VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPQDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================================
' CLASS: clsPQDatabase
' PURPOSE: Framework for managing state and variables in Power Query through M code.
' =========================================================================================

' =========================================================================================
' 1. DECLARATIONS AND STATE VARIABLES
' =========================================================================================
Private p_QueryName As String
Private p_wb As Workbook

Public ViewAsRecord As Boolean ' When True, appends the [Key=Key] record block



' =========================================================================================
' 2. CONSTRUCTOR
' =========================================================================================
' ==========================================================================
' Routine:      Class_Initialize
' Description:   Default class constructor. Initializes state by defining the
'              working environment and default query settings.
' Parameters:  None
' ==========================================================================
Private Sub Class_Initialize()
    ' Uses the active workbook as the query source
    Set p_wb = ActiveWorkbook
    p_QueryName = "Variables" ' Sets the factory default query name
    ' Ensures the final M output is structured as a record
    ViewAsRecord = True
End Sub



' =========================================================================================
' 3. PUBLIC PROPERTIES
' =========================================================================================
' ==========================================================================
' Property: QueryName (Get)
' Description:   Returns the target Power Query query name.
' Returns:     String - Current query name.
' ==========================================================================
Public Property Get queryName() As String
    queryName = p_QueryName
End Property

' ==========================================================================
' Property: QueryName (Let)
' Description:   Allows the target query name to be changed dynamically.
' Parameters:  vNewValue (String) - New name to assign.
' ==========================================================================
Public Property Let queryName(ByVal vNewValue As String)
    p_QueryName = vNewValue
End Property

' ==========================================================================
' Property: Exists (Get)
' Description:   Checks whether the query defined by p_QueryName exists in Excel.
' Returns:     Boolean - True when the query is found; otherwise False.
' ==========================================================================
Public Property Get Exists() As Boolean
    Dim query As WorkbookQuery
    ' Temporarily suppresses errors while attempting to obtain the query object
    On Error Resume Next
    Set query = p_wb.Queries(p_QueryName)
    On Error GoTo 0
    ' A non-empty object reference means that the query was found
    Exists = Not (query Is Nothing)
End Property

' ==========================================================================
' Property: CurrentMCode (Get)
' Description:   Returns the current M formula for the query when it exists.
' Returns:     String - Query M code, or an empty string.
' ==========================================================================
Public Property Get CurrentMCode() As String
    If Me.Exists Then
        CurrentMCode = p_wb.Queries(p_QueryName).Formula
    Else
        CurrentMCode = ""
    End If
End Property



' =========================================================================================
' 4. INFRASTRUCTURE AND STATE
' =========================================================================================
' ==========================================================================
' Routine:      Initialize
' Description:   Creates the base Power Query query when it does not exist,
'              providing a minimal valid M scaffold.
' Parameters:  None
' ==========================================================================
Public Sub Initialize()
    Dim baseMCode As String
    If Me.Exists Then Exit Sub
    
    ' Builds the minimal let...in structure required by Power Query
    baseMCode = "let" & vbCrLf & _
                 "    _Iniciado = true" & vbCrLf & _
                 "in" & vbCrLf & _
                 "    _Iniciado"
                 
    p_wb.Queries.Add Name:=p_QueryName, Formula:=baseMCode
End Sub

' ==========================================================================
' Routine:      ClearAll
' Description:   Resets the existing query by replacing all M code
'              with the initial base scaffold.
' Parameters:  None
' ==========================================================================
Public Sub ClearAll()
    Dim mCode As String
    If Not Me.Exists Then Exit Sub
    
    ' Defines the initial M code state
    mCode = "let" & vbCrLf & _
           "    _Iniciado = true" & vbCrLf & _
           "in" & vbCrLf & _
           "    _Iniciado"
           
    ' Applies record formatting when enabled and updates the Excel query
    UpdateInBlock mCode
    p_wb.Queries(p_QueryName).Formula = mCode
    Debug.Print "Query [" & p_QueryName & "] was successfully reset."
End Sub

' ==========================================================================
' Function:      VariableExists
' Description:   Uses a regular expression to check whether a specific variable
'              is declared in the query M code.
' Parameters:  Key (String) - Variable name to locate.
' Returns:     Boolean - True when the variable exists.
' ==========================================================================
Public Function VariableExists(ByVal Key As String) As Boolean
    Dim mCode As String, regex As Object
    If Not Me.Exists Then
        VariableExists = False
        Exit Function
    End If
    
    mCode = Me.CurrentMCode
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configures the regex to match Key as an exact identifier at the start of a line
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "^[ \t]*\b" & Key & "\b\s*="
    End With
    
    VariableExists = regex.Test(mCode)
    Set regex = Nothing
End Function



' =========================================================================================
' 5. CRUD: READ AND EXTRACTION
' =========================================================================================
' ==========================================================================
' Function:      ReadValue
' Description:   Extracts and interprets the value of a specific variable from
'              M code and converts it to an appropriate VBA data type.
' Parameters:  Key (String) - Variable to read.
' Returns:     Variant - Parsed typed value, or Empty when not found.
' ==========================================================================
Public Function ReadValue(ByVal Key As String) As Variant
    Dim mCode As String, regex As Object, matches As Object
    mCode = Me.CurrentMCode
    If mCode = "" Then Exit Function
    
    Set regex = CreateObject("VBScript.RegExp")
    ' Expression that captures everything after the equals sign on the same line
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "^[ \t]*\b" & Key & "\b\s*=\s*([^\r\n]*)"
    End With
    
    If regex.Test(mCode) Then
        Set matches = regex.Execute(mCode)
        ' Passes the captured content to the M-to-VBA conversion engine
        ReadValue = ParseMValue(Trim(matches(0).SubMatches(0)))
    Else
        ReadValue = Empty
    End If
    Set regex = Nothing
End Function

' ==========================================================================
' Function:      ListAll
' Description:   Maps and returns all variables and values contained in
'              the M query as a VBA dictionary.
' Parameters:  DebugMode (Boolean) - When True, prints results to the Immediate window.
' Returns:     Object (Scripting.Dictionary) - Dictionary containing typed keys and values.
' ==========================================================================
Public Function ListAll(Optional ByVal DebugMode As Boolean = False) As Object
    Dim mCode As String, letBlock As String
    Dim regex As Object, matches As Object, match As Object
    Dim resultDictionary As Object
    Dim Key As String, Value As Variant
    
    Set resultDictionary = CreateObject("Scripting.Dictionary")
    mCode = Me.CurrentMCode
    If mCode = "" Then GoTo Done
    
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Isolates the "let" block by removing the entire "in" block before extraction
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "\bin\b[\s\S]*"
        letBlock = .Replace(mCode, "")
        
        .Global = True
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^\r\n]*)"
    End With
    
    If regex.Test(letBlock) Then
        Set matches = regex.Execute(letBlock)
        If DebugMode Then Debug.Print "--- POWER QUERY VARIABLES [" & p_QueryName & "] ---"
        
        ' Iterates through all matches and converts M types to VBA
        For Each match In matches
            Key = Trim(match.SubMatches(0))
            Value = ParseMValue(Trim(match.SubMatches(1)))
            resultDictionary.Add Key, Value
            If DebugMode Then Debug.Print "- " & Key & " (" & TypeName(Value) & ") = " & CStr(Value)
        Next match
        
        If DebugMode Then Debug.Print "---------------------------------------------------"
    End If
Done:
    Set ListAll = resultDictionary
    Set regex = Nothing
End Function



' =========================================================================================
' 6. CRUD: WRITE AND DELETE
' =========================================================================================
' ==========================================================================
' Routine:      SetValue
' Description:   Updates or inserts a single variable and its value in M code.
' Parameters:  Key (String) - Variable name.
'              Value (Variant) - Value to insert.
' ==========================================================================
Public Sub SetValue(ByVal Key As String, ByVal Value As Variant)
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.Add Key, Value
    ' Delegates to the batch engine (SetMultiple) to avoid duplicated logic
    Me.SetMultiple dictionary
End Sub

' ==========================================================================
' Routine:      SetMultiple
' Description:   Inserts or updates a batch of variables supplied through
'              a Dictionary object directly in M code.
' Parameters:  VariablesDict (Object) - Dictionary containing keys and values.
' ==========================================================================
Public Sub SetMultiple(ByVal VariablesDict As Object)
    Dim mCode As String, formattedValue As String, newLine As String
    Dim regex As Object, Key As Variant
    
    If VariablesDict Is Nothing Then Exit Sub
    ' Ensures the base script exists before manipulating strings
    If Not Me.Exists Then Me.Initialize
    mCode = Me.CurrentMCode
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
    End With
    
    For Each Key In VariablesDict.Keys
        ' Converts the native VBA type to its corresponding M representation
        formattedValue = FormatMValue(VariablesDict(Key))
        regex.Pattern = "^([ \t]*\b" & CStr(Key) & "\b\s*=)[^\r\n]*"
        
        ' Replaces the value on the same line when the variable already exists
        If regex.Test(mCode) Then
            mCode = regex.Replace(mCode, "$1 " & formattedValue & ",")
        Else
            ' Otherwise, inserts a new declaration line immediately after 'let'
            newLine = "let" & vbCrLf & "    " & CStr(Key) & " = " & formattedValue & ","
            regex.Pattern = "^let\b"
            mCode = regex.Replace(mCode, newLine)
        End If
    Next Key
    
    ' Removes the trailing comma from the final variable before 'in'
    regex.Pattern = ",(\s*\bin\b)"
    mCode = regex.Replace(mCode, "$1")
    
    ' Rebuilds the 'in' structure when required and saves the query
    UpdateInBlock mCode
    p_wb.Queries(p_QueryName).Formula = mCode
    Set regex = Nothing
End Sub

' ==========================================================================
' Routine:      RemoveVariable
' Description:   Removes a variable declaration from M code when it exists.
' Parameters:  Key (String) - Variable name to remove.
' ==========================================================================
Public Sub RemoveVariable(ByVal Key As String)
    Dim mCode As String, regex As Object
    If Not Me.Exists Then Exit Sub
    mCode = Me.CurrentMCode
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "^[ \t]*\b" & Key & "\b\s*=[^\r\n]*\r?\n?"
    End With
    
    If regex.Test(mCode) Then
        ' Removes the entire line for the matched variable
        mCode = regex.Replace(mCode, "")
        
        ' Keeps the syntax valid by removing the comma before 'in'
        regex.Pattern = ",(\s*\bin\b)"
        mCode = regex.Replace(mCode, "$1")
        
        UpdateInBlock mCode
        p_wb.Queries(p_QueryName).Formula = mCode
    End If
    Set regex = Nothing
End Sub



' =========================================================================================
' 7. UTILITIES (Snapshots and Environments)
' =========================================================================================
' ==========================================================================
' Function:      Clone
' Description:   Creates an exact copy of an existing Power Query query.
' Parameters:  SourceQuery (String) - Name of the query to copy.
'              NewQuery (String) - Name of the cloned query.
'              Overwrite (Boolean) - When True, replaces an existing target query.
' Returns:     Boolean - True when cloning succeeds.
' ==========================================================================
Public Function Clone(ByVal SourceQuery As String, ByVal NewQuery As String, Optional ByVal Overwrite As Boolean = False) As Boolean
    Dim sourceQueryObject As WorkbookQuery, targetQueryObject As WorkbookQuery
    Dim mCode As String
    
    ' Safely validates that the source query exists
    On Error Resume Next
    Set sourceQueryObject = p_wb.Queries(SourceQuery)
    On Error GoTo 0
    
    If sourceQueryObject Is Nothing Then
        Clone = False
        Exit Function
    End If
    
    mCode = sourceQueryObject.Formula
    
    On Error Resume Next
    Set targetQueryObject = p_wb.Queries(NewQuery)
    On Error GoTo 0
    
    If Not targetQueryObject Is Nothing Then
        ' Handles an existing target query to prevent Excel errors
        If Overwrite Then
            targetQueryObject.Formula = mCode
            Clone = True
        Else
            Err.Raise vbObjectError + 2, "clsPQDatabase", "The query '" & NewQuery & "' already exists. Use Overwrite:=True."
            Clone = False
        End If
    Else
        p_wb.Queries.Add Name:=NewQuery, Formula:=mCode
        Clone = True
    End If
End Function



' =========================================================================================
' 8. INTERNAL AND PRIVATE ENGINES (Do not call externally)
' =========================================================================================
' ==========================================================================
' Routine:      UpdateInBlock (Private)
' Description:   Dynamically rebuilds the M code 'in' block to expose
'              all variables as a record for easier debugging.
' Parameters:  mCode (String, ByRef) - M code being modified in memory.
' ==========================================================================
Private Sub UpdateInBlock(ByRef mCode As String)
    If Not Me.ViewAsRecord Then Exit Sub
    
    Dim regex As Object, matches As Object, match As Object
    Dim keysBlock As String, Key As String, letBlock As String
    
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Splits the code at the 'in' keyword
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "\bin\b[\s\S]*"
    End With
    letBlock = regex.Replace(mCode, "")
    
    ' Removes trailing whitespace from the let block
    regex.Multiline = False
    regex.Pattern = "\s+$"
    letBlock = regex.Replace(letBlock, "")
    
    ' Maps all variables declared in the let block
    With regex
        .Global = True: .Multiline = True
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)[ \t]*="
    End With
    
    If regex.Test(letBlock) Then
        Set matches = regex.Execute(letBlock)
        keysBlock = "    [" & vbCrLf
        
        ' Builds the M record syntax: [var1 = var1, var2 = var2]
        For Each match In matches
            Key = match.SubMatches(0)
            keysBlock = keysBlock & "        " & Key & " = " & Key & "," & vbCrLf
        Next match
        
        ' Removes the extra comma from the final record property
        If Len(keysBlock) > 3 Then keysBlock = Left(keysBlock, Len(keysBlock) - 3) & vbCrLf
        keysBlock = keysBlock & "    ]"
        mCode = letBlock & vbCrLf & "in" & vbCrLf & keysBlock
    End If
    Set regex = Nothing
End Sub

' ==========================================================================
' Function:      ParseMValue (Private)
' Description:   Reads a textual M value representation and converts it back
'              to its equivalent original VBA data type.
' Parameters:  rawValue (String) - M string fragment containing the value.
' Returns:     Variant - Processed and typed value for VBA use.
' ==========================================================================
Private Function ParseMValue(ByVal rawValue As String) As Variant
    Dim innerValue As String, dateParts() As String
    ' Removes a syntactic comma when included in the captured string
    If Right(rawValue, 1) = "," Then rawValue = Trim(Left(rawValue, Len(rawValue) - 1))
    
    ' Conditional parser based on the M string prefix and format
    If Left(rawValue, 10) = "#datetime(" Then
        innerValue = Mid(rawValue, 11, Len(rawValue) - 11)
        dateParts = Split(innerValue, ",")
        ParseMValue = DateSerial(Trim(dateParts(0)), Trim(dateParts(1)), Trim(dateParts(2))) + _
                      TimeSerial(Trim(dateParts(3)), Trim(dateParts(4)), Trim(dateParts(5)))
    ElseIf Left(rawValue, 6) = "#date(" Then
        innerValue = Mid(rawValue, 7, Len(rawValue) - 7)
        dateParts = Split(innerValue, ",")
        ParseMValue = DateSerial(Trim(dateParts(0)), Trim(dateParts(1)), Trim(dateParts(2)))
    ElseIf Left(rawValue, 1) = """" And Right(rawValue, 1) = """" Then
        ParseMValue = Mid(rawValue, 2, Len(rawValue) - 2)
    ElseIf LCase(rawValue) = "true" Then ParseMValue = True
    ElseIf LCase(rawValue) = "false" Then ParseMValue = False
    ElseIf IsNumeric(rawValue) Then
        ' Ensures the decimal separator is understood by the regional VBA settings
        ParseMValue = CDbl(Replace(rawValue, ".", ","))
    Else
        ParseMValue = rawValue
    End If
End Function

' ==========================================================================
' Function:      FormatMValue (Private)
' Description:   Analyzes a VBA value and creates the M string literal
'              required to inject it into Power Query.
' Parameters:  Value (Variant) - VBA value to transform.
' Returns:     String - M code fragment representing the original value.
' ==========================================================================
Private Function FormatMValue(ByVal Value As Variant) As String
    ' Handles complex objects by extracting their underlying values (for example, ranges)
    If IsObject(Value) Then
        If TypeName(Value) = "Range" Then
            FormatMValue = FormatMValue(Value.Value2)
            Exit Function
        Else
            FormatMValue = "null"
            Exit Function
        End If
    End If

    ' Converts arrays into lists of M values (for example, {1, 2, "a"})
    If IsArray(Value) Then
        Dim item As Variant, mList As String
        mList = "{"
        For Each item In Value
            If Not IsEmpty(item) Then mList = mList & FormatMValue(item) & ", "
        Next item
        If Len(mList) > 2 Then mList = Left(mList, Len(mList) - 2)
        FormatMValue = mList & "}"
        Exit Function
    End If

    ' Converts scalar values by inspecting VarType and generating strict M syntax
    Select Case VarType(Value)
        Case vbString: FormatMValue = """" & Value & """"
        Case vbBoolean: FormatMValue = IIf(Value, "true", "false")
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency: FormatMValue = Replace(CStr(Value), ",", ".")
        Case vbDate: FormatMValue = "#datetime(" & Year(Value) & ", " & Month(Value) & ", " & Day(Value) & ", " & Hour(Value) & ", " & Minute(Value) & ", " & Second(Value) & ")"
        Case vbNull, vbEmpty: FormatMValue = "null"
        Case Else: FormatMValue = """" & CStr(Value) & """"
    End Select
End Function


' =========================================================================================
' 9. RAD TOOLS (Component Import and Export)
' =========================================================================================
' ==========================================================================
' Routine:      Tool_ExportQueryToTXT
' Description:   Extracts metadata (description and instructions) and M code
'              from the query into a tagged text file.
' Parameters:  QueryName (String) - Query to export.
'              FilePath (String) - Physical destination of the .txt file.
' ==========================================================================
Public Sub Tool_ExportQueryToTXT(ByVal queryName As String, ByVal filePath As String)
    Dim query As WorkbookQuery
    Dim stream As Object
    Dim fullDescription As String, descriptionPart As String, usagePart As String
    Dim usageDelimiter As String
    Dim TextContent As String
    
    ' Checks whether the query exists
    On Error Resume Next
    Set query = p_wb.Queries(queryName)
    On Error GoTo 0
    
    If query Is Nothing Then
        Err.Raise vbObjectError + 3, "clsPQDatabase", "The query '" & queryName & "' does not exist."
        Exit Sub
    End If
    
    ' Splits the description and usage instructions using a marker in the Power Query properties
    fullDescription = query.Description
    usageDelimiter = "[COMO USAR]"
    
    If InStr(1, fullDescription, usageDelimiter, vbTextCompare) > 0 Then
        descriptionPart = Trim(Split(fullDescription, usageDelimiter)(0))
        usagePart = Trim(Split(fullDescription, usageDelimiter)(1))
    Else
        descriptionPart = fullDescription
        usagePart = ""
    End If
    
    ' Builds the structured text using custom protection tags
    TextContent = "<NOME>" & vbCrLf & query.Name & vbCrLf & "</NOME>" & vbCrLf & _
                  "<DESCRICAO>" & vbCrLf & descriptionPart & vbCrLf & "</DESCRICAO>" & vbCrLf & _
                  "<COMOUSAR>" & vbCrLf & usagePart & vbCrLf & "</COMOUSAR>" & vbCrLf & _
                  "<CODIGOM>" & vbCrLf & query.Formula & vbCrLf & "</CODIGOM>"
                  
    ' Saves the file as UTF-8 through ADODB.Stream to preserve Unicode characters
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText TextContent
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub

' ==========================================================================
' Routine:      Tool_ImportQueryFromTXT
' Description:   Reads a structured .txt file through RAD Tools and rebuilds
'              the Power Query query by injecting its M code and metadata.
' Parameters:  FilePath (String) - Logical path to the text file.
'              Overwrite (Boolean) - Permission to replace an existing query.
' ==========================================================================
Public Sub Tool_ImportQueryFromTXT(ByVal filePath As String, Optional ByVal Overwrite As Boolean = False)
    Dim stream As Object
    Dim TextContent As String
    Dim importedName As String, importedDescription As String, importedUsage As String, importedMCode As String
    Dim finalDescription As String
    Dim query As WorkbookQuery
    
    ' Reads the UTF-8 file to preserve metadata encoding and Unicode characters
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    TextContent = stream.ReadText
    stream.Close
    
    ' Extracts tags through the private helper function
    importedName = ExtractTag(TextContent, "NOME")
    importedDescription = ExtractTag(TextContent, "DESCRICAO")
    importedUsage = ExtractTag(TextContent, "COMOUSAR")
    importedMCode = ExtractTag(TextContent, "CODIGOM")
    
    ' Validates the imported package
    If importedName = "" Or importedMCode = "" Then
        Err.Raise vbObjectError + 4, "clsPQDatabase", "The file is corrupted or does not follow the RAD format."
        Exit Sub
    End If
    
    ' Combines description and usage text for the native Power Query description property
    finalDescription = importedDescription
    If importedUsage <> "" Then
        finalDescription = finalDescription & vbCrLf & vbCrLf & "[COMO USAR]" & vbCrLf & importedUsage
    End If
    
    ' Creates or updates the query in Excel
    On Error Resume Next
    Set query = p_wb.Queries(importedName)
    On Error GoTo 0
    
    If Not query Is Nothing Then
        If Overwrite Then
            query.Formula = importedMCode
            query.Description = finalDescription
        Else
            Err.Raise vbObjectError + 5, "clsPQDatabase", "The component '" & importedName & "' already exists. Use Overwrite:=True."
        End If
    Else
        ' Adds the name, M code, and metadata to a new query
        Set query = p_wb.Queries.Add(Name:=importedName, Formula:=importedMCode, Description:=finalDescription)
    End If
End Sub

' ==========================================================================
' Routine:      DB_SyncDevFile
' Description:   Synchronizes a UTF-8 file with a Power Query query
'              only when the local content has changed.
' Parameters:  FilePath (String) - Physical path to the local file.
'              QueryName (String) - Name of the query to synchronize.
' ==========================================================================
Public Sub DB_SyncDevFile(ByVal filePath As String, ByVal queryName As String)
    Dim stream As Object
    Dim localText As String
    Dim storedText As String

    On Error GoTo SilentFailure
    If Dir$(filePath) = "" Then Exit Sub

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    localText = stream.ReadText
    stream.Close
    Set stream = Nothing

    Me.queryName = queryName
    storedText = Me.DB_ReadRawText()

    If localText <> storedText Then
        Me.DB_SaveRawText TextContent:=localText
    End If

    Exit Sub

SilentFailure:
    On Error Resume Next
    If Not stream Is Nothing Then
        If stream.State <> 0 Then stream.Close
    End If
    Set stream = Nothing
    On Error GoTo 0
End Sub

' ==========================================================================
' Routine:      DB_SaveRawText
' Description:   Stores raw text (HTML, CSS, or JavaScript) directly in the current query.
'              Accepts code through an argument or reads it from a physical file.
' Parameters:  TextContent (String, Optional) - String content to save.
'              FilePath (String, Optional) - Physical file path.
' ==========================================================================
Public Sub DB_SaveRawText(Optional ByVal TextContent As String = "", Optional ByVal filePath As String = "")
    Dim mCode As String
    Dim query As WorkbookQuery
    Dim finalContent As String
    Dim stream As Object
    
    ' 1. Validates that the query name was set before saving
    If p_QueryName = "" Then
        Err.Raise vbObjectError + 8, "clsPQDatabase", "Set QueryName before saving."
        Exit Sub
    End If
    
    ' 2. Selects the text source: physical file or direct string
    If filePath <> "" Then
        ' Uses ADODB.Stream to ensure correct UTF-8 decoding
        On Error Resume Next
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' adTypeText
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile filePath
        
        If Err.Number <> 0 Then
            Err.Raise vbObjectError + 6, "clsPQDatabase", "Could not read the file: " & filePath
            Exit Sub
        End If
        On Error GoTo 0
        
        finalContent = stream.ReadText
        stream.Close
    ElseIf TextContent <> "" Then
        finalContent = TextContent
    Else
        Err.Raise vbObjectError + 7, "clsPQDatabase", "Provide TextContent or FilePath to save."
        Exit Sub
    End If
    
    ' 3. Formats the string for the strict Power Query representation
    mCode = """" & Replace(finalContent, """", """""") & """"
    
    ' 4. Checks whether the query must be created or updated
    On Error Resume Next
    Set query = p_wb.Queries(p_QueryName)
    On Error GoTo 0
    
    If Not query Is Nothing Then
        query.Formula = mCode
    Else
        p_wb.Queries.Add Name:=p_QueryName, Formula:=mCode
    End If
End Sub

' ==========================================================================
' Function:      DB_ReadRawText
' Description:   Extracts raw text stored in the current query by removing
'              the M string wrapper and escaped double quotes.
' Parameters:  None.
' Returns:     String - Original raw content (HTML, CSS, JavaScript, and so on).
' ==========================================================================
Public Function DB_ReadRawText() As String
    Dim mCode As String
    Dim query As WorkbookQuery
    
    On Error Resume Next
    Set query = p_wb.Queries(p_QueryName)
    On Error GoTo 0
    
    If Not query Is Nothing Then
        mCode = query.Formula
        
        ' Checks whether the text is wrapped in M string quotes
        If Left(mCode, 1) = """" And Right(mCode, 1) = """" Then
            ' Removes the first and last quote
            mCode = Mid(mCode, 2, Len(mCode) - 2)
        End If
        
        ' Restores double quotes escaped for Power Query
        DB_ReadRawText = Replace(mCode, """""", """")
    Else
        ' Returns an empty string when the query is not found
        DB_ReadRawText = ""
    End If
End Function

' ==========================================================================
' Function:      ExtractTag (Private)
' Description:   Parses the XML-like syntax generated by the class to
'              extract content enclosed by two specific tags.
' Parameters:  Text (String) - Raw file content.
'              TagName (String) - Base opening and closing tag name.
' Returns:     String - Text contained inside the tag.
' ==========================================================================
Private Function ExtractTag(ByVal Text As String, ByVal TagName As String) As String
    Dim regex As Object, matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Global = False
    regex.IgnoreCase = True
    regex.Multiline = True
    ' Regex that reads exactly what appears between <TAG> and </TAG>, including line breaks
    regex.Pattern = "<" & TagName & ">\r?\n?([\s\S]*?)\r?\n?</" & TagName & ">"
    
    If regex.Test(Text) Then
        Set matches = regex.Execute(Text)
        ExtractTag = Trim(matches(0).SubMatches(0))
    Else
        ExtractTag = ""
    End If
    
    Set regex = Nothing
End Function



