Attribute VB_Name = "General"
Public Cmds() As CommandTypes
Public intCommandIDs() As Integer



'Public gintCommandCount As Integer

Public Type CommandTypes
    CommandText         As String
    BlockBefore         As Integer
    BlockAfter          As Integer
    WhitespaceBefore    As Integer
    WhitespaceAfter     As Integer
End Type

Public Type FunctionParams
    Name                As String
    Type                As String
    Optional            As Boolean


End Type

Public Const ILLEGAL_CHARS = "~`!@#$%^&*()+=-|\}{][;:'""""<>,.?/"



'===============================================================================
'Name:          PopulateCommands
'Purpose:       Load command parameters from table into UDTs
'Returns:       UDT Array
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:
'===============================================================================
    
    Public Function PopulateCommands()
    Dim rsCommands As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim intCmdCount As Integer
    Dim strDbPath As String
    
    On Error GoTo ERR_PopulateCommands
    
    strDbPath = GetSetting(App.Title, "startup", "datapath", App.Path & "\commands.mdb")
    While Dir(strDbPath) = "" Or strDbPath = ""
        strDbPath = InputBox("Enter the location of the command.mdb file:")
    
        If strDbPath = "" Then
            Err.Raise 2004, "PopulateCommands", "The source data file path entered (" & strDbPath & ") is invalid."
            Exit Function
        End If
    Wend
    
    SaveSetting App.Title, "startup", "datapath", strDbPath
    
    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command
    
    With conn
        .ConnectionString = "Provider=" & _
        "Microsoft.Jet.OLEDB." & _
        "4.0;Data Source=" & _
        strDbPath
    
    
        .Open
    End With
    
    With cmd
    .CommandText = "SELECT * FROM Commands"
    .ActiveConnection = conn
    Set rsCommands = .Execute
End With

Set conn = Nothing
Set cmd = Nothing

ReDim Preserve Cmds(1)

While Not rsCommands.EOF
    Cmds(intCmdCount).CommandText = rsCommands.Fields("CommandText") & ""
    Cmds(intCmdCount).BlockBefore = rsCommands.Fields("BlockBefore")
    Cmds(intCmdCount).BlockAfter = rsCommands.Fields("BlockAfter")
    Cmds(intCmdCount).WhitespaceBefore = rsCommands.Fields("WhitespaceBefore")
    Cmds(intCmdCount).WhitespaceAfter = rsCommands.Fields("WhitespaceAfter")
            
    intCmdCount = intCmdCount + 1
    ReDim Preserve Cmds(intCmdCount)
    rsCommands.MoveNext
Wend

gintCommandCount = intCmdCount

EXIT_PolulateCommands:

Exit Function

ERR_PopulateCommands:
    Resume EXIT_PolulateCommands





End Function


'===============================================================================
'Name:          MatchVal
'Purpose:       Evaluates wheter or not the MatchValue string matches any part
'               Evaluates wheter or not the MatchValue string matches any part of the ValueList string.
'Returns:       Boolean - True = Match found.
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:
'===============================================================================


Public Function MatchVal(MatchValue As String, ValueList As String)
 
    If InStr(1, ValueList, MatchValue) Then
        MatchVal = True
    End If

End Function

'===============================================================================
'Name:          BuildString
'Purpose:       Builds a string of characters (single or multiple) the length
'               Builds a string of characters (single or multiple) the length designated by RepeatCount.
'Returns:       String of repeating characters.
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:  Works like the Space() function, but for any string.
'===============================================================================

Public Function BuildString(BaseString As String, RepeatCount As Long) As String
Dim lngStringLen As Long

For lngStringLen = 1 To RepeatCount
    BuildString = BuildString & BaseString
Next lngStringLen


End Function

'===============================================================================
'Name:          PCase
'Purpose:       Returns a Proper Case string value equivilent of BaseString.Returns a Proper Case string value equivilent of BaseString.
'Returns:       String
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:  Works like UCASE or LCASE, but returns mixed case.
'===============================================================================


Public Function PCase(BaseString As String) As String

    PCase = UCase(Left(BaseString, 1)) & LCase(Right(BaseString, Len(BaseString) - 1))

End Function

Public Function RTrimChar(Source As String, RemoveChar As String) As String
Dim strConverted As String
    
    strConverted = Source
    
    While Right(strConverted, Len(RemoveChar)) = RemoveChar
        strConverted = Left(strConverted, Len(strConverted) - 1)
    Wend
    
    RTrimChar = strConverted
    
End Function

Public Function LTrimChar(Source As String, RemoveChar As String)

    strConverted = Source
    
    While Left(strConverted, Len(RemoveChar)) = RemoveChar
        strConverted = Right(strConverted, Len(strConverted) - 1)
    Wend
    
    LTrimChar = strConverted



End Function


Public Function TrimChar(Source As String, RemoveChar As String)

    strConverted = Source
    
    While Left(strConverted, Len(RemoveChar)) = RemoveChar
        strConverted = Right(strConverted, Len(strConverted) - 1)
    Wend
    
    While Left(strConverted, Len(RemoveChar)) = RemoveChar
        strConverted = Right(strConverted, Len(strConverted) - 1)
    Wend
    
    TrimChar = strConverted



End Function
