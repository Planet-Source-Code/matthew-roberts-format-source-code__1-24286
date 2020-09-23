Attribute VB_Name = "CodeFormat"
Option Explicit
'===============================================================================
'Name:          FormatCode
'Purpose:       Completely formats a given block of VB code to the desired blocking
'               Completely formats a given block of VB code to the desired blocking and whitespacing rules.
'Returns:       String - Formatted source code.
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:  Set the formatting by typing "config" in the editor window, highlighting
'               it, and right-clicking on Format Code. A configuration form will
'               open. Data is stored in a file called "commands.mdb" which should
'               Set the formatting by typing "config" in the editor window, highlighting it, and right-clicking on Format Code. A configuration form will open. Data is stored in a file called "commands.mdb" which should have a reference in the registrty.
'===============================================================================

Public Function FormatCode(strOriginal As String) As String

Dim lngStartPos         As Long
Dim strFirstword        As String
Dim lngEndpos           As Long
Dim lngCurrentpos       As Long
Dim strFormatted        As String
Dim lngCurrentLine      As Long
Dim strCurrentLine      As String
Dim strCommands()       As String
Dim intCmds             As Long
Dim intAddTabs          As Integer
Dim strConverted        As String
Dim blnLastLine         As Boolean
Dim strCommand          As String
Dim intCommandCheck     As Integer
Dim blnCommandFound     As Boolean
Dim lngWhiteSpaceBefore As Long
Dim lngWhiteSpaceAfter  As Long
Dim intBlockBefore      As Integer
Dim intBlockAfter       As Integer
Dim blnBlockBack        As Boolean
Dim strBlock            As String
Dim strWSBefore         As String
Dim strWSAfter          As String
Dim blnSingleRevIndent  As Boolean
Dim intFindKeyword      As Integer
Dim lngTabCount         As Long
Dim intAddTab           As Long
Dim strNewline          As String
Dim blnFirstLineDone    As Boolean
Dim lngStartTest         As Long
Dim strContinuedline    As String
Dim lngLastPos          As Long
Dim blnWhitespaceAdded As Boolean
Dim lngLineStart As Long
Dim lngLineEnd As Long
Dim lngNextCR As Long
Dim lngNextContinue As Long
Dim blnLineFinished As Boolean

lngStartPos = 1
Screen.MousePointer = vbHourglass
ReDim intCommandQueue(0)
'   First, remove all current blocking and spacing.
strFormatted = Replace(strOriginal, vbTab, "")

While InStr(1, strFormatted, vbCrLf & vbCrLf) > 0
    strFormatted = Replace(strFormatted, vbCrLf & vbCrLf, vbCrLf)
Wend

ReDim strCommands(gintCommandCount)

For intCmds = 0 To gintCommandCount - 1
    strCommands(intCmds) = Cmds(intCmds + 1).CommandText
Next intCmds

'   Get the first line
lngStartPos = 1
lngCurrentpos = 1

lngCurrentpos = InStr(lngStartPos, strFormatted, vbCrLf)


     Do While lngCurrentpos > 0
     
    If lngCurrentpos > 1 Then
    lngStartTest = lngCurrentpos
    If lngCurrentpos < lngLastPos Then
        Exit Do
    End If

    lngLastPos = lngCurrentpos
    strCurrentLine = Trim(Mid(strFormatted, lngStartPos, lngCurrentpos - lngStartPos))
    lngEndpos = lngCurrentpos
        
        If Right(strCurrentLine, 2) = " _" Then
            'Line continued...get the rest of it
        lngLineStart = lngCurrentpos
        strContinuedline = strCurrentLine
       Do
          '  lngLineEnd = InStr(lngLineStart + 2, strFormatted, " _")
            lngNextContinue = InStr(lngLineStart + 2, strFormatted, " _")
            lngNextCR = InStr(lngLineStart + 2, strFormatted, vbCr)
            If lngNextCR = 0 Then lngNextCR = Len(strFormatted)
            If lngNextContinue = 0 Then lngNextContinue = Len(strFormatted)
            
            If lngNextContinue < lngNextCR Then
                 lngLineEnd = lngNextContinue
            Else
                lngLineEnd = lngNextCR - 1
                blnLineFinished = True
            End If
            strContinuedline = strContinuedline & vbCrLf & BuildString(vbTab, lngTabCount) & Mid(strFormatted, lngLineStart + 2, lngLineEnd - lngLineStart)
            strContinuedline = TrimChar(strContinuedline, vbCrLf)
            lngStartTest = lngLineStart
            lngLineStart = lngLineEnd + 2
            
        Loop While Not blnLineFinished
        blnLineFinished = False
        lngEndpos = lngLineEnd

            strCurrentLine = strContinuedline
            Debug.Print strCurrentLine
            
        End If

End If
        
        lngCurrentpos = lngEndpos
        lngStartTest = 1
 '   lngEndPos = InStr(1, strCurrentLine, " ")
  '  If lngEndPos = 0 Then lngEndPos = Len(strCurrentLine)
    blnCommandFound = False
    blnWhitespaceAdded = False
If Trim(Left(strCurrentLine, 1)) <> "'" Then

    For intCommandCheck = 0 To gintCommandCount - 1

        If Cmds(intCommandCheck).CommandText > "" Then

            If Left(strCurrentLine, Len(Cmds(intCommandCheck).CommandText)) = Cmds(intCommandCheck).CommandText Then
                If Mid(strCurrentLine, Len(Cmds(intCommandCheck).CommandText) + 1, 1) = " " Or Len(strCurrentLine) = Len(Cmds(intCommandCheck).CommandText) Then
                    strCommand = Cmds(intCommandCheck).CommandText
                    blnCommandFound = True
                End If
                Exit For
             End If

        End If

    Next intCommandCheck


    If blnCommandFound Then
        
            For intFindKeyword = 0 To gintCommandCount - 1
    
                If Cmds(intFindKeyword).CommandText = strCommand Then
                    '   If is special...it can terminate on one line.
                    '   Make sure this is not a single-line if...then
    
                    If strCommand = "If" Then
    
                        If Right(Trim(strCurrentLine), 4) <> "Then" Then
                                Exit For
                        End If
    
                    End If
    
                    '   Found the command's UDT...now find check rules.
                    lngWhiteSpaceBefore = Cmds(intFindKeyword).WhitespaceBefore
                    lngWhiteSpaceAfter = Cmds(intFindKeyword).WhitespaceAfter
                    intBlockBefore = Cmds(intFindKeyword).BlockBefore
                    intBlockAfter = Cmds(intFindKeyword).BlockAfter
    
                    If intBlockBefore Then
                        lngTabCount = lngTabCount + Cmds(intFindKeyword).BlockBefore
                    ElseIf Cmds(intFindKeyword).BlockBefore = -99 Then
                        lngTabCount = 0
                    ElseIf Cmds(intFindKeyword).BlockBefore = 99 Then
                        lngTabCount = 1
                    ElseIf intBlockAfter Then
                        intAddTab = Cmds(intFindKeyword).BlockAfter
                    ElseIf Cmds(intFindKeyword).BlockBefore = -1 Then
                        lngTabCount = lngTabCount - 1
                        '                 blnCommandFound = True
                    End If
    
                    intAddTab = Cmds(intFindKeyword).BlockAfter
                    Exit For
                End If
    
            Next intFindKeyword
        End If
    End If


    If blnSingleRevIndent Then
        lngTabCount = lngTabCount - 1
    End If

    If lngTabCount < 0 Then lngTabCount = 0
    strBlock = BuildString(vbTab, lngTabCount)
    strWSBefore = ""
    strWSAfter = ""

    If lngWhiteSpaceBefore Then
        If Not blnWhitespaceAdded Then
            strWSBefore = BuildString(vbCrLf, lngWhiteSpaceBefore)
            lngWhiteSpaceBefore = 0
            blnWhitespaceAdded = False
        End If
    End If


    If lngWhiteSpaceAfter Then
        strWSAfter = BuildString(vbCrLf, lngWhiteSpaceAfter)
        lngWhiteSpaceAfter = 0
        blnWhitespaceAdded = True
    End If

    strConverted = strConverted & strWSBefore & strBlock & strCurrentLine & vbCrLf & strWSAfter

    If blnSingleRevIndent Then
        lngTabCount = lngTabCount + 1
        blnSingleRevIndent = False
    End If

    intBlockBefore = 0
    intBlockAfter = 0
    lngWhiteSpaceAfter = 0
    lngWhiteSpaceAfter = 0
    strBlock = ""
    strWSAfter = ""
    strWSBefore = ""
    lngStartPos = lngCurrentpos + 2

    If Not blnLastLine Then
        lngCurrentpos = InStr(lngStartPos, strFormatted, vbCrLf)

        If lngCurrentpos = 0 Then
            blnLastLine = True
            lngCurrentpos = Len(strFormatted) + 2
        End If

    Else
        Exit Do
    End If

    If lngTabCount < 0 Then lngTabCount = 0

    If intAddTab > 0 Then
        lngTabCount = lngTabCount + intAddTab
        intAddTab = 0
    End If

Loop


FormatCode = strConverted
EXIT_FormatCode:
    Screen.MousePointer = vbNormal

End Function




Public Function CreateContinueLines(ConvertString As String, ColCount As Long) As String
Dim strCurrentLine As String
Dim lngCurrentpos       As Long
Dim strContinuedline        As String
Dim lngEndpos       As Long
Dim strNewline      As String
Dim lngStartTest As Long

        
        If Right(ConvertString, 1) = "_" Then
            'Line continued...get the rest of it
            
            strContinuedline = strCurrentLine & vbCrLf & BuildString(vbTab, 5)
            '   As long as we hit line continuations before ")", we are still on the continued block
            lngStartTest = lngCurrentpos
            
            Do
            
                lngEndpos = InStr(lngStartTest + 1, ConvertString, "_")
                If Mid(ConvertString, lngStartTest, 2) = vbCrLf Then
                    lngStartTest = lngStartTest + 1
                End If
                If lngEndpos = 0 Then
                    Exit Do
                End If
                strNewline = Mid(ConvertString, lngStartTest + 1, lngEndpos - lngStartTest)
                strContinuedline = strContinuedline & strNewline & vbCrLf & BuildString(vbTab, 5)
                lngStartTest = lngEndpos + 2
    
            Loop While InStr(lngStartTest, ConvertString, ")") >= InStr(lngStartTest, ConvertString, "_")
            
            lngEndpos = InStr(lngStartTest, ConvertString, ")")
            strNewline = Mid(ConvertString, lngStartTest + 1, lngEndpos - lngStartTest)
            strContinuedline = strContinuedline & strNewline & vbCrLf & BuildString(vbTab, 5)
            strCurrentLine = strContinuedline

        End If




End Function
