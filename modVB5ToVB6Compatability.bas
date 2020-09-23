Attribute VB_Name = "modVB5ToVB6Compatability"
Option Explicit

'       UNCOMMENT THESE FUNCTION IF YOU ARE RUNNING VB5!!!


'===============================================================================
'Name:
'Purpose:       This module contains function that are present in VB6 but not
'               having to code the functions.
'Returns:
'Created By:    Matthew M. Roberts (M@)
'Date:          6/1/2001
'Comments:  Just add this to any VB5 project.               '
'===============================================================================
'===============================================================================
'Name:          Replace
'Purpose:       Provides VB6 compatability to VB5 for the Replace function.
'Returns:
'Created By:    Matthew M. Roberts (M@)
'Date:          6/1/2001
'Comments:
'===============================================================================

'Public Function Replace(strRawString As String, _
'             strFind As String, _
'             strReplaceWith As String) As String
'
'
'
'    Dim lngCurrentChar As Long
'    Dim strReplaced As String
'
'
'    For lngCurrentChar = 1 To Len(strRawString)
'
'        If Mid(strRawString, lngCurrentChar, Len(strFind)) = strFind Then
'            strReplaced = strReplaced & strReplaceWith
'            lngCurrentChar = lngCurrentChar + Len(strFind) - 1
'        Else
'            strReplaced = strReplaced & Mid(strRawString, lngCurrentChar, 1)
'        End If
'
'
'
'    Next lngCurrentChar
'
'
'    Replace = strReplaced
'
'End Function
'
''===============================================================================
''Name:          InstrRev
''Purpose:       VB5 Equivilent of VB6's InstrRev() function.
''Returns:       Long - Character position.
''Created By:    Matthew M. Roberts (M@)
''Date:          6/20/2001
''Comments:
''===============================================================================
'
'Public Function InstrRev(StartPos As Long, sValue As String, Find As String) As Long
'
'    Dim lngChar As Long
'
'    For lngChar = StartPos To 1 Step -1
'
'        '    Debug.Print Mid(sValue, lngChar, Len(Find))
'
'
'        If Mid(sValue, lngChar, Len(Find)) = Find Then
'            InstrRev = lngChar
'            Exit For
'        End If
'
'    Next lngChar
'
'
'
'End Function
'
'
''
