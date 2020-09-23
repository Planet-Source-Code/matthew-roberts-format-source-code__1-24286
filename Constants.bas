Attribute VB_Name = "Constants"
Option Explicit


'Public Cmds() As CommandInfo


Public Enum Commands
    vbcmWhile = 1
    vbcmDo = 2
    vbcmif = 20
    vbcmEndIf = 3
    vbcmWith = 4
    vbcmSelect = 5
    vbcmFor = 7
    vbcmSub = 8
    vbcmFunction = 9
    vbcmPrivatesub = 10
    vbcmPrivateFunction = 11
    vbcmPublicFunction = 12
    vbcmPublicSub = 13
    vbcmLoop = 14
    vbcmCase = 15
    vbcmElse = 16
    vbcmEndWith = 17
    vbcmNext = 18
    vbcmWend = 19
    vbcmEndSub = 21
    vbcmEndFunction = 22
    
End Enum

Type CommandInfo
    CommandText     As String
    BlockBefore      As Boolean
    BlockAfter      As Boolean
    BlockVal        As Integer
    WhitespaceBefore      As Boolean
    WhitespaceAfter     As Boolean
End Type

'   Change this constant if any commands are added:
Public gintCommandCount As Integer



Sub Main()
'   Initalize command structures

ReDim Cmds(gintCommandCount)
PopulateCommands


frmCleanup.Show
End Sub
