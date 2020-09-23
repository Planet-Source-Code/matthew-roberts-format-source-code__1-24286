VERSION 5.00
Begin VB.Form frmConfigFormat 
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   LinkTopic       =   "frmConfigFormat"
   ScaleHeight     =   5445
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1740
      TabIndex        =   19
      Top             =   4380
      Width           =   765
   End
   Begin VB.TextBox txtAddCommand 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   4800
      Width           =   2280
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   270
      TabIndex        =   17
      Top             =   4380
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   7650
      TabIndex        =   15
      Top             =   5040
      Width           =   885
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Settings:"
      Height          =   1575
      Left            =   2790
      TabIndex        =   5
      Top             =   3375
      Width           =   5805
      Begin VB.TextBox txtTabAfter 
         Height          =   360
         Left            =   2910
         TabIndex        =   11
         Text            =   "0"
         Top             =   555
         Width           =   375
      End
      Begin VB.TextBox txtWhitespaceafter 
         Height          =   360
         Left            =   2910
         TabIndex        =   12
         Text            =   "0"
         Top             =   1005
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   4365
         TabIndex        =   13
         Top             =   765
         Width           =   735
      End
      Begin VB.CheckBox chkForceLeft 
         Caption         =   "Force To Margin"
         Height          =   240
         Left            =   4050
         TabIndex        =   16
         Top             =   270
         Width           =   1470
      End
      Begin VB.TextBox txtWhitespaceBefore 
         Height          =   360
         Left            =   1560
         TabIndex        =   7
         Text            =   "0"
         Top             =   990
         Width           =   375
      End
      Begin VB.TextBox txtTabBefore 
         Height          =   360
         Left            =   1545
         TabIndex        =   6
         Text            =   "0"
         Top             =   540
         Width           =   375
      End
      Begin VB.Label lbltabs 
         BackStyle       =   0  'Transparent
         Caption         =   "Tabs:"
         Height          =   255
         Left            =   465
         TabIndex        =   14
         Top             =   600
         Width           =   465
      End
      Begin VB.Shape Shape2 
         Height          =   480
         Left            =   375
         Top             =   930
         Width           =   3435
      End
      Begin VB.Shape Shape4 
         Height          =   1215
         Left            =   2550
         Top             =   195
         Width           =   1260
      End
      Begin VB.Shape Shape3 
         Height          =   1215
         Left            =   1410
         Top             =   195
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   375
         Top             =   510
         Width           =   3435
      End
      Begin VB.Label lblWhiteSpace 
         Caption         =   "Whitespace:"
         Height          =   255
         Left            =   465
         TabIndex        =   10
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label lblAfter 
         Caption         =   "After:"
         Height          =   210
         Left            =   2970
         TabIndex        =   9
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblBefore 
         Caption         =   "Before:"
         Height          =   285
         Left            =   1530
         TabIndex        =   8
         Top             =   240
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   2775
      ScaleHeight     =   3045
      ScaleWidth      =   5775
      TabIndex        =   1
      Top             =   120
      Width           =   5805
      Begin VB.TextBox txtCommandPlacer 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   -30
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   225
         Width           =   5820
      End
      Begin VB.Line linMargin 
         BorderColor     =   &H00000080&
         X1              =   1485
         X2              =   1485
         Y1              =   3045
         Y2              =   -270
      End
      Begin VB.Label lblBottomSample 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   15
         TabIndex        =   3
         Top             =   2790
         Width           =   5760
      End
      Begin VB.Label lblTopSample 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   5775
      End
   End
   Begin VB.ListBox lstCommands 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   2280
   End
End
Attribute VB_Name = "frmConfigFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intBlockVal As Integer
Dim blnChanged As Boolean

'===============================================================================
'Name:          frmConfigFormat
'Purpose:       This form allows the user to configure settings for automatic
'               blocking and whitespacing of VB code within the code editor window.
'
'Returns:       None
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:  Saves data in "comands.mdb". Reference to this file is set in
'               Saves data in "comands.mdb". Reference to this file is set in the registry and will be prompted for if missing.             '
'===============================================================================


Private Sub Text1_Change()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub chkForceLeft_Click()

    If chkForceLeft Then
        txtTabBefore = -99
    End If

End Sub

Private Sub cmdAdd_Click()
    Static blnEditMode  As Boolean
    blnEditMode = Not blnEditMode
    
    If blnEditMode Then
        cmdAdd.Caption = "&Save"
        txtAddCommand.BackColor = vbWindowBackground
        txtAddCommand.Enabled = True
    Else
        cmdAdd.Caption = "&Add"
        txtAddCommand.BackColor = vbButtonFace
        AddCommand
    End If
    

    txtAddCommand.Text = ""
    PopulateCommandList

End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub



Private Sub cmdRemove_Click()
If lstCommands.SelCount > 0 Then
    RemoveCommand
    PopulateCommandList
End If
End Sub

Private Sub cmdSave_Click()
    UpdateTable
    blnChanged = False
End Sub


Private Sub Form_Load()

PopulateCommandList



End Sub

Private Sub lstCommands_Click()

PopulateVals
BuildCommand



End Sub

Private Sub lstCommands_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If blnChanged Then
            If MsgBox("You have made changes to this command...save?", vbQuestion + vbYesNo) = vbYes Then
                UpdateTable

            End If
        End If
    End If
blnChanged = False
End Sub

Private Sub txtAddCommand_Change()
    cmdRemove.Enabled = False
End Sub


Private Sub txtAddCommand_GotFocus()
    cmdRemove.Enabled = False
End Sub

Private Sub txtTabAfter_Change()
 

 If Val(txtTabAfter) > -99 And Val(txtTabAfter) < -5 Then
    txtTabAfter = -5
 End If
 BuildCommand
End Sub

Private Sub txtTabBefore_Change()
 
 If Val(txtTabBefore) > -99 And Val(txtTabBefore) < -5 Then
    txtTabBefore = -5

 End If
 chkForceLeft.Value = 0
 BuildCommand
 

End Sub


Private Function BuildString(BaseValue As String, RepeatCount As Integer) As String
       
Dim intCountLoop As Integer

For intCountLoop = 1 To RepeatCount

    BuildString = BuildString & BaseValue
Next intCountLoop



End Function

Private Sub txtWhitespaceafter_Change()

    BuildCommand
End Sub

Private Sub txtWhitespaceBefore_Change()

    BuildCommand
End Sub

'===============================================================================
'Name:          BuildCommand
'Purpose:       Formats the command in the example windows to match the selected
'               Formats the command in the example windows to match the selected parameters on the form.
'Returns:       None
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:
'===============================================================================


Private Function BuildCommand() As String
Dim strBuild As String
Dim intTabsAfter As Integer


If chkForceLeft <> 0 Then
    intTabsAfter = Val(txtTabAfter) - 5
Else
    intTabsAfter = Val(txtTabAfter)
End If


strBuild = BuildString(Space(5), 5) & "Preceding line of code..."
strBuild = strBuild & BuildString(vbCrLf, Val(txtWhitespaceBefore)) & vbCrLf
strBuild = strBuild & BuildString(Space(5), Val(txtTabBefore) + 5) & lstCommands.Text
strBuild = strBuild & BuildString(vbCrLf, Val(txtWhitespaceafter)) & vbCrLf
strBuild = strBuild & BuildString(Space(5), intTabsAfter + Val(txtTabAfter) + 5) & "Following Line of code.."

txtCommandPlacer = strBuild

End Function

'===============================================================================
'Name:          PopulateCommandList
'Purpose:       Populates the commands list with all commands in the database.
'
'Returns:       None
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:
'===============================================================================


Sub PopulateCommandList()

Dim rsCommands As ADODB.Recordset
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim intCmdCount As Integer
Dim strDbPath As String



strDbPath = GetSetting(App.Title, "startup", "datapath", App.Path & "\commands.mdb")

While Dir(strDbPath) = ""
    MsgBox "Commands data file not found at: " & strDbPath & ".", vbExclamation
    strDbPath = InputBox("Enter the location of the command.mdb file:")
    If strDbPath = "" Then
        MsgBox "Invalid path"
        Unload Me
        Exit Sub
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
    .CommandText = "SELECT * FROM Commands ORDER BY CommandText"
    .ActiveConnection = conn
    Set rsCommands = .Execute
End With

Set conn = Nothing
Set cmd = Nothing

lstCommands.Clear

While Not rsCommands.EOF
    lstCommands.AddItem rsCommands.Fields("CommandText") & ""
    lstCommands.ItemData(lstCommands.ListCount - 1) = rsCommands.Fields("ID")
    ReDim Preserve intCommandIDs(intCmdCount)
    intCommandIDs(intCmdCount) = rsCommands.Fields("ID")
    rsCommands.MoveNext
    intCmdCount = intCmdCount + 1
Wend


End Sub

'===============================================================================
'Name:          UpdateTable
'Purpose:       Saves changes made to configuration info to the commands table.
'
'Returns:       None
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:
'===============================================================================


Private Sub UpdateTable()

Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim intCmdCount As Integer
Dim strUpdate As String
Dim intCommandID As Integer
Dim strDbPath As String
Dim intBlockBefore As Integer

Set conn = New ADODB.Connection
Set cmd = New ADODB.Command

strDbPath = GetSetting(App.Title, "startup", "datapath", App.Path)
While Dir(strDbPath) = ""
    strDbPath = InputBox("Enter the location of the command.mdb file:")
    If strDbPath = "" Then
        MsgBox "Invalid path"
    End If
Wend


With conn
    .ConnectionString = "Provider=" & _
      "Microsoft.Jet.OLEDB." & _
      "4.0;Data Source=" & _
      strDbPath
    .Open
End With

If chkForceLeft.Value <> 0 Then
    intBlockBefore = -99
Else
    intBlockBefore = Val(txtTabBefore)
End If



intCommandID = intCommandIDs(lstCommands.ListIndex)

    strUpdate = "UPDATE Commands SET BlockBefore=" & intBlockBefore _
                                & ",BlockAfter=" & Val(txtTabAfter) _
                                & ",WhitespaceBefore=" & Val(txtWhitespaceBefore) _
                                & ",Whitespaceafter=" & Val(txtWhitespaceafter) _
                                & " WHERE ID  = " & intCommandID
                                


With cmd
    .CommandText = strUpdate
    .ActiveConnection = conn
    .Execute
End With

Set conn = Nothing
Set cmd = Nothing


End Sub

'===============================================================================
'Name:          PopulateVals
'Purpose:       Populate the information in the text boxes for this command.                '
'
'Returns:
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:
'===============================================================================


Private Sub PopulateVals()

Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim intCmdCount As Integer
Dim strPopulate As String
Dim rsPopulate As ADODB.Recordset
Dim intCommandID As Integer
Dim strDbPath As String

Set conn = New ADODB.Connection
Set cmd = New ADODB.Command

strDbPath = GetSetting(App.Title, "startup", "datapath", App.Path)
While Dir(strDbPath) = ""
    strDbPath = InputBox("Enter the location of the command.mdb file:")
    If strDbPath = "" Then
        MsgBox "Invalid path"
    End If
Wend

With conn
    .ConnectionString = "Provider=" & _
      "Microsoft.Jet.OLEDB." & _
      "4.0;Data Source=" & _
      strDbPath
    .Open
End With

intCommandID = intCommandIDs(lstCommands.ListIndex)
If lstCommands.ListIndex >= 0 Then
    strPopulate = "SELECT * FROM Commands WHERE ID = " & intCommandID
    
    With cmd
        .CommandText = strPopulate
        .ActiveConnection = conn
       Set rsPopulate = .Execute
    End With
    
    Set conn = Nothing
    Set cmd = Nothing
    
    While Not rsPopulate.EOF
    
        txtTabBefore = Val(rsPopulate.Fields("BlockBefore"))
        txtTabAfter = Val(rsPopulate.Fields("BlockAfter"))
        txtWhitespaceBefore = Val(rsPopulate.Fields("WhitespaceBefore"))
        txtWhitespaceafter = Val(rsPopulate.Fields("WhitespaceAfter"))
        'txtTabBefore = rsPopulate.Fields("BlockBefore")
        
        If rsPopulate.Fields("BlockBefore") = -99 Then
            chkForceLeft.Value = 1
        Else
            chkForceLeft.Value = 0
        End If
        rsPopulate.MoveNext
    Wend
    
    If txtTabBefore = -99 Then
  '      txtTabBefore = 0
    End If
    
End If
End Sub

Private Sub AddCommand()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim strDbPath As String
Dim strSQL As String


Set conn = New ADODB.Connection
Set cmd = New ADODB.Command

strDbPath = GetSetting(App.Title, "startup", "datapath", App.Path)
While Dir(strDbPath) = ""
    strDbPath = InputBox("Enter the location of the command.mdb file:")
    If strDbPath = "" Then
        MsgBox "Invalid path"
    End If
Wend

With conn
    .ConnectionString = "Provider=" & _
      "Microsoft.Jet.OLEDB." & _
      "4.0;Data Source=" & _
      strDbPath
    .Open
End With


    strSQL = "INSERT INTO Commands  (CommandText) VALUES ( '" & txtAddCommand.Text & "');"
    
    With cmd
        .CommandText = strSQL
        .ActiveConnection = conn
        .Execute
    End With
    
    Set conn = Nothing
    Set cmd = Nothing
    

End Sub


'===============================================================================
'Name:          UpdateTable
'Purpose:       Saves changes made to configuration info to the commands table.
'
'Returns:       None
'Created By:    M@ (Matthew M. Roberts)
'Date:          1/3/2001
'Comments:
'===============================================================================


Private Sub RemoveCommand()

Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim intCmdCount As Integer
Dim strSQL As String
Dim intCommandID As Integer
Dim strDbPath As String
Dim intBlockBefore As Integer

Set conn = New ADODB.Connection
Set cmd = New ADODB.Command

strDbPath = GetSetting(App.Title, "startup", "datapath", App.Path)
While Dir(strDbPath) = ""
    strDbPath = InputBox("Enter the location of the command.mdb file:")
    If strDbPath = "" Then
        MsgBox "Invalid path"
    End If
Wend


With conn
    .ConnectionString = "Provider=" & _
      "Microsoft.Jet.OLEDB." & _
      "4.0;Data Source=" & _
      strDbPath
    .Open
End With

If chkForceLeft.Value <> 0 Then
    intBlockBefore = -99
Else
    intBlockBefore = Val(txtTabBefore)
End If


intCommandID = intCommandIDs(lstCommands.ListIndex)

    strSQL = "DELETE FROM Commands WHERE ID = " & intCommandID
                                
With cmd
    .CommandText = strSQL
    .ActiveConnection = conn
    .Execute
End With

Set conn = Nothing
Set cmd = Nothing


End Sub


