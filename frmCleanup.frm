VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCleanup 
   Caption         =   "Code Cleanup"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtSource 
      Height          =   6225
      Left            =   135
      TabIndex        =   1
      Top             =   495
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   10980
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   ""
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "&Format"
      Height          =   375
      Left            =   8460
      TabIndex        =   0
      Top             =   45
      Width           =   1005
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuConfigure 
         Caption         =   "&Configure"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "frmCleanup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFormat_Click()
    txtSource.Text = FormatCode(txtSource.Text)
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
    With txtSource
        .Width = Me.Width - .Left - 200
        .Height = Me.Height - .Top - 800
    
        cmdFormat.Left = .Left + .Width - cmdFormat.Width
    End With
    

    
End Sub

Private Sub mnuClear_Click()
    txtSource.Text = ""
End Sub

Private Sub mnuConfigure_Click()
    On Error Resume Next
    frmConfigFormat.Show
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText txtSource.Text
End Sub

Private Sub mnuPaste_Click()
    txtSource = Clipboard.GetText
End Sub

Private Sub txtSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuRightClick
    End If
End Sub
