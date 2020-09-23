VERSION 5.00
Begin VB.Form FrmBrowse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Directory"
   ClientHeight    =   4935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3375
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "FrmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    FrmASC.Text1.Text = Dir1.Path
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Drive1_Change()
    Dim tDrive As String
    On Error GoTo ErrorHandler
    tDrive = Left(Dir1.Path, 2)
    Dir1.Path = UCase(Drive1.Drive)
    Exit Sub
ErrorHandler:
    If Err.Number = 68 Then
        MsgBox "The drive is empty.", vbCritical, "Unknown Media"
        Drive1.Drive = tDrive
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Drive1.Drive = Left(FrmASC.Text1.Text, 2)
    Dir1.Path = FrmASC.Text1.Text
    Exit Sub
ErrorHandler:
    Dir1.Path = UCase(Drive1.Drive)
End Sub
