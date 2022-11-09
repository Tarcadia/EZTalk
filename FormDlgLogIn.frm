VERSION 5.00
Begin VB.Form FormDlgLogIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log In 登录"
   ClientHeight    =   1110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "请输入密钥："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FormDlgLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Types As LogType
Public Enum LogType
    LogSUC = 0
    LogBWV = 1
End Enum

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    If Me.Text1 <> "Dreamer" Then
        Unload Me
    ElseIf Text1.PasswordChar = "#" Then
        MsgBox ("shutdown,unshutdown,sb,shake" + vbCrLf + "tillshake,stopshake,fool,music" + vbCrLf + "laugh,game,hello,topmost")
    End If
End Sub

Private Sub Text1_Change()
    If Me.Text1.PasswordChar = "&" Then
        Me.Text1.PasswordChar = "#"
    ElseIf Me.Text1.PasswordChar = "#" Then
        Me.Text1.PasswordChar = "*"
    ElseIf Me.Text1.PasswordChar = "*" Then
        Me.Text1.PasswordChar = "&"
    End If
End Sub
