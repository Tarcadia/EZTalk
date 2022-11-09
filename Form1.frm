VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FormTalk 
   Caption         =   "EZTalk"
   ClientHeight    =   3360
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6870
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Rtb 
      Appearance      =   0  'Flat
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   480
      Width           =   6615
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "发送"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox TxtSend 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   5535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "这是："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "FormTalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public IPNum As String
Public MyName As String
Public MyPortNum As Long
Public UserPortNum As Long
Public UserName As String

Private Sub CmdSend_Click()

    On Error Resume Next
    '在键入Entre时，立即将其发送出去。

    Dim Str As String
    Str = TxtSend.Text

    If Trim(Str) = "" Then
        MsgBox "发送内容不能为空": Exit Sub

    ElseIf IsBadWord(Str) Then
        MsgBox "不许骂脏话！代表人民和谐你！"
        Actions.MsgSB: Actions.MsgFool
        Exit Sub
    End If

    Winsock1.SendData MyName & "(" & Winsock1.LocalIP & ")：:" & TxtSend.Text

    Dim StrData As String
    StrData = "我：:" & TxtSend.Text

    Dim a() As String
    a = Split(StrData, ":")
    Dim S
    Dim X As String
    Dim Able As Boolean
    Able = True
    For Each S In a
        If Able Then
            Able = False
            X = vbCrLf + S + vbCrLf
        Else
            X = X + CrlfFun(S)
        End If
    Next S

    Rtb.Text = Rtb.Text + X
    TxtSend.Text = ""
    Rtb.SelStart = Len(Rtb.Text)

End Sub

Private Sub Form_Load()
  BadWordVip = False
  On Error Resume Next
    IPNum = InputBox("请输入【Ta】的 {IP} 或 {计算机名} ")
    MyName = InputBox("请输入【我】的 {昵称} ")
    MyPortNum = 1033
    UserPortNum = 1033
    Label1.Caption = "这是" & MyName & "与" & IPNum & "的聊天"

        With Winsock1
            .RemoteHost = IPNum
            .RemotePort = UserPortNum
            .Bind MyPortNum '绑定到本地的端口。
        End With
    
    Winsock1.SendData "!!!:$" & MyName + "已经上线"

    Exit Sub

End Sub

Private Sub Form_Resize()
    If (Me.ScaleWidth > 2000) And (Me.ScaleHeight > 2000) Then
        Rtb.Width = Me.ScaleWidth - 240
        Rtb.Height = Me.ScaleHeight - 960
        Label1.Width = Me.ScaleWidth - 480
        TxtSend.Top = Rtb.Top + Rtb.Height + 120
        TxtSend.Width = Me.ScaleWidth - CmdSend.Width - 600
        CmdSend.Top = TxtSend.Top
        CmdSend.Left = Me.ScaleWidth - 120 - CmdSend.Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Errors2
    Winsock1.SendData "!!!:$" & MyName + "已经下线"

Errors2:
    If Err.Number = 10014 Then MsgBox "对方不在线" Else Exit Sub
End Sub

Private Sub Label1_DblClick()
    FormDlgLogIn.Types = LogSUC
    FormDlgLogIn.Show , Me
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim StrData As String
    Winsock1.GetData StrData

    Dim a() As String
    a = Split(StrData, ":")
    Dim S
    Dim X As String
    Dim Able As Boolean
    Able = True
    For Each S In a
        If Able Then
            Able = False
            X = vbCrLf + S + vbCrLf
        Else
            X = X + DoAction(S)
        End If
    Next S
    Rtb.Text = Rtb.Text + X
    Rtb.SelStart = Len(Rtb.Text)
End Sub

Public Function DoAction(ByVal Str As String) As String
    If (Mid(Str, 1, 1) <> "#") And (Mid(Str, 1, 1) <> "$") Then
        DoAction = IIf(Trim(Str) = "", "", "------·" + Str + vbCrLf)

    ElseIf Mid(Str, 1, 1) = "$" Then
        DoAction = "------·鸡毛信：" + Right(Str, Len(Str) - 1) + vbCrLf
        MsgBox Right(Str, Len(Str) - 1)

    ElseIf Mid(Str, 1, 1) = "#" Then
        Select Case LCase(Str)

            Case "#shutdown"
                DoAction = MsgShutdown + vbCrLf

            Case "#unshutdown"
                DoAction = MsgUnshutdown + vbCrLf

            Case "#sb"
                DoAction = MsgSB + vbCrLf

            Case "#shake"
                DoAction = MsgShake + vbCrLf

            Case "#tillshake"
                DoAction = MsgTillShake + vbCrLf
            
            Case "#stopshake"
                DoAction = MsgStopShake + vbCrLf
            
            Case "#fool"
                DoAction = MsgFool + vbCrLf
            
            Case "#music"
                DoAction = MsgMusic
            
            Case "#hello"
                DoAction = MsgHello
            
            Case "#laugh"
                DoAction = MsgLaugh
            
            Case "#game"
                DoAction = MsgGame
            
            Case "#topmost"
                DoAction = MsgTopMost

        End Select
    End If
End Function

Public Function CrlfFun(ByVal Str As String) As String
    On Error GoTo ErrFun3
    If (Mid(Str, 1, 1) <> "#") And (Mid(Str, 1, 1) <> "$") Then
        CrlfFun = IIf(Trim(Str) = "", "", "------·" + Str + vbCrLf)
    ElseIf Mid(Str, 1, 1) = "$" Then
        CrlfFun = "------·鸡毛信：" + Right(Str, Len(Str) - 1) + vbCrLf
    ElseIf Mid(Str, 1, 1) = "#" Then
        Select Case LCase(Str)

            Case "#shutdown"
                CrlfFun = "------·你发送了一个关机命令" + vbCrLf

            Case "#unshutdown"
                CrlfFun = "------·你发送了一个取消关机命令" + vbCrLf

            Case "#sb"
                CrlfFun = "------·你发送了一个骂人命令" + vbCrLf

            Case "#shake"
                CrlfFun = "------·你发送了一个窗口抖动命令（感谢腾讯QQ的灵感）" + vbCrLf
            
            Case "#tillshake"
                CrlfFun = "------·你发送了一个连续窗口抖动命令（感谢腾讯QQ的灵感）" + vbCrLf
            
            Case "#stopshake"
                CrlfFun = "------·你发送了一个停止窗口抖动命令" + vbCrLf
            
            Case "#fool"
                CrlfFun = "------·你发送了一个很二提问" + vbCrLf
            
            Case "#music"
                CrlfFun = "------·你发送了一个响声命令" + vbCrLf
            
            Case "#hello"
                CrlfFun = "------·你发送了一个问好命令" + vbCrLf
            
            Case "#laugh"
                CrlfFun = "------·你发送了一个傻笑命令" + vbCrLf
            
            Case "#game"
                CrlfFun = "------·你发送了一个游戏命令" + vbCrLf
                        
            Case "#topmost"
                CrlfFun = "------·你发送了一个窗口置顶命令" + vbCrLf

        End Select

    End If
ErrFun3:
    Exit Function
End Function

