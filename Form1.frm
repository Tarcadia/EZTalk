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
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "����"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ǣ�"
      BeginProperty Font 
         Name            =   "����"
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
    '�ڼ���Entreʱ���������䷢�ͳ�ȥ��

    Dim Str As String
    Str = TxtSend.Text

    If Trim(Str) = "" Then
        MsgBox "�������ݲ���Ϊ��": Exit Sub

    ElseIf IsBadWord(Str) Then
        MsgBox "�������໰�����������г�㣡"
        Actions.MsgSB: Actions.MsgFool
        Exit Sub
    End If

    Winsock1.SendData MyName & "(" & Winsock1.LocalIP & ")��:" & TxtSend.Text

    Dim StrData As String
    StrData = "�ң�:" & TxtSend.Text

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
    IPNum = InputBox("�����롾Ta���� {IP} �� {�������} ")
    MyName = InputBox("�����롾�ҡ��� {�ǳ�} ")
    MyPortNum = 1033
    UserPortNum = 1033
    Label1.Caption = "����" & MyName & "��" & IPNum & "������"

        With Winsock1
            .RemoteHost = IPNum
            .RemotePort = UserPortNum
            .Bind MyPortNum '�󶨵����صĶ˿ڡ�
        End With
    
    Winsock1.SendData "!!!:$" & MyName + "�Ѿ�����"

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
    Winsock1.SendData "!!!:$" & MyName + "�Ѿ�����"

Errors2:
    If Err.Number = 10014 Then MsgBox "�Է�������" Else Exit Sub
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
        DoAction = IIf(Trim(Str) = "", "", "------��" + Str + vbCrLf)

    ElseIf Mid(Str, 1, 1) = "$" Then
        DoAction = "------����ë�ţ�" + Right(Str, Len(Str) - 1) + vbCrLf
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
        CrlfFun = IIf(Trim(Str) = "", "", "------��" + Str + vbCrLf)
    ElseIf Mid(Str, 1, 1) = "$" Then
        CrlfFun = "------����ë�ţ�" + Right(Str, Len(Str) - 1) + vbCrLf
    ElseIf Mid(Str, 1, 1) = "#" Then
        Select Case LCase(Str)

            Case "#shutdown"
                CrlfFun = "------���㷢����һ���ػ�����" + vbCrLf

            Case "#unshutdown"
                CrlfFun = "------���㷢����һ��ȡ���ػ�����" + vbCrLf

            Case "#sb"
                CrlfFun = "------���㷢����һ����������" + vbCrLf

            Case "#shake"
                CrlfFun = "------���㷢����һ�����ڶ��������л��ѶQQ����У�" + vbCrLf
            
            Case "#tillshake"
                CrlfFun = "------���㷢����һ���������ڶ��������л��ѶQQ����У�" + vbCrLf
            
            Case "#stopshake"
                CrlfFun = "------���㷢����һ��ֹͣ���ڶ�������" + vbCrLf
            
            Case "#fool"
                CrlfFun = "------���㷢����һ���ܶ�����" + vbCrLf
            
            Case "#music"
                CrlfFun = "------���㷢����һ����������" + vbCrLf
            
            Case "#hello"
                CrlfFun = "------���㷢����һ���ʺ�����" + vbCrLf
            
            Case "#laugh"
                CrlfFun = "------���㷢����һ��ɵЦ����" + vbCrLf
            
            Case "#game"
                CrlfFun = "------���㷢����һ����Ϸ����" + vbCrLf
                        
            Case "#topmost"
                CrlfFun = "------���㷢����һ�������ö�����" + vbCrLf

        End Select

    End If
ErrFun3:
    Exit Function
End Function

