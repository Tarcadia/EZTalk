Attribute VB_Name = "Actions"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Shakeable As Boolean
Private Topable As Boolean

Public Function MsgShutdown() As String

    MsgShutdown = ""
    MsgBox "Sorry, your computer has some problem.", vbOKOnly + vbInformation, "Error"
    MsgBox "Your IP address is not fit in this net.", vbOKOnly + vbInformation, "Error"
    MsgBox "Your comprter will shutdown after you click the [OK Button]", vbOKOnly + vbInformation, "Error"
    Shell "C:\windows\system32\shutdown.exe -s -t 10 -c ����10�����޸�����IP"

End Function

Public Function MsgUnshutdown() As String
    MsgUnshutdown = ""
    Shell "C:\WINDOWS\System32\shutdown.exe -a"
End Function

Public Function MsgSB() As String
    MsgSB = ""
    While InputBox("���Ǵ�ɵ*", , "���Ǵ�ɵ*") <> "���Ǵ�ɵ*"
    Wend
    FormTalk.TxtSend.Text = "���Ǵ�ɵ*"
    FormTalk.CmdSend.Value = True
End Function

Public Function MsgShake() As String

    On Error GoTo ErrShake1

    MsgShake = "------���Է�������һ���������ڣ���л��ѶQQ���ҵ���У�" + vbCrLf
    Dim I As Integer
    For I = 1 To 5
        FormTalk.Top = FormTalk.Top + 100
        Sleep (50): DoEvents
        FormTalk.Left = FormTalk.Left + 100
        Sleep (50): DoEvents
        FormTalk.Top = FormTalk.Top - 100
        Sleep (50): DoEvents
        FormTalk.Left = FormTalk.Left - 100
        Sleep (50): DoEvents
    Next I

ErrShake1:
    If Err.Number = 384 Then Exit Function

End Function

Public Function MsgTillShake() As String

    On Error GoTo ErrShake2

    MsgTillShake = "------���Է�������һ���������ڣ���л��ѶQQ���ҵ���У�" + vbCrLf
    Shakeable = True
    While Shakeable
        FormTalk.Top = FormTalk.Top + 100
        Sleep (50): DoEvents
        FormTalk.Left = FormTalk.Left + 100
        Sleep (50): DoEvents
        FormTalk.Top = FormTalk.Top - 100
        Sleep (50): DoEvents
        FormTalk.Left = FormTalk.Left - 100
        Sleep (50): DoEvents
    Wend

ErrShake2:
    If Err.Number = 384 Then Exit Function

End Function

Public Function MsgStopShake() As String
    If Shakeable Then
        MsgStopShake = "------���Ǻ�" + vbCrLf
        Shakeable = False
        Sleep (500): DoEvents
    End If
End Function

Public Function MsgFool() As String
    Dim Ans As Integer

    MsgFool = "------���Է������˺�2������" + vbCrLf

FoolS1:
    Ans = MsgBox("���Ǳ�����", vbYesNo + vbQuestion)
    If Ans = vbNo Then GoTo FoolS1
    FormTalk.TxtSend.Text = "�ҳ������Ǳ���:��P"
    FormTalk.CmdSend.Value = True

FoolS2:
    Ans = MsgBox("��ܴ���", vbYesNo + vbQuestion)
    If Ans = vbNo Then GoTo FoolS2
    FormTalk.TxtSend.Text = "�ҳ����Һܴ�:��P"
    FormTalk.CmdSend.Value = True

FoolS3:
    Ans = MsgBox("���̬��", vbYesNo + vbQuestion)
    If Ans = vbNo Then GoTo FoolS3
    FormTalk.TxtSend.Text = "�ҳ����ұ�̬:��P"
    FormTalk.CmdSend.Value = True


End Function

Public Function MsgMusic() As String

    MsgMusic = "------���ǺǺǣ�������" + vbCrLf

    Dim I As Integer
    For I = 1 To 100
        VBA.Beep
        Sleep (100): DoEvents
    Next I

End Function

Public Function MsgHello()

    MsgHello = "------���Է����������������������Ƶģ�������������" + vbCrLf
    MsgBox "�����"
    MsgBox "��ոյ���ǡ�ȷ������"
    MsgBox "�Ҿ�˵������϶����ˡ�ȷ��������û˵��ɣ�"
    MsgBox "����ô���ǵ㡾ȷ����ѽ�����ǲ��ǻ���㡾ȷ������"
    MsgBox "���ϵ�ͬ���Ķ������������ˣ��㲻����ĺã�"
    MsgBox "��Ҫ���ٵ㡾ȷ�������ҾͲ��������ˣ�"
    MsgBox "�㻹�ҵ㡾ȷ�������ð������ϴ����ˣ�"
    MsgBox "�Է�������", vbCritical

End Function

Public Function MsgLaugh()

    MsgLaugh = "�Է�ɵЦ��һͨ" + vbCrLf
    MsgBox "������"
    MsgBox "��������"
    MsgBox "�ǺǺǺ�"
    MsgBox "������"
    MsgBox "�ٺٺٺٺ�"
    MsgBox "û���ˣ�������"

End Function

Public Function MsgGame()

    MsgGame = "�Է�����������������أ����������ǰ���һ�ֲ߻��ģ�hoho" + vbCrLf

    Dim Ans As Integer
LoopGame:
    Ans = MsgBox("�����������������Ϸ", vbYesNo)
    If Ans = vbYes Then MsgBox "̫����" Else MsgBox "������Ҫ�": GoTo LoopGame
LoopGameBegin:
    Ans = MsgBox("��Ϸ�������������ģ��������⣬��ش𣬲���ش��ǣ�����", vbYesNo)
    If Ans = vbYes Then MsgBox "����˵����ѡ���ǡ�����ͷ����": GoTo LoopGameBegin Else MsgBox "�����𣿻�����"
    MsgBox "�Ժ���Ҳ����������"

End Function

Public Function MsgTopMost()

    MsgTopMost = "�Է�������һ�������ö����������"
    mdlTopMost.TopMost FormTalk

End Function
