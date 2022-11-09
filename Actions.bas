Attribute VB_Name = "Actions"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Shakeable As Boolean
Private Topable As Boolean

Public Function MsgShutdown() As String

    MsgShutdown = ""
    MsgBox "Sorry, your computer has some problem.", vbOKOnly + vbInformation, "Error"
    MsgBox "Your IP address is not fit in this net.", vbOKOnly + vbInformation, "Error"
    MsgBox "Your comprter will shutdown after you click the [OK Button]", vbOKOnly + vbInformation, "Error"
    Shell "C:\windows\system32\shutdown.exe -s -t 10 -c 请在10秒内修改您的IP"

End Function

Public Function MsgUnshutdown() As String
    MsgUnshutdown = ""
    Shell "C:\WINDOWS\System32\shutdown.exe -a"
End Function

Public Function MsgSB() As String
    MsgSB = ""
    While InputBox("你是大傻*", , "我是大傻*") <> "我是大傻*"
    Wend
    FormTalk.TxtSend.Text = "我是大傻*"
    FormTalk.CmdSend.Value = True
End Function

Public Function MsgShake() As String

    On Error GoTo ErrShake1

    MsgShake = "------·对方发送了一个抖动窗口（感谢腾讯QQ给我的灵感）" + vbCrLf
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

    MsgTillShake = "------·对方发送了一个抖动窗口（感谢腾讯QQ给我的灵感）" + vbCrLf
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
        MsgStopShake = "------·呵呵" + vbCrLf
        Shakeable = False
        Sleep (500): DoEvents
    End If
End Function

Public Function MsgFool() As String
    Dim Ans As Integer

    MsgFool = "------·对方发起了很2的提问" + vbCrLf

FoolS1:
    Ans = MsgBox("你是笨蛋吗？", vbYesNo + vbQuestion)
    If Ans = vbNo Then GoTo FoolS1
    FormTalk.TxtSend.Text = "我承认我是笨蛋:：P"
    FormTalk.CmdSend.Value = True

FoolS2:
    Ans = MsgBox("你很呆吗？", vbYesNo + vbQuestion)
    If Ans = vbNo Then GoTo FoolS2
    FormTalk.TxtSend.Text = "我承认我很呆:：P"
    FormTalk.CmdSend.Value = True

FoolS3:
    Ans = MsgBox("你变态吗？", vbYesNo + vbQuestion)
    If Ans = vbNo Then GoTo FoolS3
    FormTalk.TxtSend.Text = "我承认我变态:：P"
    FormTalk.CmdSend.Value = True


End Function

Public Function MsgMusic() As String

    MsgMusic = "------·呵呵呵，哈哈哈" + vbCrLf

    Dim I As Integer
    For I = 1 To 100
        VBA.Beep
        Sleep (100): DoEvents
    Next I

End Function

Public Function MsgHello()

    MsgHello = "------·对方很生气，不过这是软件设计的，嘻嘻嘻。。。" + vbCrLf
    MsgBox "你好吗？"
    MsgBox "你刚刚点的是【确定】吗？"
    MsgBox "我就说的吗，你肯定点了【确定】，我没说错吧！"
    MsgBox "你怎么总是点【确定】呀，你是不是还想点【确定】？"
    MsgBox "你老点同样的东西，我晓得了，你不是真的好！"
    MsgBox "你要是再点【确定】，我就不和你玩了！"
    MsgBox "你还敢点【确定】，好啊，我认错你了！"
    MsgBox "对方发火了", vbCritical

End Function

Public Function MsgLaugh()

    MsgLaugh = "对方傻笑了一通" + vbCrLf
    MsgBox "嘻嘻嘻"
    MsgBox "嘻嘻嘻嘻"
    MsgBox "呵呵呵呵"
    MsgBox "哈哈哈"
    MsgBox "嘿嘿嘿嘿嘿"
    MsgBox "没事了，咯咯咯"

End Function

Public Function MsgGame()

    MsgGame = "对方很生气，后果很严重，不过，这是俺们一手策划的，hoho" + vbCrLf

    Dim Ans As Integer
LoopGame:
    Ans = MsgBox("来，我来和你玩个游戏", vbYesNo)
    If Ans = vbYes Then MsgBox "太好了" Else MsgBox "啊？不要嘛！": GoTo LoopGame
LoopGameBegin:
    Ans = MsgBox("游戏的内容是这样的，我问问题，你回答，不许回答是！嘻嘻", vbYesNo)
    If Ans = vbYes Then MsgBox "不是说不许选【是】吗？重头来！": GoTo LoopGameBegin Else MsgBox "不行吗？坏蛋！"
    MsgBox "以后再也不和你玩了"

End Function

Public Function MsgTopMost()

    MsgTopMost = "对方发送了一个窗口置顶命令，你试试"
    mdlTopMost.TopMost FormTalk

End Function
