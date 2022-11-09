Attribute VB_Name = "脏话处理"

Public Function IsBadWord(Str As String) As Boolean

    Dim strLC As String
    Dim Ans As Boolean
    Dim File As Integer
    Dim Dat As String
    strLC = LCase(Str)
    File = FreeFile

    Ans = _
    (strLC Like "*fuck*") _
 Or (strLC Like "*tmd*") _
 Or (strLC Like "*尼玛*") _
 Or (strLC Like "*肏*") _
 Or (strLC Like "*屄*") _
 Or (strLC Like "*艹*") _
 Or (strLC Like "*2b*") _
 Or (strLC Like "*zb*")

    Ans = Ans Or _
    ((strLC Like "*sb*") And Not (strLC Like "*#sb")) _
 Or (strLC Like "*你妈*") _
 Or (strLC Like "*tmd*") _
 Or (strLC Like "*他妈的*") _
 Or (strLC Like "*傻逼*") _
 Or (strLC Like "*煞笔*") _
 Or (strLC Like "*二逼*") _
 Or (strLC Like "*你他妈*")

    Ans = Ans Or _
    (strLC Like "*bitch*") _
 Or (strLC Like "*asshole*") _
 Or (strLC Like "*dick*") _
 Or (strLC Like "*damn*") _
 Or (strLC Like "*你娘*") _
 Or (strLC Like "*我靠*") _
 Or (strLC Like "*cao*") _
 Or (strLC Like "*kao*")

    Ans = Ans Or _
    (strLC Like "*靠！*") _
 Or (strLC Like "*娘希匹*") _
 Or (strLC Like "*妈了个*") _
 Or (strLC Like "*我操*") _
 Or (strLC Like "*呸*") _
 Or (strLC Like "*dork*") _
 Or (strLC Like "*nerd*") _
 Or (strLC Like "*geek*")

    Ans = Ans Or _
    (strLC Like "*dammit*") _
 Or (strLC Like "*phycho*") _
 Or (strLC Like "*shit*") _
 Or (strLC Like "*dense*") _
 Or (strLC Like "*stupid*") _
 Or (strLC Like "*foolish*") _
 Or (strLC Like "*bastard*") _
 Or (strLC Like "*faq*")
 
     Ans = Ans Or _
    (strLC Like "*dammit*") _
 Or (strLC Like "*陈水扁*") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "")

    If Dir("脏话列表.bwl") <> "" Then
        Open "脏话列表.bwl" For Input As #File
        While Not EOF(File)
            Line Input #File, Dat
            Ans = Ans Or (strLC Like Dat)
        Wend
        Close #File
    End If
    IsBadWord = Ans
End Function
