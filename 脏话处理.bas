Attribute VB_Name = "�໰����"

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
 Or (strLC Like "*����*") _
 Or (strLC Like "*�H*") _
 Or (strLC Like "*��*") _
 Or (strLC Like "*ܳ*") _
 Or (strLC Like "*2b*") _
 Or (strLC Like "*zb*")

    Ans = Ans Or _
    ((strLC Like "*sb*") And Not (strLC Like "*#sb")) _
 Or (strLC Like "*����*") _
 Or (strLC Like "*tmd*") _
 Or (strLC Like "*�����*") _
 Or (strLC Like "*ɵ��*") _
 Or (strLC Like "*ɷ��*") _
 Or (strLC Like "*����*") _
 Or (strLC Like "*������*")

    Ans = Ans Or _
    (strLC Like "*bitch*") _
 Or (strLC Like "*asshole*") _
 Or (strLC Like "*dick*") _
 Or (strLC Like "*damn*") _
 Or (strLC Like "*����*") _
 Or (strLC Like "*�ҿ�*") _
 Or (strLC Like "*cao*") _
 Or (strLC Like "*kao*")

    Ans = Ans Or _
    (strLC Like "*����*") _
 Or (strLC Like "*��ϣƥ*") _
 Or (strLC Like "*���˸�*") _
 Or (strLC Like "*�Ҳ�*") _
 Or (strLC Like "*��*") _
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
 Or (strLC Like "*��ˮ��*") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "") _
 Or (strLC Like "")

    If Dir("�໰�б�.bwl") <> "" Then
        Open "�໰�б�.bwl" For Input As #File
        While Not EOF(File)
            Line Input #File, Dat
            Ans = Ans Or (strLC Like Dat)
        Wend
        Close #File
    End If
    IsBadWord = Ans
End Function
