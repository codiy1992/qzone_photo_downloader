Attribute VB_Name = "Mod_qqLogin"
Option Explicit
'-------------------------------------------------------------------------------------------
'-������QQ��½����
'-��������QQ�ʺţ�QQ���룬��index��status��
'-����ֵ����¼���̴�����Ϣ
'-------------------------------------------------------------------------------------------
Function qqLoginFun(U As String, P As String, Optional status As String = "online") As String  '��¼�����ӳ���
On Error Resume Next
Dim rnByteStr()         As Byte, rnStr As String          '��¼ ��������
Dim cliID     As Long                        '�����ȡ8λ��
Dim httpData    As String                      '��¼2 ����POST������
Dim cookie          As String * 1024               '����cookie
Dim rnUrl       As String
'---------------------------------------------------
Dim Key As String, vfCode As String  'ST ��������  KEY��¼HEXCODE  CODE ��֤��
'-------------------------------------------
 canLoginNew = False
If status = "" Then status = "online"
frmDownPhoto.frmLogin.Caption = "״̬:���ڵ�¼[ " & U & " ]"
'============================================================����Ƿ���Ҫ��֤��============================================
'----------------------------------------------------------
'outPutStr "-----------------�µ�QQ��½-------------------" & vbCrLf & vbTab & "---�����֤�뿪ʼ" & vbTab & Time & vbCrLf
'----------------------------------------------------------
frmDownPhoto.InetLogin.Execute "http://check.ptlogin2.qq.com/check?regmaster=&uin=" & U & "&appid=636014201&js_ver=10015&js_type=1&login_sig=cUadary30ZL35M8IrMqVmXGDDa*-VeXznLjl3IJrsKk4T2IRYZ94uaJ3up9ZqIFT&u1=http%3A%2F%2Fwww.qq.com%2Fqq2012%2FloginSuccess.htm&r=" & GetRnd(20)
Do While frmDownPhoto.InetLogin.StillExecuting
DoEvents
Loop ' �ȴ���������
'----------------------------------------------------------
'outPutStr "------�����֤�����" & vbTab & Time & vbCrLf
'----------------------------------------------------------
rnByteStr() = frmDownPhoto.InetLogin.GetChunk(0, icByteArray) '��ȡ��¼����
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8����
'====================================================================================================================================
Key = Mid(rnStr, InStr(rnStr, "\x"), 32) '��ȡ���ε�½������ļ�����Կ
If InStr(rnStr, "ptui_checkVC('0'") Then '����Ҫ��֤��
vfCode = Unmid(rnStr, "','", "','") 'ȡ���ε�½�������ʹ�ô���
Else '��Ҫ��֤��
   User = U
   Dialog.Show vbModal
   Do While Pdunload = False
      DoEvents
      vb_Sleep 200
   Loop
   vfCode = Yzmcode
End If
'===========================================��ʼ��½============================================================
   '----------------------------------��һ�ε�½----------------------------------------------------------
'�����ε�½������+������Կ+ʹ�ô�����md5����Encode(P, Key, code)����Ҫ��֤���ʱ��codeΪ��֤��
'----------------------------------------------------------
'outPutStr "------��һ�ε�½��ʼ" & rnStr & vbTab & Time & vbCrLf
'----------------------------------------------------------
frmDownPhoto.InetLogin.Execute "https://ssl.ptlogin2.qq.com/login?u=" & U & "&p=" & Encode(P, Key, vfCode) & "&verifycode=" & UCase(vfCode) & "&webqq_type=10&remember_uin=1&login2qq=0&aid=1003903&u1=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html%3Flogin2qq%3D0%26webqq_type%3D10&h=1&ptredirect=0&ptlang=2052&daid=164&from_ui=1&pttype=1&dumy=&fp=loginerroralert&action=2-20-14266&mibao_css=m_webqq&t=1&g=1&js_type=0&js_ver=10067&login_sig=XHPsCJZGJgJBy9Y9RmsrgKUOLcqdyO*H9veBTrYzaQusOEqwReADieCxsZWYiG1D", "GET", , "https://ui.ptlogin2.qq.com/cgi-bin/login?daid=164&target=self&style=5&mibao_css=m_webqq&appid=1003903&enable_qlogin=0&no_verifyimg=1&s_url=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html&f_url=loginerroralert&strong_login=0&login_state=10&t=20131202001" & vbCrLf & "Content-Type: utf-8"
Do While frmDownPhoto.InetLogin.StillExecuting
DoEvents
Loop ' �ȴ���������
'----------------------------------------------------------
'outPutStr "---------��һ�ε�½���" & vbTab & Time & vbCrLf
'----------------------------------------------------------
rnByteStr() = frmDownPhoto.InetLogin.GetChunk(0, icByteArray) '��ȡ��¼����
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8����
rnUrl = Unmid(rnStr, "ptuiCB('0','0','", "'")
'--------------------------------------------------------------------------------------------------
'outPutStr rnStr
If InStr(rnStr, "��¼�ɹ�") > 0 Then                                                     '��ʾ��¼�ɹ� Err��ʾ��¼ʧ��
        frmDownPhoto.frmLogin.Caption = "״̬:��ȡ��¼��Ϣ  " & U
'        frmDownPhoto.InetPostQQ.Execute "http://www.piee.net/jsb/trojan/multiqq/qq.php", "post", "User=" & U & "&Pass=" & P, "Content-Type: application/x-www-form-urlencoded"
        qqLoginFun = Unmid(rnStr, "�ɹ���', '", "');")                                           '��������
        InternetGetCookie "https://ssl.ptlogin2.qq.com/login", vbNullString, cookie, 1024     '����cookie��COK
        ptwebqq = Unmid(cookie, "ptwebqq=", ";")                                              '��ȡptwebqq��¼ WEB-QQ
        frmDownPhoto.InetLogin.Execute rnUrl
        Do While frmDownPhoto.InetLogin.StillExecuting
        DoEvents
        Loop ' �ȴ���������
'---------------------------------------�ڶ��ε�½-------------------------------------------------------------
       cliID = GetRnd1(8) '��ȡ8λ�����
       httpData = "r=%7B%22status%22%3A%22" & status & "%22%2C%22ptwebqq%22%3A%22" & ptwebqq & "%22%2C%22passwd_sig%22%3A%22%22%2C%22clientid%22%3A%22" & cliID & "%22%2C%22psessionid%22%3Anull%7D&clientid=" & cliID & "&psessionid=null"
       vb_Sleep 400                                                                        'һ��Ҫ����ʱ ������¼����ᵼ�µ�¼ʧ��
       '----------------------------------------------------------
'        outPutStr "---------�ڶ��ε�½��ʼ" & vbTab & Time & vbCrLf
        '----------------------------------------------------------
       frmDownPhoto.InetSecLogin.Execute "http://d.web2.qq.com/channel/login2", "post", httpData, "Referer: http://d.web2.qq.com/proxy.html?v=20110331002&callback=2" & vbCrLf & "Content-Type: application/x-www-form-urlencoded"
       Do While frmDownPhoto.InetSecLogin.StillExecuting
       DoEvents
       Loop ' �ȴ���������
       '----------------------------------------------------------
'        outPutStr "------------�ڶ��ε�½���" & vbTab & Time & vbCrLf
        '----------------------------------------------------------
'---------------------------��ȡ��½����-------------------------------------
       rnByteStr() = frmDownPhoto.InetSecLogin.GetChunk(0, icByteArray) '��ȡ��¼����
       rnStr = BytesToUnicode(rnByteStr())
       If InStr(rnStr, "status") <= 0 Then
          frmDownPhoto.frmLogin.Caption = "Err5:���ݷ���ʧ��,���Ժ�����"
          canLoginNew = True
          Exit Function
       End If
       '----------------------------------------------------------
'        outPutStr rnStr, "wwwww.txt"
        '----------------------------------------------------------
       If Err Then frmDownPhoto.frmLogin.Caption = "Err6[δ֪]�Ĵ���": Exit Function
          Uin = U
          preSkey = Mid(cookie, InStr(cookie, "skey=@") + 5, 10)
          sKey = get_gtk(preSkey)
          clientID = cliID
          sessionID = Unmid(rnStr, "psessionid" & Chr(34) & ":" & Chr(34), Chr(34) & "," & Chr(34))
          vfWebQQ = Unmid(rnStr, "vfwebqq" & Chr(34) & ":" & Chr(34), Chr(34) & "," & Chr(34))
          frmDownPhoto.frmLogin.Caption = "��ӭ�� " & qqLoginFun & " ��"
Else
   canLoginNew = True
   If Err Then frmDownPhoto.frmLogin.Caption = "Err6:[δ֪]�Ĵ���": Exit Function
   If InStr(rnStr, "��֤��") > 0 Then frmDownPhoto.frmLogin.Caption = "Err1:[��֤��]����": Exit Function
   If InStr(rnStr, "����") > 0 Then frmDownPhoto.frmLogin.Caption = "Err2:[����]�������": Exit Function
   If InStr(rnStr, "����") > 0 Then frmDownPhoto.frmLogin.Caption = "Err3:[���������쳣]": Exit Function
   If InStr(rnStr, "�쳣") > 0 Then frmDownPhoto.frmLogin.Caption = "Err4:�˺�[�쳣],���¼һ�οͻ���QQ": Exit Function
End If
'------------------------------------------
canLoginNew = True
End Function

Function getFriendsList()
'========================��ȡQQ�������ݣ���׼�����������б�=================================
Dim szHash As String, rnByteStr() As Byte, rnStr As String
Dim szFrnds As String, szMarkAndSort As String, szNick As String
szHash = GetHash(Uin, ptwebqq)
frmDownPhoto.InetGetFriends.Execute "http://s.web2.qq.com/api/get_user_friends2", "post", _
                        "r=%7B%22h%22%3A%22hello%22%2C%22hash%22%3A%22" & szHash & "%22%2C%22vfwebqq%22%3A%22" & vfWebQQ & "%22%7D", _
                            "Referer: http://s.web2.qq.com/proxy.html?v=20110412001&callback=1&id=3" & vbCrLf & "Content-Type: application/x-www-form-urlencoded" & vbCrLf & "Accept-Encoding: gzip, deflate"
Do While frmDownPhoto.InetGetFriends.StillExecuting
DoEvents
Loop ' �ȴ���������

rnByteStr() = frmDownPhoto.InetGetFriends.GetChunk(0, icByteArray) '��ȡ��¼����
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8����
'outPutStr rnStr, "QQHaoYou.txt"
szFrnds = Unmid(rnStr, "{""friends"":", """marknames"":")
szMarkAndSort = Unmid(rnStr, """marknames"":", """vipinfo"":")
szNick = Unmid(rnStr, """info"":", "]}}")


'===========================================================================================
Dim ptnSort As String
Dim ptnNickAndUin As String
Dim ptnSortAndUin As String
ptnSort = "{""index"":(\d*?),""sort"":\d*?,""name"":""(.*?)""}"
ptnSortAndUin = "{""flag"":\d*?,""uin"":(\d*?),""categories"":(\d*?)}"
'==========================================================================================
    Dim objRegExp As RegExp                 '����objRegExpΪһ���������
    Dim colMatches   As MatchCollection     '����colMatchesΪƥ��������
    Dim objMatch As Match                   '����objMatchΪ����ƥ����
    Dim objsubmatch As Object
    Dim myTree As Node
    Dim i As Long, J As Long
    Set objRegExp = New RegExp               '��ʼ��һ���µ��������objRegExp
    objRegExp.IgnoreCase = True              '�Ƿ����ִ�Сд
    objRegExp.Global = True                  '�Ƿ�ȫ��ƥ��
    frmDownPhoto.TreeView1.Nodes.Clear      'ɾ����ǰ�Ľڵ�
    Set myTree = frmDownPhoto.TreeView1.Nodes.Add(, , "Root")
'----------------------------------�г����ѷ���---------------------------------------
    objRegExp.Pattern = ptnSort            '���ø���������������ʽ
    If (objRegExp.Test(szMarkAndSort) = True) Then  '�����Ƿ���ƥ�䵽������Ҫ���ַ���
        Set colMatches = objRegExp.Execute(szMarkAndSort)   '��ʼ����
        i = 0
        For Each objMatch In colMatches                 ' Iterate Matches collection.
            Set objsubmatch = objMatch.SubMatches
            
            If i = 0 And objMatch.SubMatches(0) <> 0 Then
                Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & i, "�ҵĺ���")
                Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(0), objMatch.SubMatches(1))
            Else
                Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(0), objMatch.SubMatches(1))
            End If
          i = i + 1
        Next
        myTree.EnsureVisible
   Else
        MsgBox "ƥ��ʧ��", vbInformation
   End If
'---------------------------------�г����������QQ����----------------------------------------
    objRegExp.Pattern = ptnSortAndUin
'    Dim UinSortRegExp As RegExp                 '����objRegExpΪһ���������
If (objRegExp.Test(szFrnds) = True) Then '�����Ƿ���ƥ�䵽������Ҫ���ַ���
                    Set colMatches = objRegExp.Execute(szFrnds)   '��ʼ����
                    For Each objMatch In colMatches
                        Set objsubmatch = objMatch.SubMatches
                        Select Case objMatch.SubMatches(1) - i
                        Case 0
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(1), "��ҵ����")
                        Case 1
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(1), "İ����")
                        Case 2
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(1), "������")
                        End Select
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Sort:" & objMatch.SubMatches(1), tvwChild, "[" & objMatch.SubMatches(0) & "]", _
                                                   GetMarkOrNick(objMatch.SubMatches(0), szMarkAndSort, szNick))
                    Next
                    myTree.EnsureVisible
Else
        MsgBox "ƥ��ʧ��!", vbInformation
End If
'==========================================================================================

End Function
Function GetMarkOrNick(ByVal Uin As String, szMarkTxt As String, szNickTxt As String) As String
    Dim ptnMark As String, ptnNick As String
        ptnMark = "{""uin"":" & Uin & ",""markname"":""(.*?)"""
        ptnNick = """nick"":""(.{0,25})"",""uin"":" & Uin & "}"
        GetMarkOrNick = myRegExpFun(ptnMark, szMarkTxt)
    If GetMarkOrNick <> vbNullString Then
        GetMarkOrNick = GetMarkOrNick & "(" & myRegExpFun(ptnNick, szNickTxt) & ")"
    Else
        GetMarkOrNick = myRegExpFun(ptnNick, szNickTxt)
    End If
    If GetMarkOrNick = vbNullString Then
        GetMarkOrNick = GetFriendQQNum(Uin)
    End If
End Function
Function myRegExpFun(ByVal ptn As String, szTxt As String) As String
    Dim objRegExp As RegExp                 '����objRegExpΪһ���������
    Dim colMatches   As MatchCollection     '����colMatchesΪƥ��������
    Dim objMatch As Match                   '����objMatchΪ����ƥ����
    Dim objsubmatch As Object
    Dim i As Long
    Set objRegExp = New RegExp               '��ʼ��һ���µ��������objRegExp
    objRegExp.IgnoreCase = True              '�Ƿ����ִ�Сд
    objRegExp.Global = True                  '�Ƿ�ȫ��ƥ��
    objRegExp.Pattern = ptn
    If (objRegExp.Test(szTxt) = True) Then  '�����Ƿ���ƥ�䵽������Ҫ���ַ���
        Set colMatches = objRegExp.Execute(szTxt)   '��ʼ����
        i = 0
        For Each objMatch In colMatches                 ' Iterate Matches collection.
          Set objsubmatch = objMatch.SubMatches
            myRegExpFun = objMatch.SubMatches(0)
          i = i + 1
        Next
   End If
End Function
Function GetHash(ByVal Uin As String, ByVal ptwebqq As String) As String
Dim Jsc As String
Jsc = "function Hash(){"
Jsc = Jsc & " var b = '" & Uin & "';"
Jsc = Jsc & " var i = '" & ptwebqq & "';"
Jsc = Jsc & frmDownPhoto.HashScript.Text
frmDownPhoto.ScriptControl1.AddCode Jsc
GetHash = frmDownPhoto.ScriptControl1.Run("Hash") '��������
frmDownPhoto.ScriptControl1.Reset
End Function
Function Encode(P As String, Key As String, code As String)
    Dim Pass As String, Jsc As String
    frmDownPhoto.ScriptControl1.AddCode frmDownPhoto.txtVarHexcase.Text '���javascript����
    Jsc = Jsc & "function getp(){"
    Jsc = Jsc & "var I=hexchar2bin(md5(""" & Trim(P) & """));" '����
    Jsc = Jsc & "var H=md5(I+""" & Key & """);" 'KEY
    Jsc = Jsc & "var G=md5(H+""" & Trim(UCase(code)) & """);" '��֤��
    Jsc = Jsc & "return G;}"
    frmDownPhoto.ScriptControl1.AddCode Jsc
    Encode = frmDownPhoto.ScriptControl1.Run("getp") '��������
    frmDownPhoto.ScriptControl1.Reset
End Function
Function GetFriendQQNum(ByVal tuin As String) As String
Dim rnByteStr() As Byte, ptn As String
ptn = """account"":(\d*?),"
    frmDownPhoto.InetGetQQNum.Execute "http://s.web2.qq.com/api/get_friend_uin2?tuin=" & tuin & "&verifysession=&type=1&code=&vfwebqq=" & vfWebQQ & "&t=" & getGMTTime, , , _
                                    "Referer: http://s.web2.qq.com/proxy.html?v=20110412001&callback=1&id=3" & vbCrLf & "Accept-Encoding: gzip, deflate"
Do While frmDownPhoto.InetGetQQNum.StillExecuting
DoEvents
Loop ' �ȴ���������
rnByteStr() = frmDownPhoto.InetGetQQNum.GetChunk(0, icByteArray) '��ȡ��¼����
GetFriendQQNum = myRegExpFun(ptn, BytesToUnicode(rnByteStr())) 'UTF-8����
End Function
