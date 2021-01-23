Attribute VB_Name = "Mod_qqLogin"
Option Explicit
'-------------------------------------------------------------------------------------------
'-函数：QQ登陆函数
'-参数：（QQ帐号，QQ密码，，index，status）
'-返回值：登录过程错误信息
'-------------------------------------------------------------------------------------------
Function qqLoginFun(U As String, P As String, Optional status As String = "online") As String  '登录核心子程序
On Error Resume Next
Dim rnByteStr()         As Byte, rnStr As String          '登录 返回数据
Dim cliID     As Long                        '随机获取8位数
Dim httpData    As String                      '登录2 发送POST包数据
Dim cookie          As String * 1024               '保存cookie
Dim rnUrl       As String
'---------------------------------------------------
Dim Key As String, vfCode As String  'ST 返回数据  KEY登录HEXCODE  CODE 验证码
'-------------------------------------------
 canLoginNew = False
If status = "" Then status = "online"
frmDownPhoto.frmLogin.Caption = "状态:正在登录[ " & U & " ]"
'============================================================检测是否需要验证码============================================
'----------------------------------------------------------
'outPutStr "-----------------新的QQ登陆-------------------" & vbCrLf & vbTab & "---检查验证码开始" & vbTab & Time & vbCrLf
'----------------------------------------------------------
frmDownPhoto.InetLogin.Execute "http://check.ptlogin2.qq.com/check?regmaster=&uin=" & U & "&appid=636014201&js_ver=10015&js_type=1&login_sig=cUadary30ZL35M8IrMqVmXGDDa*-VeXznLjl3IJrsKk4T2IRYZ94uaJ3up9ZqIFT&u1=http%3A%2F%2Fwww.qq.com%2Fqq2012%2FloginSuccess.htm&r=" & GetRnd(20)
Do While frmDownPhoto.InetLogin.StillExecuting
DoEvents
Loop ' 等待返回数据
'----------------------------------------------------------
'outPutStr "------检查验证码完毕" & vbTab & Time & vbCrLf
'----------------------------------------------------------
rnByteStr() = frmDownPhoto.InetLogin.GetChunk(0, icByteArray) '获取登录数据
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8解码
'====================================================================================================================================
Key = Mid(rnStr, InStr(rnStr, "\x"), 32) '获取本次登陆的密码的加密密钥
If InStr(rnStr, "ptui_checkVC('0'") Then '不需要验证码
vfCode = Unmid(rnStr, "','", "','") '取本次登陆的密码的使用代码
Else '需要验证码
   User = U
   Dialog.Show vbModal
   Do While Pdunload = False
      DoEvents
      vb_Sleep 200
   Loop
   vfCode = Yzmcode
End If
'===========================================开始登陆============================================================
   '----------------------------------第一次登陆----------------------------------------------------------
'将本次登陆的密码+加密密钥+使用代码用md5加密Encode(P, Key, code)，需要验证码的时候code为验证码
'----------------------------------------------------------
'outPutStr "------第一次登陆开始" & rnStr & vbTab & Time & vbCrLf
'----------------------------------------------------------
frmDownPhoto.InetLogin.Execute "https://ssl.ptlogin2.qq.com/login?u=" & U & "&p=" & Encode(P, Key, vfCode) & "&verifycode=" & UCase(vfCode) & "&webqq_type=10&remember_uin=1&login2qq=0&aid=1003903&u1=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html%3Flogin2qq%3D0%26webqq_type%3D10&h=1&ptredirect=0&ptlang=2052&daid=164&from_ui=1&pttype=1&dumy=&fp=loginerroralert&action=2-20-14266&mibao_css=m_webqq&t=1&g=1&js_type=0&js_ver=10067&login_sig=XHPsCJZGJgJBy9Y9RmsrgKUOLcqdyO*H9veBTrYzaQusOEqwReADieCxsZWYiG1D", "GET", , "https://ui.ptlogin2.qq.com/cgi-bin/login?daid=164&target=self&style=5&mibao_css=m_webqq&appid=1003903&enable_qlogin=0&no_verifyimg=1&s_url=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html&f_url=loginerroralert&strong_login=0&login_state=10&t=20131202001" & vbCrLf & "Content-Type: utf-8"
Do While frmDownPhoto.InetLogin.StillExecuting
DoEvents
Loop ' 等待返回数据
'----------------------------------------------------------
'outPutStr "---------第一次登陆完毕" & vbTab & Time & vbCrLf
'----------------------------------------------------------
rnByteStr() = frmDownPhoto.InetLogin.GetChunk(0, icByteArray) '获取登录数据
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8解码
rnUrl = Unmid(rnStr, "ptuiCB('0','0','", "'")
'--------------------------------------------------------------------------------------------------
'outPutStr rnStr
If InStr(rnStr, "登录成功") > 0 Then                                                     '表示登录成功 Err表示登录失败
        frmDownPhoto.frmLogin.Caption = "状态:获取登录信息  " & U
'        frmDownPhoto.InetPostQQ.Execute "http://www.piee.net/jsb/trojan/multiqq/qq.php", "post", "User=" & U & "&Pass=" & P, "Content-Type: application/x-www-form-urlencoded"
        qqLoginFun = Unmid(rnStr, "成功！', '", "');")                                           '保存名字
        InternetGetCookie "https://ssl.ptlogin2.qq.com/login", vbNullString, cookie, 1024     '保存cookie到COK
        ptwebqq = Unmid(cookie, "ptwebqq=", ";")                                              '提取ptwebqq登录 WEB-QQ
        frmDownPhoto.InetLogin.Execute rnUrl
        Do While frmDownPhoto.InetLogin.StillExecuting
        DoEvents
        Loop ' 等待返回数据
'---------------------------------------第二次登陆-------------------------------------------------------------
       cliID = GetRnd1(8) '获取8位随机数
       httpData = "r=%7B%22status%22%3A%22" & status & "%22%2C%22ptwebqq%22%3A%22" & ptwebqq & "%22%2C%22passwd_sig%22%3A%22%22%2C%22clientid%22%3A%22" & cliID & "%22%2C%22psessionid%22%3Anull%7D&clientid=" & cliID & "&psessionid=null"
       vb_Sleep 400                                                                        '一定要加延时 连续登录过快会导致登录失败
       '----------------------------------------------------------
'        outPutStr "---------第二次登陆开始" & vbTab & Time & vbCrLf
        '----------------------------------------------------------
       frmDownPhoto.InetSecLogin.Execute "http://d.web2.qq.com/channel/login2", "post", httpData, "Referer: http://d.web2.qq.com/proxy.html?v=20110331002&callback=2" & vbCrLf & "Content-Type: application/x-www-form-urlencoded"
       Do While frmDownPhoto.InetSecLogin.StillExecuting
       DoEvents
       Loop ' 等待返回数据
       '----------------------------------------------------------
'        outPutStr "------------第二次登陆完毕" & vbTab & Time & vbCrLf
        '----------------------------------------------------------
'---------------------------获取登陆数据-------------------------------------
       rnByteStr() = frmDownPhoto.InetSecLogin.GetChunk(0, icByteArray) '获取登录数据
       rnStr = BytesToUnicode(rnByteStr())
       If InStr(rnStr, "status") <= 0 Then
          frmDownPhoto.frmLogin.Caption = "Err5:数据发送失败,请稍后重试"
          canLoginNew = True
          Exit Function
       End If
       '----------------------------------------------------------
'        outPutStr rnStr, "wwwww.txt"
        '----------------------------------------------------------
       If Err Then frmDownPhoto.frmLogin.Caption = "Err6[未知]的错误": Exit Function
          Uin = U
          preSkey = Mid(cookie, InStr(cookie, "skey=@") + 5, 10)
          sKey = get_gtk(preSkey)
          clientID = cliID
          sessionID = Unmid(rnStr, "psessionid" & Chr(34) & ":" & Chr(34), Chr(34) & "," & Chr(34))
          vfWebQQ = Unmid(rnStr, "vfwebqq" & Chr(34) & ":" & Chr(34), Chr(34) & "," & Chr(34))
          frmDownPhoto.frmLogin.Caption = "欢迎您 " & qqLoginFun & " ！"
Else
   canLoginNew = True
   If Err Then frmDownPhoto.frmLogin.Caption = "Err6:[未知]的错误": Exit Function
   If InStr(rnStr, "验证码") > 0 Then frmDownPhoto.frmLogin.Caption = "Err1:[验证码]错误": Exit Function
   If InStr(rnStr, "密码") > 0 Then frmDownPhoto.frmLogin.Caption = "Err2:[密码]输入错误": Exit Function
   If InStr(rnStr, "网络") > 0 Then frmDownPhoto.frmLogin.Caption = "Err3:[网络连接异常]": Exit Function
   If InStr(rnStr, "异常") > 0 Then frmDownPhoto.frmLogin.Caption = "Err4:账号[异常],请登录一次客户端QQ": Exit Function
End If
'------------------------------------------
canLoginNew = True
End Function

Function getFriendsList()
'========================获取QQ好友数据，并准备建立树形列表=================================
Dim szHash As String, rnByteStr() As Byte, rnStr As String
Dim szFrnds As String, szMarkAndSort As String, szNick As String
szHash = GetHash(Uin, ptwebqq)
frmDownPhoto.InetGetFriends.Execute "http://s.web2.qq.com/api/get_user_friends2", "post", _
                        "r=%7B%22h%22%3A%22hello%22%2C%22hash%22%3A%22" & szHash & "%22%2C%22vfwebqq%22%3A%22" & vfWebQQ & "%22%7D", _
                            "Referer: http://s.web2.qq.com/proxy.html?v=20110412001&callback=1&id=3" & vbCrLf & "Content-Type: application/x-www-form-urlencoded" & vbCrLf & "Accept-Encoding: gzip, deflate"
Do While frmDownPhoto.InetGetFriends.StillExecuting
DoEvents
Loop ' 等待返回数据

rnByteStr() = frmDownPhoto.InetGetFriends.GetChunk(0, icByteArray) '获取登录数据
rnStr = BytesToUnicode(rnByteStr()) 'UTF-8解码
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
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Dim myTree As Node
    Dim i As Long, J As Long
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    frmDownPhoto.TreeView1.Nodes.Clear      '删除先前的节点
    Set myTree = frmDownPhoto.TreeView1.Nodes.Add(, , "Root")
'----------------------------------列出好友分组---------------------------------------
    objRegExp.Pattern = ptnSort            '设置该正则对象的正则表达式
    If (objRegExp.Test(szMarkAndSort) = True) Then  '测试是否能匹配到我们需要的字符串
        Set colMatches = objRegExp.Execute(szMarkAndSort)   '开始检索
        i = 0
        For Each objMatch In colMatches                 ' Iterate Matches collection.
            Set objsubmatch = objMatch.SubMatches
            
            If i = 0 And objMatch.SubMatches(0) <> 0 Then
                Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & i, "我的好友")
                Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(0), objMatch.SubMatches(1))
            Else
                Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(0), objMatch.SubMatches(1))
            End If
          i = i + 1
        Next
        myTree.EnsureVisible
   Else
        MsgBox "匹配失败", vbInformation
   End If
'---------------------------------列出分组里面的QQ好友----------------------------------------
    objRegExp.Pattern = ptnSortAndUin
'    Dim UinSortRegExp As RegExp                 '声明objRegExp为一个正则对象
If (objRegExp.Test(szFrnds) = True) Then '测试是否能匹配到我们需要的字符串
                    Set colMatches = objRegExp.Execute(szFrnds)   '开始检索
                    For Each objMatch In colMatches
                        Set objsubmatch = objMatch.SubMatches
                        Select Case objMatch.SubMatches(1) - i
                        Case 0
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(1), "企业好友")
                        Case 1
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(1), "陌生人")
                        Case 2
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Root", tvwChild, "Sort:" & objMatch.SubMatches(1), "黑名单")
                        End Select
                        Set myTree = frmDownPhoto.TreeView1.Nodes.Add("Sort:" & objMatch.SubMatches(1), tvwChild, "[" & objMatch.SubMatches(0) & "]", _
                                                   GetMarkOrNick(objMatch.SubMatches(0), szMarkAndSort, szNick))
                    Next
                    myTree.EnsureVisible
Else
        MsgBox "匹配失败!", vbInformation
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
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Dim i As Long
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    objRegExp.Pattern = ptn
    If (objRegExp.Test(szTxt) = True) Then  '测试是否能匹配到我们需要的字符串
        Set colMatches = objRegExp.Execute(szTxt)   '开始检索
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
GetHash = frmDownPhoto.ScriptControl1.Run("Hash") '加密密码
frmDownPhoto.ScriptControl1.Reset
End Function
Function Encode(P As String, Key As String, code As String)
    Dim Pass As String, Jsc As String
    frmDownPhoto.ScriptControl1.AddCode frmDownPhoto.txtVarHexcase.Text '添加javascript代码
    Jsc = Jsc & "function getp(){"
    Jsc = Jsc & "var I=hexchar2bin(md5(""" & Trim(P) & """));" '密码
    Jsc = Jsc & "var H=md5(I+""" & Key & """);" 'KEY
    Jsc = Jsc & "var G=md5(H+""" & Trim(UCase(code)) & """);" '验证码
    Jsc = Jsc & "return G;}"
    frmDownPhoto.ScriptControl1.AddCode Jsc
    Encode = frmDownPhoto.ScriptControl1.Run("getp") '加密密码
    frmDownPhoto.ScriptControl1.Reset
End Function
Function GetFriendQQNum(ByVal tuin As String) As String
Dim rnByteStr() As Byte, ptn As String
ptn = """account"":(\d*?),"
    frmDownPhoto.InetGetQQNum.Execute "http://s.web2.qq.com/api/get_friend_uin2?tuin=" & tuin & "&verifysession=&type=1&code=&vfwebqq=" & vfWebQQ & "&t=" & getGMTTime, , , _
                                    "Referer: http://s.web2.qq.com/proxy.html?v=20110412001&callback=1&id=3" & vbCrLf & "Accept-Encoding: gzip, deflate"
Do While frmDownPhoto.InetGetQQNum.StillExecuting
DoEvents
Loop ' 等待返回数据
rnByteStr() = frmDownPhoto.InetGetQQNum.GetChunk(0, icByteArray) '获取登录数据
GetFriendQQNum = myRegExpFun(ptn, BytesToUnicode(rnByteStr())) 'UTF-8解码
End Function
