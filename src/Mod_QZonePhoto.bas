Attribute VB_Name = "Mod_QZonePhoto"
Option Explicit
Function GetAlbumList(ByVal friQQNum As String)
Dim rnByteStr() As Byte, rnStr As String, beDone As Boolean
Dim tolNum As Long, curPageNum As Long, startNum As Long
Dim i As Long
i = 1
startNum = 0
beDone = False
frmDownPhoto.lstAlbumName.Clear '相册名称
frmDownPhoto.lstAlbumID.Clear '相册id
frmDownPhoto.InetGetAlbum.Execute "http://user.qzone.qq.com/" & friQQNum
Do While frmDownPhoto.InetGetAlbum.StillExecuting
DoEvents
Loop ' 等待返回数据
'----------------------------------
GetRoute
'----------------------------------
Do While beDone <> True
ReStart:
'    MsgBox aHost
    frmDownPhoto.InetGetAlbum.Execute "http://" & aHost & "/fcgi-bin/fcg_list_album_v3?g_tk=" & Getg_tkValue & "&callback=shine" & i - 1 & "_Callback&t=" & GetRnd1(9) & "&hostUin=" & friQQNum & "&uin=" & Uin & "&appid=4&inCharset=utf-8&outCharset=utf-8&source=qzone&plat=qzone&format=jsonp&notice=0&filter=1&handset=4&pageNumModeSort=40&pageNumModeClass=10&needUserInfo=1&pageStart=" & startNum & "&pageNum=20&idcNum=" & idcNum_A & "&callbackFun=shine" & i - 1 & "&_=" & getGMTTime, , , _
                                "Referer: http://cnc.qzs.qq.com/qzone/photo/v7/page/photo.html?init=photo.v7/module/albumList/index&navBar=1"
    Do While frmDownPhoto.InetGetAlbum.StillExecuting
    DoEvents
    Loop ' 等待返回数据
    rnByteStr() = frmDownPhoto.InetGetAlbum.GetChunk(0, icByteArray) '获取登录数据
    rnStr = BytesToUnicode(rnByteStr())
'    MsgBox Len(rnStr)
    GetAlbumNum rnStr, tolNum, startNum, curPageNum
    If Len(rnStr) = 0 Then
    MsgBox "操作过频繁，请稍后再试！"
    i = 6
    End If
    If i <= 5 And Len(rnStr) < 200 Then
        Select Case i
                Case 1
                aHost = DOMAIN_0_A
                idcNum_A = "0"
                Case 2
                aHost = DOMAIN_1_A
                idcNum_A = "102"
                Case 3
                aHost = DOMAIN_2_A
                idcNum_A = "2"
                Case 4
                aHost = DOMAIN_3_A
                idcNum_A = "3"
                Case 5
                aHost = DOMAIN_4_A
                idcNum_A = "4"
            End Select
        i = i + 1
        GoTo ReStart
    End If
    If curPageNum < startNum Or startNum = tolNum Then: beDone = True
'    MsgBox frmDownPhoto.InetGetAlbum.GetHeader
    albumListRegExp rnStr
'    outPutStr rnStr, "album.txt"
Loop
End Function

Function GetPhotoList(ByVal friQQNum As String, ByVal albumID As String) As Long
Dim rnByteStr() As Byte, rnStr As String, beDone As Boolean
Dim tolNum As Long, curPageNum As Long, startNum As Long
Dim i As Long
i = 1
beDone = False
startNum = 0
frmDownPhoto.lstPhotoURL.Clear '某个相册的所有相片地址
frmDownPhoto.lstPhotoName.Clear '某个相册的所有相片名称
'MsgBox "http://" & pHost & "/fcgi-bin/cgi_list_photo?g_tk=" & Getg_tkValue & "&callback=shine" & i - 1 & "_Callback&t=" & GetRnd1(9) & "&mode=0&idcNum=" & idcNum_P & "&hostUin=" & friQQNum & "&topicId=" & albumID & "&noTopic=0&uin=" & Uin & "&pageStart=" & startNum & "&pageNum=20&singleurl=1&notice=0&appid=4&inCharset=utf-8&outCharset=utf-8&source=qzone&plat=qzone&outstyle=json&format=jsonp&json_esc=1&question=&answer=&callbackFun=shine0&_=" & getGMTTime
Do While beDone <> True

ReGet:
    frmDownPhoto.InetGetPhoto.Execute "http://" & pHost & "/fcgi-bin/cgi_list_photo?g_tk=" & Getg_tkValue & "&callback=shine" & i - 1 & "_Callback&t=" & GetRnd1(9) & "&mode=0&idcNum=" & idcNum_P & "&hostUin=" & friQQNum & "&topicId=" & albumID & "&noTopic=0&uin=" & Uin & "&pageStart=" & startNum & "&pageNum=20&singleurl=1&notice=0&appid=4&inCharset=utf-8&outCharset=utf-8&source=qzone&plat=qzone&outstyle=json&format=jsonp&json_esc=1&question=&answer=&callbackFun=shine" & i - 1 & "&_=" & getGMTTime, , , _
                                "Referer: http://cnc.qzs.qq.com/qzone/photo/v7/page/photo.html?init=photo.v7/module/photoList2/index&navBar=1&normal=1&aid=" & albumID
    Do While frmDownPhoto.InetGetPhoto.StillExecuting
        DoEvents
    Loop ' 等待返回数据
    rnByteStr() = frmDownPhoto.InetGetPhoto.GetChunk(0, icByteArray) '获取登录数据
    rnStr = BytesToUnicode(rnByteStr())
'    MsgBox Len(rnStr)
        If Len(rnStr) = 0 Then
    MsgBox "操作过频繁，请稍后再试！"
    i = 6
    End If
        If i <= 5 And Len(rnStr) < 200 Then
        Select Case i
                Case 1
                pHost = DOMAIN_0_P
                idcNum_P = "0"
                Case 2
                pHost = DOMAIN_1_P
                idcNum_P = "102"
                Case 3
                pHost = DOMAIN_2_P
                idcNum_P = "2"
                Case 4
                pHost = DOMAIN_3_P
                idcNum_P = "3"
                Case 5
                pHost = DOMAIN_4_P
                idcNum_P = "4"
            End Select
        i = i + 1
        GoTo ReGet
    End If
    GetPhotoNum rnStr, tolNum, curPageNum
    startNum = startNum + curPageNum
    If curPageNum < 20 Or tolNum = 20 Then: beDone = True
    photoListRegExp rnStr, curPageNum
Loop
'MsgBox frmDownPhoto.InetGetPhoto.GetHeader
'Do While beDone <> True
'MsgBox Len(rnStr)
'If Len(rnStr) = 0 Then
'beDone = True
'Else
'frmDownPhoto.txtPhoto.Text = frmDownPhoto.txtPhoto.Text & rnStr
'End If
'Loop
'MsgBox Len(frmDownPhoto.txtPhoto.Text)
'rnByteStr() = frmDownPhoto.InetGetPhoto.GetChunk(65536, icByteArray) '获取登录数据
'rnStr = BytesToUnicode(rnByteStr())
'frmDownPhoto.txtPhoto.Text = BytesToUnicode(rnByteStr())
GetPhotoList = frmDownPhoto.lstPhotoName.ListCount
End Function
Function albumListRegExp(szTxt As String)   'ByVal ptn As String, ByVal szTxt As String) As String
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Dim i As Long, Num As Long
    i = 0
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    objRegExp.Pattern = """id"" : ""(.*?)"",\s*.*\s*.*\s*""name"" : ""(.*?)""," '\s*.*\s*.*\s*.*\s*.*\s*""total"" : (\d*?),"
    If (objRegExp.Test(szTxt) = True) Then  '测试是否能匹配到我们需要的字符串
        Set colMatches = objRegExp.Execute(szTxt)   '开始检索
        For Each objMatch In colMatches                 ' Iterate Matches collection.
          Set objsubmatch = objMatch.SubMatches
            Num = photoNumRegExp(szTxt, i)
            If Num <> "0" Then
            frmDownPhoto.lstAlbumID.AddItem objMatch.SubMatches(0)
            frmDownPhoto.lstAlbumName.AddItem "【" & objMatch.SubMatches(1) & " [" & Num & "]】"
            End If
            i = i + 1
        Next
   End If
End Function
Function photoNumRegExp(szTxt As String, index As Long) As String   'ByVal ptn As String, ByVal szTxt As String) As String
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    objRegExp.Pattern = """total"" : (\d*?),"
    If (objRegExp.Test(szTxt) = True) Then  '测试是否能匹配到我们需要的字符串
        Set colMatches = objRegExp.Execute(szTxt)   '开始检索
            Set objMatch = colMatches.Item(index)
            Set objsubmatch = objMatch.SubMatches
            photoNumRegExp = objMatch.SubMatches(0)
   End If
End Function

Function photoListRegExp(szTxt As String, Num As Long)   'ByVal szTxt As String) As String
    Dim ptnName As String
    Dim ptnUrl As String
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Dim i As Long, J As Long
    i = 1: J = 1
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    ptnName = """name"" : ""(.*?)"""
    ptnUrl = """url"" : ""(.*?)"""
    objRegExp.Pattern = ptnName
    If (objRegExp.Test(szTxt) = True) Then  '测试是否能匹配到我们需要的字符串
        Set colMatches = objRegExp.Execute(szTxt)   '开始检索
        For Each objMatch In colMatches                 ' Iterate Matches collection.
          Set objsubmatch = objMatch.SubMatches
            If i <= Num Then
            frmDownPhoto.lstPhotoName.AddItem objMatch.SubMatches(0)
            End If
            i = i + 1
        Next
   End If
   objRegExp.Pattern = ptnUrl
    If (objRegExp.Test(szTxt) = True) Then  '测试是否能匹配到我们需要的字符串
        Set colMatches = objRegExp.Execute(szTxt)   '开始检索
        For Each objMatch In colMatches                 ' Iterate Matches collection.
          Set objsubmatch = objMatch.SubMatches
            If J <= Num Then
            frmDownPhoto.lstPhotoURL.AddItem urlRelpaceRegExp(objMatch.SubMatches(0))
            End If
            J = J + 1
        Next
   End If
End Function
Function urlRelpaceRegExp(ByVal szTxt As String) As String
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    objRegExp.Pattern = "\\"
    If (objRegExp.Test(szTxt) = True) Then  '测试是否能匹配到我们需要的字符串
        urlRelpaceRegExp = objRegExp.Replace(szTxt, vbNullString)
   End If
End Function
Function Getg_tkValue() As String
Dim Jsc As String
Jsc = Jsc & "function g_tk(){"
Jsc = Jsc & "var str = '" & preSkey & "';"
Jsc = Jsc & "var hash=5381;for(var i=0,len=str.length;i<len;++i)hash+=(hash<<5)+str.charCodeAt(i);return hash&2147483647};"
frmDownPhoto.ScriptControl1.AddCode Jsc
Getg_tkValue = frmDownPhoto.ScriptControl1.Run("g_tk")  '加密密码
frmDownPhoto.ScriptControl1.Reset
End Function
Function GetPhotoNum(szTxt As String, tolNum As Long, curPageNum As Long)
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    objRegExp.Pattern = """totalInAlbum"" : (\d*?),\s*?""totalInPage"" : (\d*?)\s*?}"
    If (objRegExp.Test(szTxt) = True) Then  '测试是否能匹配到我们需要的字符串
        Set colMatches = objRegExp.Execute(szTxt)   '开始检索
        For Each objMatch In colMatches                 ' Iterate Matches collection.
          Set objsubmatch = objMatch.SubMatches
            tolNum = objMatch.SubMatches(0)
            curPageNum = objMatch.SubMatches(1)
        Next
   End If
End Function
Function GetAlbumNum(szTxt As String, tolNum As Long, startNum As Long, pageNum As Long)
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    objRegExp.Pattern = "albumsInUser.*?(\d*?),"
    Set colMatches = objRegExp.Execute(szTxt)   '开始检索
    For Each objMatch In colMatches                 ' Iterate Matches collection.
      Set objsubmatch = objMatch.SubMatches
        tolNum = objMatch.SubMatches(0)
    Next
    objRegExp.Pattern = "nextPageStart.*?(\d*?),"
    Set colMatches = objRegExp.Execute(szTxt)   '开始检索
    For Each objMatch In colMatches                 ' Iterate Matches collection.
      Set objsubmatch = objMatch.SubMatches
        startNum = objMatch.SubMatches(0)
    Next
    objRegExp.Pattern = "totalInPage.*?\s(\d*?)[,\s]"
    Set colMatches = objRegExp.Execute(szTxt)   '开始检索
    For Each objMatch In colMatches                 ' Iterate Matches collection.
      Set objsubmatch = objMatch.SubMatches
        pageNum = objMatch.SubMatches(0)
    Next
End Function

Function GetRoute()
    Dim rnByteStr() As Byte, rnStr As String
    Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
    Dim colMatches   As MatchCollection     '声明colMatches为匹配结果集合
    Dim objMatch As Match                   '声明objMatch为单个匹配结果
    Dim objsubmatch As Object
    Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
    objRegExp.IgnoreCase = True              '是否区分大小写
    objRegExp.Global = True                  '是否全局匹配
    
    frmDownPhoto.InetGetRoute.Execute "http://route.store.qq.com/GetRoute?UIN=" & friQQNum & "&type=json&version=2&json_esc=1&g_tk=" & Getg_tkValue, , , _
                                "Referer: http://cnc.qzs.qq.com/qzone/client/photo/pages/photocanvas.html"
                
    Do While frmDownPhoto.InetGetRoute.StillExecuting
    DoEvents
    Loop ' 等待返回数据
    rnByteStr() = frmDownPhoto.InetGetRoute.GetChunk(0, icByteArray) '获取登录数据
    rnStr = BytesToUnicode(rnByteStr())
    objRegExp.Pattern = """default"":""(.*?)""\s*.*,""\1"":\s*?.*?""p"":""(.*?)"",""s"":""(.*?)"""
    Set colMatches = objRegExp.Execute(rnStr)   '开始检索
    For Each objMatch In colMatches                 ' Iterate Matches collection.
      Set objsubmatch = objMatch.SubMatches
        aHost = objMatch.SubMatches(1)
        Select Case aHost
            Case DOMAIN_0_A
                idcNum_A = "0"
            Case DOMAIN_1_A
                idcNum_A = "102"
            Case DOMAIN_2_A
                idcNum_A = "2"
            Case DOMAIN_3_A
                idcNum_A = "3"
            Case DOMAIN_4_A
                idcNum_A = "4"
        End Select
        pHost = objMatch.SubMatches(2)
    Next
    If aHost = vbNullString Then: aHost = "xalist.photo.qq.com": idcNum_A = "102"
    If pHost = vbNullString Then: pHost = "xaplist.photo.qq.com": idcNum_P = "102"
'    MsgBox aHost & vbCrLf & pHost
    idcNum_P = idcNum_A
End Function

