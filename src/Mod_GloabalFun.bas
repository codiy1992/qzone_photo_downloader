Attribute VB_Name = "Mod_GloabalFun"
Option Explicit
Public Declare Function InternetGetCookie Lib "wininet.dll" Alias "InternetGetCookieA" (ByVal lpszUrlName As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String, lpdwSize As Long) As Boolean
Public Declare Function MultiByteToWideChar Lib "KERNEL32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function DeleteFile Lib "KERNEL32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const CP_UTF8 = 65001

'--------------�����������ص�����------------------
Public canDownLoad As Boolean
'--------------����������ص�����------------------
Public idcNum_A As String
Public idcNum_P As String
Public AlbumCount As Integer
Public PhotoCount As Integer
Public friQQNum     As String  '��ǰѡ�еĺ���QQ��
Public aHost        As String   '������ڷ�����
Public pHost        As String   '��Ƭ���ڷ�����

Public Const DOMAIN_0_A = "alist.photo.qq.com"
Public Const DOMAIN_1_A = "xalist.photo.qq.com"
Public Const DOMAIN_2_A = "hzalist.photo.qq.com"
Public Const DOMAIN_3_A = "gzalist.photo.qq.com"
Public Const DOMAIN_4_A = "shalist.photo.qq.com"
Public Const DOMAIN_0_P = "plist.photo.qq.com"
Public Const DOMAIN_1_P = "xaplist.photo.qq.com"
Public Const DOMAIN_2_P = "hzplist.photo.qq.com"
Public Const DOMAIN_3_P = "gzplist.photo.qq.com"
Public Const DOMAIN_4_P = "shplist.photo.qq.com"
'--------------���½��ص�����------------------
Public User            As Long      'Dialog  ���ڶ�ȡ��֤��QQ����
Public Yzmcode         As String    ' Dialog  ���ڱ�����֤��
Public Pdunload        As Boolean   'Dialog  �����Ƿ�����
Public canLoginNew     As Boolean
'--------------��WebQQ������ص��߸�Ȩ��ֵ-----------
Public Uin          As String       '��QQ�����йصģ����������ظ��ͻ��˵���QQ����ʹ�õ�һ��ֵ
Public Hash         As String       '���ȡ�����б���ص�16λ��Hashֵ
Public preSkey      As String       'ԭ��skey
Public sKey         As String       '��g_tk�㷨������
Public ptwebqq      As String       'cook��ȡ��¼��Ϣ
Public vfWebQQ      As String
Public clientID     As String
Public sessionID    As String
'-----------------�����ģ��"clswaitabletimer"ʹ��-------------
Public mobjWaitTimer As clswaitabletimer
Function GetRnd(ByVal n As Integer) As String '��ȡNλ�����
Randomize
Const Cstring As String = "1234567890"
GetRnd = Mid("0" & Rnd(1) & Cstring, 1, n)
End Function
Function GetRnd1(ByVal n As Integer)  '��ȡNλ�����
Dim X
Randomize
GetRnd1 = Int(8 * Rnd + 1)
For X = 2 To n
    Randomize
    GetRnd1 = GetRnd1 & Int(9 * Rnd + 0)
Next
End Function
Function getGMTTime() As String
frmDownPhoto.ScriptControl1.AddCode "function getGTM(){return (new Date).getTime();}"
getGMTTime = frmDownPhoto.ScriptControl1.Run("getGTM")
frmDownPhoto.ScriptControl1.Reset
End Function
'GetHash�������ڻ�ȡ��ȡQQ�������������Hashֵ

Function get_gtk(sk_ey As String)
    Dim js(6) As String
    js(0) = "function getGTK(str){" & vbCrLf
    js(1) = "var hash = 5381;" & vbCrLf
    js(2) = "for(var i = 0, len = str.length; i < len; ++i){" & vbCrLf
    js(3) = "    hash += (hash << 5) + str.charAt(i).charCodeAt();" & vbCrLf
    js(4) = "}" & vbCrLf
    js(5) = " return hash & 0x7fffffff;" & vbCrLf
    js(6) = "}"
  
    frmDownPhoto.ScriptControl1.AddCode js(0) & js(1) & js(2) & js(3) & js(4) & js(5) & js(6)
    get_gtk = frmDownPhoto.ScriptControl1.Run("getGTK", sk_ey)
    frmDownPhoto.ScriptControl1.Reset
End Function

'inet�����ȡ������ʱ��������룬��Ҫ�ô˺���ת��
Function BytesToUnicode(ByRef Utf() As Byte) As String
    Dim lRet As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    lLength = UBound(Utf) - LBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    BytesToUnicode = String$(lBufferSize, Chr(0))
    lRet = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(BytesToUnicode), lBufferSize)
    If lRet <> 0 Then
        BytesToUnicode = Left(BytesToUnicode, lRet)
    Else
        BytesToUnicode = ""
    End If
End Function
Function Unmid(StrU, Minstr, Maxstr) 'ȡ�м��ı�
'If InStr(StrU, Minstr) > 0 And InStr(StrU, Maxstr) > 0 Then
   Dim q1 As Long, q2 As Long
   q1 = InStr(StrU, Minstr) + Len(Minstr)
   q2 = InStr(q1, StrU, Maxstr)
   'Debug.Print "q2=" & q2
   If q2 = 0 Then Unmid = Replace(StrU, Left(StrU, q1), ""): Exit Function
   Unmid = Mid(StrU, q1, q2 - q1)
'Else
'   Unmid = 0
'End If
End Function
Function GetTimerc() 'ȡʱ���
Dim cs As Date, xs As Date, t As Long
cs = CDate(Now)
xs = CDate("1970-01-01 08:00:00")
Randomize
GetTimerc = DateDiff("s", xs, cs) * 1000 + Int(1 * Rnd + 999)
End Function



'--------------------------------------------------------------
'**************�����ģ��"clswaitabletimer"ʹ��'***************
'������vb_Sleep ��ʹ���ں˶���WaitableTimerʵ�֡�
'������(ʱ��) ��λΪ������
'����ֵ���޷���ֵ
'--------------------------------------------------------------
Public Function vb_Sleep(dwMilliseconds As Long)
    Set mobjWaitTimer = New clswaitabletimer
            mobjWaitTimer.Wait (dwMilliseconds)
    Set mobjWaitTimer = Nothing
End Function

Function outPutStr(ByVal str As String, Optional ByVal path As String = "log.txt")
    Open path For Append As #1
    Print #1, , str
    Close #1
End Function

Public Function Encript(ByVal strValue As String) As String
    Dim byteValue() As Byte
    Dim i As Integer
    byteValue = StrConv(strValue, vbFromUnicode)  '�ַ���ת������
    For i = 0 To UBound(byteValue)
    byteValue(i) = (byteValue(i) + 5) Mod 255
    Next
    Encript = BytesToUnicode(byteValue())
End Function
Public Function Decript(ByVal strValue As String) As String
    Dim byteValue() As Byte
    Dim i As Integer
    byteValue = StrConv(strValue, vbFromUnicode)  '�ַ���ת������
    For i = 0 To UBound(byteValue)
    byteValue(i) = (byteValue(i) - 5) Mod 255
    Next
    Decript = BytesToUnicode(byteValue())
End Function


