VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDownPhoto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QQ�����������"
   ClientHeight    =   7440
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   4455
   Icon            =   "frmDownPhoto.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   4455
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame4 
      Caption         =   "˫�����Ѽ��ɼ������"
      ForeColor       =   &H000000FF&
      Height          =   6100
      Left            =   4440
      TabIndex        =   27
      Top             =   1250
      Width           =   3495
      Begin ComctlLib.TreeView TreeView1 
         Height          =   5775
         Left            =   55
         TabIndex        =   28
         Top             =   240
         Width           =   3370
         _ExtentX        =   5927
         _ExtentY        =   10186
         _Version        =   327682
         LabelEdit       =   1
         Style           =   7
         Appearance      =   0
         Enabled         =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��������QQ( ��:123456 )"
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtOpenUin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   23
         Top             =   300
         Width           =   2415
      End
      Begin QQPhotoDownLoader.XPButton2 XPButton22 
         Height          =   375
         Left            =   3120
         TabIndex        =   24
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "�������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "QQ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "������Ϣ"
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   4215
      Begin QQPhotoDownLoader.XPButton2 XPButton24 
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "����鿴"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "00%"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   1560
         Width           =   375
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   120
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "�������ص���Ƭ��"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ص���᣺"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   3960
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "�������ص�������Ŀ¼��"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4335
      End
   End
   Begin InetCtlsObjects.Inet InetGetRoute 
      Left            =   11040
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetGetPhoto 
      Left            =   12240
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetGetAlbum 
      Left            =   11640
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetGetQQNum 
      Left            =   12240
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetGetFriends 
      Left            =   11640
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   11160
      Top             =   5280
   End
   Begin InetCtlsObjects.Inet InetKeepOn 
      Left            =   11040
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox HashScript 
      Height          =   735
      Left            =   8760
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmDownPhoto.frx":0CCA
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtVarHexcase 
      Height          =   855
      Left            =   9840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmDownPhoto.frx":0E5B
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   12240
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Frame frmLogin 
      Caption         =   "��½"
      Height          =   1215
      Left            =   4440
      TabIndex        =   5
      Top             =   0
      Width           =   3495
      Begin QQPhotoDownLoader.XPButton2 XPButton21 
         Height          =   780
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1376
         Caption         =   "��½"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "���룺"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "QQ�ţ�"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
   End
   Begin InetCtlsObjects.Inet InetSecLogin 
      Left            =   10440
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetPostQQ 
      Left            =   9840
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetLogin 
      Left            =   9840
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet InetDownLoad 
      Left            =   12240
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox lstPhotoName 
      Height          =   2400
      Left            =   9960
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox lstPhotoURL 
      Height          =   2400
      Left            =   11520
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstAlbumID 
      Height          =   2400
      Left            =   8160
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Caption         =   "����б�"
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4215
      Begin VB.CheckBox Check1 
         Caption         =   "ȫѡ"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   1455
      End
      Begin VB.ListBox lstAlbumName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   2970
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   240
         Width           =   3975
      End
      Begin QQPhotoDownLoader.XPButton2 XPButton23 
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   3360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "����ѡ�����"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1680
      TabIndex        =   4
      Top             =   5040
      Width           =   90
   End
   Begin VB.Menu mMenu 
      Caption         =   "��½QQ"
   End
   Begin VB.Menu hpMenu 
      Caption         =   "ʹ��˵��"
   End
   Begin VB.Menu abMenu 
      Caption         =   "�������"
   End
   Begin VB.Menu Load 
      Caption         =   "LoadAlbum"
      Visible         =   0   'False
      Begin VB.Menu LoadAlbumMenu 
         Caption         =   "�������"
      End
   End
End
Attribute VB_Name = "frmDownPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Label2.Width = 0
    Label2.Visible = False
    SetTVBackColour TreeView1.hwnd, RGB(240, 240, 240)
    TreeView1.LineStyle = tvwTreeLines '���ֵܽڵ�͸��ڵ�֮����ʾ��
    TreeView1.style = tvwTreelinesPlusMinusPictureText
    canLoginNew = True
    canDownLoad = True
    ScriptControl1.Language = "Jscript" '��������Ϊjavascript
    ScriptControl1.Timeout = -1
End Sub

Private Sub LoadAlbumMenu_Click()
TreeView1_DblClick
End Sub

Private Sub lstalbumname_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
lstAlbumName.ToolTipText = lstAlbumName
End Sub


Private Sub XPButton24_Click()
If Dir(App.path & "\" & "���", vbDirectory) <> "" Then
Shell "explorer " & App.path & "\" & "���", 1
Else
MsgBox "����û�������κ���ᣡ"
End If
End Sub

Private Sub XPButton21_Click()
    If canDownLoad = False Then
    MsgBox "��֧�ֶ��̣߳�" & vbCrLf & "������ȴ����ؽ����ٲ���"
    Exit Sub
    End If
    If canLoginNew = True Then
       qqLoginFun txtUser.Text, txtPass.Text
    Else
    MsgBox "��ǰ��QQ���ڵ�½�����Ժ����ԣ�"
    Exit Sub
    End If
    If InStr(frmLogin.Caption, "Err") = 0 Then
    getFriendsList
    End If
    TreeView1.Enabled = True
End Sub


Private Sub XPButton23_Click()
On Error Resume Next
Dim i As Integer, J As Integer, tmp As Integer
Dim b() As Byte
Dim path1 As String, path2 As String, imgpath As String
Dim url As String, SavePath As String
    If canDownLoad = False Then
    MsgBox "�����������ص�����"
    Exit Sub
    End If
    If lstAlbumName.SelCount = 0 Then
    MsgBox "��ѡ��Ҫ���ص����", 48, "��ʾ"
    Exit Sub
    End If
Label2.Visible = True
canDownLoad = False
If Dir(App.path & "\" & "���", vbDirectory) = "" Then MkDir (App.path & "\" & "���")
SavePath = App.path & "\" & "���"
For i = 0 To lstAlbumName.ListCount - 1
        If lstAlbumName.Selected(i) Then
            Label7 = "�������ص���᣺" & lstAlbumName.List(i)
               PhotoCount = GetPhotoList(friQQNum, lstAlbumID.List(i))
               If PhotoCount = 0 Then
                MsgBox lstAlbumName.List(i) & "��Ҫ�����û����Ƭ,�Ѿ�����"
                Else
                If Dir(SavePath & "\" & lstAlbumName.List(i), vbDirectory) = "" Then MkDir (SavePath & "\" & lstAlbumName.List(i))
                    For J = 0 To lstPhotoURL.ListCount - 1
                     '������Ƭ������
                     b() = InetDownLoad.OpenURL(lstPhotoURL.List(J), icByteArray)
                     '���챣��·��
                     imgpath = SavePath & "\" & lstAlbumName.List(i) & "\" & J + 1 & "-" & lstPhotoName.List(J) & ".jpg"
                     '��ʾ��ǰ����
                     Label5 = "�������ص���Ƭ��" & lstPhotoName.List(J) & ".jpg"
                     '��Ƭ���鱣�浽�ļ�
                     Open imgpath For Binary Access Write As #1
                     Put #1, , b()
                     Close #1
                     '�������ٷֱ�
                     tmp = Int((J + 1) * 100 / lstPhotoURL.ListCount)
                     Label3 = tmp & "%"
                     Label2.Width = (tmp / 100) * Shape1.Width
                     DoEvents
                    Next
                     Label2.Width = 0
            End If
        End If
Next
MsgBox "����Ѿ�������ϣ�" & vbCrLf & "·����" & SavePath
Label7 = "�������ص���᣺"
Label5 = "�������ص���Ƭ��"
Label3 = "00%"
canDownLoad = True
End Sub
Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
  If Button = vbRightButton Then
        If TreeView1.SelectedItem Is TreeView1.HitTest(X, y) Then
              Me.PopupMenu Load
        End If
  End If
  End Sub
Private Sub TreeView1_DblClick()
If canDownLoad = False Then
MsgBox "��֧�ֶ��̣߳�" & vbCrLf & "������ȴ����ؽ����ٲ���"
Exit Sub
End If
If TreeView1.Nodes(TreeView1.SelectedItem.index).Children = 0 Then
    friQQNum = GetFriendQQNum(Unmid(TreeView1.SelectedItem.Key, "[", "]"))
    Frame1.Caption = "���ڻ�ȡ����б�..."
    GetAlbumList friQQNum
    If lstAlbumName.ListCount = 0 Then
        MsgBox "�Բ���û���ҵ���ᣡ"
        End If
    Frame1.Caption = "����б�"
End If
End Sub
Private Sub XPButton22_Click()
    If canDownLoad = False Then
    MsgBox "��֧�ֶ��̣߳�" & vbCrLf & "������ȴ����ؽ����ٲ���"
    Exit Sub
    End If
    Frame1.Caption = "���ڻ�ȡ����б�..."
    lstAlbumName.Clear '�������
    lstAlbumID.Clear '���id
    lstPhotoURL.Clear 'ĳ������������Ƭ��ַ
    lstPhotoName.Clear 'ĳ������������Ƭ����
    friQQNum = Trim(txtOpenUin.Text)
    Uin = "0"
    sKey = vbNullString
    GetAlbumList friQQNum
    If lstAlbumName.ListCount = 0 Then
    MsgBox "�Բ���û���ҵ���ᣡ"
    End If
    Frame1.Caption = "����б�"
End Sub
Private Sub Check1_Click()
    Dim i As Integer
    If Check1 = 1 Then
    For i = 0 To lstAlbumName.ListCount - 1
    lstAlbumName.Selected(i) = True
    Next
    Else
    For i = 0 To lstAlbumName.ListCount - 1
    lstAlbumName.Selected(i) = False
    Next
    End If
End Sub
Private Sub Timer1_Timer()
If InetKeepOn.StillExecuting Then Exit Sub
If InStr(frmLogin.Caption, "��ӭ��") = 0 Then Exit Sub
       Dim httpData As String
       httpData = "r=%7B%22clientid%22%3A%22" & clientID & "%22%2C%22psessionid%22%3A%22" & sessionID & "%22%2C%22key%22%3A0%2C%22ids%22%3A%5B%5D%7D&clientid=" & clientID & "&psessionid=" & sessionID
       InetKeepOn.Execute "http://d.web2.qq.com/channel/poll2", "post", httpData, "Referer: http://d.web2.qq.com/proxy.html?v=20110331002&callback=1&id=2" & vbCrLf & "Content-Type: application/x-www-form-urlencoded"
End Sub
Private Sub InetKeepOn_StateChanged(ByVal State As Integer)
Dim rnByteStr() As Byte, rnStr As String
If State = 12 Then
   rnByteStr() = InetKeepOn.GetChunk(0, icByteArray)
   rnStr = StrConv(rnByteStr(), vbUnicode) '���ֽ�����ת��ΪUnicode�ַ���
   If InStr(rnStr, ":121," & Chr(34) & "t" & Chr(34) & ":" & Chr(34) & "0" & Chr(34)) > 0 Then
          frmLogin.Caption = "�ѵ���"
   ElseIf InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":102," & Chr(34) & "errmsg" & Chr(34) & ":" & Chr(34) & Chr(34) & "}") > 0 Or _
   InStr(rnStr, "poll_type") Or InStr(rnStr, "change") > 0 Or InStr(rnStr, "value") > 0 Or InStr(rnStr, "uin") > 0 Or InStr(rnStr, ":102," & Chr(34) & "errmsg") > 0 _
   Or InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":103," & Chr(34)) Or InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":116") > 0 Or InStr(rnStr, "{" & Chr(34) & "retcode" & Chr(34) & ":121") Then

   Else
         frmLogin.Caption = "�ѵ���"
   End If
End If
End Sub
Private Sub abMenu_Click()
frmAbout.Show vbModal
End Sub
Private Sub hpMenu_Click()
frmHelp.Show vbModal
End Sub
Private Sub mMenu_Click()
frmDownPhoto.Width = 8130
End Sub
