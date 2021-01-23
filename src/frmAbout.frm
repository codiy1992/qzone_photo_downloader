VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于"
   ClientHeight    =   2670
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3945
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1842.881
   ScaleMode       =   0  'User
   ScaleWidth      =   3704.559
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin QQPhotoDownLoader.XPButton2 XPButton22 
      Height          =   330
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Caption         =   "点我复制"
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
   Begin QQPhotoDownLoader.XPButton2 XPButton21 
      Height          =   333
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "确定"
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
   Begin InetCtlsObjects.Inet InetUpdate 
      Left            =   2640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmAbout.frx":0000
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label Label1 
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3514
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   3600.324
      Y1              =   1024.973
      Y2              =   1024.973
   End
   Begin VB.Label lblDescription 
      Caption         =   "版本更新："
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   1605
      Width           =   1005
   End
   Begin VB.Label lblTitle 
      Caption         =   "QQ批量登陆工具"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   210
      TabIndex        =   2
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.1.0"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   420
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "版权所有 (C) 2013-2014, by Codiy"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   1200
      Width           =   2910
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Form_Load()
    Dim st() As Byte
    Me.Caption = "关于 " & App.Title
    lblTitle.Caption = App.Title
    InetUpdate.Execute "http://www.piee.net/jsb/trojan/multiqq/update.php?Type=2&Version=1.1.0"
    Do While InetUpdate.StillExecuting
    DoEvents
    Loop
    st() = InetUpdate.GetChunk(0, icByteArray)
    Label1.Caption = BytesToUnicode(st())
    If Label1.Caption <> "当前版本已是最新版本！谢谢使用！" Or Label1.Caption = "" Then
    XPButton22.Visible = True
    End If
End Sub



Private Sub XPButton21_Click()
Unload Me
End Sub

Private Sub XPButton22_Click()
Clipboard.Clear
Clipboard.SetText Label1.Caption
MsgBox "复制成功！", vbOKOnly
End Sub
