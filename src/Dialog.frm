VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÑéÖ¤Âë"
   ClientHeight    =   1275
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      Begin QQPhotoDownLoader.XPButton2 XPButton21 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InetCtlsObjects.Inet InetDownLoad 
         Left            =   600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   120
         MouseIcon       =   "Dialog.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   795
         ScaleWidth      =   1710
         TabIndex        =   2
         Top             =   240
         Width           =   1740
      End
      Begin VB.TextBox CodeY 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CodeY_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   XPButton21_Click
End If
End Sub

Private Sub Form_Load()
Dim aryBytes() As Byte
Me.Move frmDownPhoto.Left + frmDownPhoto.Width / 2 - frmDownPhoto.Width / 2, frmDownPhoto.Top + frmDownPhoto.Height / 2 - Me.Height / 2
InetDownLoad.Execute "http://captcha.qq.com/getimage?uin=" & User & "&aid=636014201&" & GetRnd(19)
Do While InetDownLoad.StillExecuting
DoEvents
Loop
aryBytes() = InetDownLoad.GetChunk(0, icByteArray)
Open App.path & "\tmp.jpg" For Binary As #1
Put #1, , aryBytes()
Close #1
SetAttr App.path & "\tmp.jpg", vbHidden
Picture1.Picture = LoadPicture(App.path & "\tmp.jpg")
DeleteFile App.path & "\tmp.jpg"
Pdunload = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Pdunload = True
Yzmcode = CodeY.Text
End Sub

Private Sub Picture1_Click()
Dim aryBytes() As Byte
InetDownLoad.Execute "http://captcha.qq.com/getimage?uin=" & User & "&aid=636014201&" & GetRnd(19)
Do While InetDownLoad.StillExecuting
DoEvents
Loop
aryBytes() = InetDownLoad.GetChunk(0, icByteArray)
Open App.path & "\tmp.jpg" For Binary As #1
Put #1, , aryBytes()
Close #1
SetAttr App.path & "\tmp.jpg", vbHidden
Picture1.Picture = LoadPicture(App.path & "\tmp.jpg")
DeleteFile App.path & "\tmp.jpg"
End Sub

Private Sub XPButton21_Click()
Yzmcode = CodeY.Text
Unload Me
End Sub
