Attribute VB_Name = "Mod_TVColor"
Option Explicit

Public Const GWL_STYLE As Long = (-16)
Public Const COLOR_WINDOW As Long = 5
Public Const COLOR_WINDOWTEXT As Long = 8

Public Const TVI_ROOT   As Long = &HFFFF0000
Public Const TVI_FIRST  As Long = &HFFFF0001
Public Const TVI_LAST   As Long = &HFFFF0002
Public Const TVI_SORT   As Long = &HFFFF0003

Public Const TVIF_STATE As Long = &H8

Public Const TVS_HASLINES As Long = 2
Public Const TVS_FULLROWSELECT As Long = &H1000

Public Const TVIS_BOLD  As Long = &H10

Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

Public Const TVGN_ROOT                As Long = &H0
Public Const TVGN_NEXT                As Long = &H1
Public Const TVGN_PREVIOUS            As Long = &H2
Public Const TVGN_PARENT              As Long = &H3
Public Const TVGN_CHILD               As Long = &H4
Public Const TVGN_FIRSTVISIBLE        As Long = &H5
Public Const TVGN_NEXTVISIBLE         As Long = &H6
Public Const TVGN_PREVIOUSVISIBLE     As Long = &H7
Public Const TVGN_DROPHILITE          As Long = &H8
Public Const TVGN_CARET               As Long = &H9

Public Type TV_ITEM
   mask As Long
   hItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Public Declare Function GetSysColor Lib "user32" _
   (ByVal nIndex As Long) As Long






Public Function SetTVBackColour(hwndTV As Long, clrref As Long)

'   Dim hwndTV As Long
   Dim style As Long
   
   Call SendMessage(hwndTV, TVM_SETBKCOLOR, 0, ByVal clrref)
   
   style = GetWindowLong(hwndTV, GWL_STYLE)

   If style And TVS_HASLINES Then
      Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
      Call SetWindowLong(hwndTV, GWL_STYLE, style)
   End If
  
End Function


Public Function SetTVForeColour(hwndTV As Long, clrref As Long)
'   Dim hwndTV As Long
   Dim style As Long
   
   Call SendMessage(hwndTV, TVM_SETTEXTCOLOR, 0, ByVal clrref)

   style = GetWindowLong(hwndTV, GWL_STYLE)
   
   If style And TVS_HASLINES Then
      Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
      Call SetWindowLong(hwndTV, GWL_STYLE, style)
   End If
   
End Function

