Attribute VB_Name = "modLV"
Option Explicit
Public Const LVS_SHAREIMAGELISTS = &H40
Public Const LVM_FIRST = &H1000
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_SETITEM = (LVM_FIRST + 6)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVSIL_NORMAL = 0
Public Const LVSIL_SMALL = 1
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2
Public Const LVIF_IMAGE = &H2
Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2
Public Const MAX_PATH = 260
Private Const WM_SETREDRAW As Long = &HB
Private Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Private Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
Private Const LVIF_STATE As Long = &H8
Public Enum IL_DrawStyle
 ILD_NORMAL = &H0
 ILD_TRANSPARENT = &H1
 ILD_MASK = &H10
 ILD_IMAGE = &H20
 ILD_ROP = &H40
 ILD_BLEND25 = &H2
 ILD_BLEND50 = &H4
 ILD_OVERLAYMASK = &HF00
 ILD_SELECTED = ILD_BLEND50
 ILD_FOCUS = ILD_BLEND25
 ILD_BLEND = ILD_BLEND50
End Enum
Public Enum SHGFI_flags
 SHGFI_LARGEICON = &H0 ' sfi.hIcon is large icon
 SHGFI_SMALLICON = &H1 ' sfi.hIcon is small icon
 SHGFI_OPENICON = &H2 ' sfi.hIcon is open icon
 SHGFI_SHELLICONSIZE = &H4 ' sfi.hIcon is shell size (not system size), rtns BOOL
 SHGFI_PIDL = &H8 ' pszPath is pidl, rtns BOOL
 SHGFI_USEFILEATTRIBUTES = &H10 ' pretend pszPath exists, rtns BOOL
 SHGFI_ICON = &H100 ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
 SHGFI_DISPLAYNAME = &H200 ' isf.szDisplayName is filled, rtns BOOL
 SHGFI_TYPENAME = &H400 ' isf.szTypeName is filled, rtns BOOL
 SHGFI_ATTRIBUTES = &H800 ' rtns IShellFolder::GetAttributesOf SFGAO_* flags
 SHGFI_ICONLOCATION = &H1000 ' fills sfi.szDisplayName with filename
' containing the icon, rtns BOOL
 SHGFI_EXETYPE = &H2000 ' rtns two ASCII chars of exe type
 SHGFI_SYSICONINDEX = &H4000 ' sfi.iIcon is sys il icon index, rtns hImagelist
 SHGFI_LINKOVERLAY = &H8000 ' add shortcut overlay to sfi.hIcon
 SHGFI_SELECTED = &H10000 ' sfi.hIcon is selected icon
End Enum
Private Const BASIC_SHGFI_FLAGS As Long = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Type LVITEM ' was LV_ITEM
 mask As Long
 iItem As Long
 iSubItem As Long
 State As Long
 stateMask As Long
 pszText As Long ' if String, must be pre-allocated before filled
 cchTextMax As Long
 iImage As Long
 lParam As Long
 iIndent As Long
End Type
Private Type SHFILEINFO ' shfi
 hIcon As Long
 iIcon As Long
 dwAttributes As Long
 szDisplayName As String * MAX_PATH
 szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As SHGFI_flags) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long ' <---

Public Sub LVRedraw(ByVal LVWnd As Long, ByVal Redraw As Boolean)
 SendMessage LVWnd, WM_SETREDRAW, Redraw, ByVal 0&
End Sub
Public Sub LVClear(ByVal LVWnd As Long)
 SendMessage LVWnd, LVM_DELETEALLITEMS, 0&, ByVal 0&
End Sub
Function LVSelCount(ByVal LVWnd As Long)
 LVSelCount = SendMessage(LVWnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)
End Function
Sub LVSelect(ByVal LVWnd As Long, ByVal Selected As Boolean)
 Dim LV As LV_ITEM
 With LV
  .mask = LVIF_STATE
  .State = Selected
  .stateMask = LVIS_SELECTED
 End With
 Call SendMessage(LVWnd, LVM_SETITEMSTATE, -1, LV)
End Sub
Public Function LVMax(ByVal LVWnd As Long) As Long
 LVMax = SendMessage(LVWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)
End Function

'SIL Macros
Public Function GetSystemImagelist(uFlags As Long) As Long
 Dim sfi As SHFILEINFO
 GetSystemImagelist = SHGetFileInfo("C:\", 0, sfi, Len(sfi), SHGFI_SYSICONINDEX Or uFlags)
End Function
Public Function GetFileIconIndex(sFile As String, uFlags As SHGFI_flags) As Long
 Dim sfi As SHFILEINFO
 If SHGetFileInfo(sFile, 0, sfi, Len(sfi), uFlags Or BASIC_SHGFI_FLAGS) Then '
  GetFileIconIndex = sfi.iIcon
 End If
End Function
Public Function GetFileTypeName(sFile As String, Optional Ext As Boolean) As String
 Dim sfi As SHFILEINFO
 If Ext Then
  If SHGetFileInfo(sFile, 0, sfi, Len(sfi), SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
   GetFileTypeName = TrimNull(sfi.szTypeName)
  End If
 Else
  If SHGetFileInfo(sFile, 0, sfi, Len(sfi), BASIC_SHGFI_FLAGS) Then
   GetFileTypeName = TrimNull(sfi.szTypeName)
  End If
 End If
End Function
' =============================================================================
' listview macros
Public Function ListView_SetImageList(hWnd As Long, himl As Long, iImageList As Long) As Long
 ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, iImageList, ByVal himl)
End Function

Public Function ListView_GetItem(hWnd As Long, pitem As LVITEM) As Boolean
 ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pitem)
End Function

Public Function ListView_SetItem(hWnd As Long, pitem As LVITEM) As Boolean
 ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pitem)
End Function

Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As Long) As Long
 ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal flags) 'MAKELPARAM(flags, 0))
End Function

Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As Boolean) As Boolean
 ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal Abs(fPartialOK)) 'MAKELPARAM(Abs(fPartialOK), 0))
End Function

Public Function ListView_SetItemState(hwndLV As Long, i As Long, State As Long, mask As Long) As Boolean
 Dim LVI As LVITEM
 LVI.State = State
 LVI.stateMask = mask
 ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, LVI)
End Function

' Returns the index of the item that is selected and has the focus rectangle (user-defined macro)

Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
 ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function

' Selects the specified item and gives it the focus rectangle.
' If the listview is multiselect (not LVS_SINGLESEL), does not
' de-select any currently selected items (user-defined macro)

Public Function ListView_SetSelectedItem(hwndLV As Long, i As Long) As Boolean
 ListView_SetSelectedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, _
   LVIS_FOCUSED Or LVIS_SELECTED)
End Function
