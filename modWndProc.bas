Attribute VB_Name = "modWndProc"
Option Explicit
'Courtesy Brad Martinez (http://btmtz.mvps.org/)
Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"
Public Enum GWL_nIndex
 GWL_WNDPROC = (-4)
 GWL_ID = (-12)
 GWL_STYLE = (-16)
 GWL_EXSTYLE = (-20)
End Enum
Public Type NMHDR
 hwndFrom As Long ' Window handle of control sending message
 idFrom As Long ' Identifier of control sending message
 code As Long ' Specifies the notification code
End Type
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long
Public mHWndLV As Long

Public Sub SetSystemImageLists()
 Dim dwStyle As Long
 Dim m_himlSysSmall As Long
 Dim m_himlSysLarge As Long

 dwStyle = GetWindowLong(mHWndLV, GWL_STYLE)
 If ((dwStyle And LVS_SHAREIMAGELISTS) = False) Then
  Call SetWindowLong(mHWndLV, GWL_STYLE, dwStyle Or LVS_SHAREIMAGELISTS)
 End If

 ' Next get the handles of the system's small and large icon imagelists
 ' in the moudle level variables
 m_himlSysSmall = GetSystemImagelist(SHGFI_SMALLICON)
 m_himlSysLarge = GetSystemImagelist(SHGFI_LARGEICON)
 If (m_himlSysSmall <> 0) And (m_himlSysLarge <> 0) Then

  ' Assign the respective handles of the imagelists to the ListView. We
  ' will set the ListItem image indices directly in LoadSysILIcons proc below.
  ' As far as the VB ListView's internal code is concerned, it's not using
  ' any imagelists, both ListItem icon properties will return Empty.
  Call ListView_SetImageList(mHWndLV, m_himlSysSmall, LVSIL_SMALL)
  Call ListView_SetImageList(mHWndLV, m_himlSysLarge, LVSIL_NORMAL)

  ' The only reason we need to subclass the ListView is to prevent it from
  ' removing our system imagelist assignments (which it will do if left unchecked...)
  Call SubClass(mHWndLV, AddressOf WndProc)
 End If
End Sub
Public Function SubClass(hWnd As Long, lpfnNew As Long, Optional objNotify As Object = Nothing) As Boolean
 Dim lpfnOld As Long
 Dim fSuccess As Boolean
 On Error GoTo Out
 If GetProp(hWnd, OLDWNDPROC) Then
  SubClass = True
  Exit Function
 End If
 lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)
 If lpfnOld Then
  fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
  If (objNotify Is Nothing) = False Then
   fSuccess = fSuccess And SetProp(hWnd, OBJECTPTR, ObjPtr(objNotify))
  End If
 End If
Out:
 If fSuccess Then
  SubClass = True
 Else
  If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
  MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
    "Err# " & Err.Number & ": " & Err.Description, vbExclamation
 End If
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
 Dim lpfnOld As Long
 lpfnOld = GetProp(hWnd, OLDWNDPROC)
 If lpfnOld Then
  If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
   Call RemoveProp(hWnd, OLDWNDPROC)
   Call RemoveProp(hWnd, OBJECTPTR)
   UnSubClass = True
  End If ' SetWindowLong
 End If ' lpfnOld
End Function
Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Select Case uMsg
   ' ======================================================
   ' Prevent the ListView from removing our system imagelist assignment,
   ' which it will do when it sees no VB ImageList associated with it.
   ' (the ListView can't be subclassed when we're assigning imagelists...)
  Case LVM_SETIMAGELIST
   Exit Function
 End Select
 WndProc = CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
End Function
