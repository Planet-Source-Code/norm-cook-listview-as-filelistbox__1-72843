Attribute VB_Name = "modBrowse"
Option Explicit
Private Const WM_USER = &H400
Private Const MAX_PATH As Long = 260
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
Private Const BIF_USENEWUI As Long = &H40
Private Type BROWSEINFO
 hOwner As Long
 pidlRoot As Long
 pszDisplayName As String
 lpszTitle As String
 ulFlags As Long
 lpfn As Long
 lParam As Long
 iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
 Select Case uMsg
  Case BFFM_INITIALIZED
   Call SendMessage(hWnd, BFFM_SETSELECTIONA, 1&, ByVal lpData)
  Case Else:
 End Select
End Function

Private Function FARPROC(pfn As Long) As Long
 FARPROC = pfn
End Function

Public Function BrowseForFolderByPath(ByVal StartPath As String, _
 Optional ByVal FHWnd As Long, Optional ByVal Title As String, _
 Optional ByVal NewFolder As Boolean) As String
 Dim BI As BROWSEINFO
 Dim pidl As Long
 Dim lpSelPath As Long
 Dim sPath As String * MAX_PATH
 With BI
  .hOwner = FHWnd
  .pidlRoot = 0
  .lpszTitle = Title
  .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
  lpSelPath = LocalAlloc(LPTR, Len(StartPath) + 1)
  CopyMemory ByVal lpSelPath, ByVal StartPath, Len(StartPath) + 1
  .lParam = lpSelPath
  If NewFolder Then
   .ulFlags = BIF_USENEWUI
  End If
 End With
 pidl = SHBrowseForFolder(BI)
 If pidl Then
  If SHGetPathFromIDList(pidl, sPath) Then
   BrowseForFolderByPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
  Else
   BrowseForFolderByPath = ""
  End If
  Call CoTaskMemFree(pidl)
 Else
  BrowseForFolderByPath = ""
 End If
 Call LocalFree(lpSelPath)
End Function



