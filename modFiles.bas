Attribute VB_Name = "modFiles"
Option Explicit
Private Const ARR_MAX As Long = &H3FF& 'used to avoid redim preserve
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const rDayZeroBias             As Double = 109205# ' Abs(CDbl(#01-01-1601#))
Private Const rMillisecondPerDay       As Double = 10000000# * 60# * 60# * 24# / 10000#
'note: uses Currency vice FILETIME--much faster
Private Type WIN32_FIND_DATA
 dwFileAttributes As Long
 ftCreationTime As Currency
 ftLastAccessTime As Currency
 ftLastWriteTime As Currency
 nFileSizeHigh As Long
 nFileSizeLow As Long
 dwReserved0 As Long
 dwReserved1 As Long
 cFileName As String * MAX_PATH
 cAlternate As String * 14
End Type
Public Type TFileData
 TName As String
 TSize As String
 TType As String
 TDate As String
End Type
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Currency, lpLocalFileTime As Currency) As Long

Public Sub EFIStrArr(ByVal StartPath As String, _
          ByRef Arr() As TFileData, _
          ByRef Count As Long, _
          Optional ByVal Pattern As String = "*.*")
 Dim Found As Boolean
 Dim hFile As Long
 Dim mFD As WIN32_FIND_DATA
 hFile = FindFirstFile(StartPath & Pattern, mFD)
 Found = (hFile <> INVALID_HANDLE_VALUE)
 Do While Found
  If Not (mFD.dwFileAttributes And vbSystem) = vbSystem Then
   If Not (mFD.dwFileAttributes And vbDirectory) = vbDirectory Then
    Count = Count + 1
    If Count > UBound(Arr) Then
     ReDim Preserve Arr(1 To Count + ARR_MAX)
    End If
    With Arr(Count)
     .TName = StartPath & TrimNull(mFD.cFileName)
     .TSize = FileFmt(mFD.nFileSizeLow)
     .TType = GetFileTypeName(.TName)
     .TDate = GetDate(mFD.ftLastWriteTime)
    End With
   End If
  End If
  Found = FindNextFile(hFile, mFD)
 Loop
 FindClose hFile
End Sub
'much faster than Instr, etc
Public Function TrimNull(ByVal StrZ As String) As String
 TrimNull = Left$(StrZ, lstrlenW(StrPtr(StrZ)))
End Function
Public Function QualifyPath(ByVal sPath As String) As String
 If Right$(sPath, 1) <> "\" Then
  QualifyPath = sPath & "\"
 Else
  QualifyPath = sPath
 End If
End Function

Public Sub FilesinFolder(ByVal Path As String, Arr() As TFileData, ACnt As Long, Optional ByVal Pattern As String = "*.*")
 ACnt = 0
 ReDim Arr(1 To ARR_MAX)
 EFIStrArr QualifyPath(Path), Arr, ACnt, Pattern
 If ACnt Then
  ReDim Preserve Arr(1 To ACnt)
 End If
End Sub
Public Function FileTitle(ByVal Pth As String) As String
 FileTitle = Mid$(Pth, InStrRev(Pth, "\") + 1)
End Function

Public Function FileExists(ByVal sFile As String) As Boolean
 Dim eAttr As Long
 On Error Resume Next
 eAttr = GetAttr(sFile)
 FileExists = (Err.Number = 0) And ((eAttr And vbDirectory) = 0)
 On Error GoTo 0
End Function
Public Function FolderExists(ByVal sPath As String) As Boolean
 Dim eAttr As Long
 On Error Resume Next
 eAttr = GetAttr(sPath)
 FolderExists = (Err.Number = 0) And ((eAttr And vbDirectory) = vbDirectory)
 On Error GoTo 0
End Function
Public Function GetDate(WhichTime As Currency) As String
 Dim ftl As Currency
 Dim d As Date
 If FileTimeToLocalFileTime(WhichTime, ftl) Then
  d = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
  GetDate = Format$(d, "m/d/yyyy h:nn AM/PM")
 End If
End Function
Public Function FileFmt(ByVal Size As Long) As String
 If Size = 0 Then
  FileFmt = "0 KB"
 Else
  FileFmt = Format$(Size \ 1024 + 1, "###,###,###,##0") & " KB"
 End If
End Function

