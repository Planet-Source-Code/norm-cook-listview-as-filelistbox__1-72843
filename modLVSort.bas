Attribute VB_Name = "modLVSort"
Option Explicit
'Courtesy Randy Birch (http://vbnet.mvps.org/)
Public Const LVFI_PARAM As Long = &H1
Public Const LVIF_TEXT As Long = &H1
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_FINDITEM As Long = (LVM_FIRST + 13)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
Public Type POINTAPI
 x As Long
 y As Long
End Type
Public Type LV_FINDINFO
 flags As Long
 psz As String
 lParam As Long
 pt As POINTAPI
 vkDirection As Long
End Type
Public Type LV_ITEM
 mask As Long
 iItem As Long
 iSubItem As Long
 State As Long
 stateMask As Long
 pszText As String
 cchTextMax As Long
 iImage As Long
 lParam As Long
 iIndent As Long
End Type
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public objFind As LV_FINDINFO
Public objItem As LV_ITEM
Public fDescending As Boolean



Public Function CompareDates(ByVal lParam1 As Long, _
                             ByVal lParam2 As Long, _
                             ByVal hWnd As Long) As Long

'CompareDates: This is the sorting routine that gets passed to the
'ListView control to provide the comparison test for date values.

'Compare returns:
' 0 = Less Than
' 1 = Equal
' 2 = Greater Than

 Dim dDate1 As Date
 Dim dDate2 As Date

 'Obtain the item names and dates corresponding to the
 'input parameters
 dDate1 = ListView_GetItemDate(hWnd, lParam1)
 dDate2 = ListView_GetItemDate(hWnd, lParam2)

 'based on the Public variable fDescending set in the
 'columnheader click sub, sort the dates appropriately:
 Select Case fDescending
  Case True: 'sort descending

   If dDate1 < dDate2 Then
    CompareDates = 0
   ElseIf dDate1 = dDate2 Then
    CompareDates = 1
   Else: CompareDates = 2
   End If

  Case Else: 'sort ascending

   If dDate1 > dDate2 Then
    CompareDates = 0
   ElseIf dDate1 = dDate2 Then
    CompareDates = 1
   Else: CompareDates = 2
   End If

 End Select

End Function


Public Function CompareValues(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hWnd As Long) As Long

'CompareValues: This is the sorting routine that gets passed to the
'ListView control to provide the comparison test for numeric values.

'Compare returns:
' 0 = Less Than
' 1 = Equal
' 2 = Greater Than

 Dim val1 As String
 Dim val2 As String

 'Obtain the item names and values corresponding
 'to the input parameters
 val1 = ListView_GetItemValueStr(hWnd, lParam1)
 val2 = ListView_GetItemValueStr(hWnd, lParam2)

 'based on the Public variable fDescending set in the
 'columnheader click sub, sort the values appropriately:
 Select Case fDescending
  Case True: 'sort descending

   If val1 < val2 Then
    CompareValues = 0
   ElseIf val1 = val2 Then
    CompareValues = 1
   Else: CompareValues = 2
   End If

  Case Else: 'sort ascending

   If val1 > val2 Then
    CompareValues = 0
   ElseIf val1 = val2 Then
    CompareValues = 1
   Else: CompareValues = 2
   End If

 End Select

End Function


Public Function ListView_GetItemDate(hWnd As Long, lParam As Long) As Date

 Dim hIndex As Long
 Dim r As Long

 'Convert the input parameter to an index in the list view
 objFind.flags = LVFI_PARAM
 objFind.lParam = lParam
 hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)

 'Obtain the value of the specified list view item.
 'The objItem.iSubItem member is set to the index
 'of the column that is being retrieved.
 objItem.mask = LVIF_TEXT
 objItem.iSubItem = 3 'date modified column
 objItem.pszText = Space$(32)
 objItem.cchTextMax = Len(objItem.pszText)

 'get the string at subitem 1
 'and convert it into a date and exit
 r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
 If r > 0 Then
  ListView_GetItemDate = CDate(Left$(objItem.pszText, r))
 End If


End Function


Public Function ListView_GetItemValueStr(hWnd As Long, lParam As Long) As String

 Dim hIndex As Long
 Dim r As Long

 'Convert the input parameter to an index in the list view
 objFind.flags = LVFI_PARAM
 objFind.lParam = lParam
 hIndex = SendMessage(hWnd, LVM_FINDITEM, -1, objFind)

 'Obtain the value of the specified list view item.
 'The objItem.iSubItem member is set to the index
 'of the column that is being retrieved.
 objItem.mask = LVIF_TEXT
 objItem.iSubItem = 1 'size column
 objItem.pszText = Space$(32)
 objItem.cchTextMax = Len(objItem.pszText)

 'get the string at subitem 2
 'and convert it into a long
 Dim tmp As String
 r = SendMessage(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
 If r > 0 Then
  'make it right justified for sorting numbers
  tmp = Format$(Left$(objItem.pszText, r), "@@@@@@@@@@@@")
  ' remove the comma from the string for sorting
  ListView_GetItemValueStr = Replace$(tmp, ",", vbNullString)
 End If
End Function

Public Function FARPROC(ByVal pfn As Long) As Long

'A procedure that receives and returns
'the value of the AddressOf operator.
'This workaround is needed as you can't assign
'AddressOf directly to an API when you are also
'passing the value ByVal in the statement
'(as is being done with SendMessage)

 FARPROC = pfn

End Function



