VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl FileList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   PropertyPages   =   "FileList.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4155
   ToolboxBitmap   =   "FileList.ctx":0023
   Begin VB.PictureBox picLV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "imlHdr"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Date"
         Text            =   "Date Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlHdr 
      Left            =   3480
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileList.ctx":0335
            Key             =   "Desc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FileList.ctx":048F
            Key             =   "Asc"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Event Declarations:
Event ItemCheck(ByVal Item As ListItem) 'MappingInfo=LV,LV,-1,ItemCheck
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=LV,LV,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=LV,LV,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=LV,LV,-1,MouseUp
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=LV,LV,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=LV,LV,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=LV,LV,-1,KeyUp
Event ItemClick(ByVal Item As ListItem) 'MappingInfo=LV,LV,-1,ItemClick
'Default Property Values:
Const m_def_LedgerColor1 = &HD0FFCC
Const m_def_LedgerColor2 = &HE1FAFF
Public Enum eBS
 None
 [Fixed Single]
End Enum
Public Enum eApp
 Flat
 [3D]
End Enum
'Property Variables:
Private mHorizontalLedger As Boolean
Private mVerticalLedger As Boolean
Private mLedgerColor1 As OLE_COLOR
Private mLedgerColor2 As OLE_COLOR
Private mPath As String
Private mPattern As String
Private mShowSize As Boolean
Private mShowType As Boolean
Private mShowDate As Boolean
Private mDescend(3) As Boolean

Public Property Get ShowSize() As Boolean
Attribute ShowSize.VB_ProcData.VB_Invoke_Property = "FileListPage"
 ShowSize = mShowSize
End Property
Public Property Let ShowSize(ByVal NewVal As Boolean)
 mShowSize = NewVal
 RedoCols
 LoadFolder mPath
 SortAscending
 PropertyChanged "ShowSize"
End Property
Public Property Get ShowType() As Boolean
Attribute ShowType.VB_ProcData.VB_Invoke_Property = "FileListPage"
 ShowType = mShowType
End Property
Public Property Let ShowType(ByVal NewVal As Boolean)
 mShowType = NewVal
 RedoCols
 LoadFolder mPath
 SortAscending
 PropertyChanged "ShowType"
End Property
Public Property Get ShowDate() As Boolean
Attribute ShowDate.VB_ProcData.VB_Invoke_Property = "FileListPage"
 ShowDate = mShowDate
End Property
Public Property Let ShowDate(ByVal NewVal As Boolean)
 mShowDate = NewVal
 RedoCols
 LoadFolder mPath
 SortAscending
 PropertyChanged "ShowDate"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get Path() As String
Attribute Path.VB_MemberFlags = "400"
 Path = mPath
End Property

Public Property Let Path(ByVal NewVal As String)
 If Ambient.UserMode = False Then Err.Raise 387
 If FolderExists(NewVal) Then
  mPath = NewVal
  LoadFolder mPath
  SortAscending
  PropertyChanged "Path"
 Else
  Cleanup
  Err.Raise 53
 End If
End Property
Public Property Get Pattern() As String
Attribute Pattern.VB_ProcData.VB_Invoke_Property = "FileListPage"
 Pattern = mPattern
End Property
Public Property Let Pattern(ByVal NewVal As String)
 mPattern = NewVal
 LoadFolder mPath
 PropertyChanged "Pattern"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,View
Public Property Get View() As ListViewConstants
 View = LV.View
End Property

Public Property Let View(ByVal New_View As ListViewConstants)
 LV.View() = New_View
 If LV.View = lvwReport Then
  If mHorizontalLedger Or mVerticalLedger Then
   Set LV.Picture = picLV.Image
  End If
 Else
  Set LV.Picture = Nothing
 End If
 PropertyChanged "View"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_ProcData.VB_Invoke_Property = "FileListPage"
 FullRowSelect = LV.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
 LV.FullRowSelect() = New_FullRowSelect
 PropertyChanged "FullRowSelect"
End Property

Private Sub LV_ItemClick(ByVal Item As ListItem)
 RaiseEvent ItemClick(Item)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,MultiSelect
Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made."
Attribute MultiSelect.VB_ProcData.VB_Invoke_Property = "FileListPage"
 MultiSelect = LV.MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
 LV.MultiSelect() = New_MultiSelect
 PropertyChanged "MultiSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determines whether the selected item will display as selected when the ListView loses focus"
Attribute HideSelection.VB_ProcData.VB_Invoke_Property = "FileListPage"
 HideSelection = LV.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
 LV.HideSelection() = New_HideSelection
 PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,GridLines
Public Property Get GridLines() As Boolean
Attribute GridLines.VB_ProcData.VB_Invoke_Property = "FileListPage"
 GridLines = LV.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As Boolean)
 LV.GridLines() = New_GridLines
 PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,Checkboxes
Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_ProcData.VB_Invoke_Property = "FileListPage"
 Checkboxes = LV.Checkboxes
End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
 Dim i As Long
 LV.Checkboxes() = New_Checkboxes
 If New_Checkboxes Then
  For i = 1 To LV.ListItems.Count 'other wise they won't show
   LV.ListItems(i).Checked = False
  Next
 End If
 PropertyChanged "Checkboxes"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,HotTracking
Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_ProcData.VB_Invoke_Property = "FileListPage"
 HotTracking = LV.HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
 LV.HotTracking() = New_HotTracking
 PropertyChanged "HotTracking"
End Property

Public Function HitTest(x As Single, y As Single) As ListItem
 Set HitTest = LV.HitTest(x, y)
End Function
Public Function SelectCount() As Long
 SelectCount = LVSelCount(mHWndLV)
End Function
Public Function Selected(ByVal Index As Long) As Boolean
 Selected = LV.ListItems(Index).Selected
End Function
Public Function Count() As Long
 Count = LV.ListItems.Count
End Function
Public Function Item(ByVal Index As Long) As ListItem
 Set Item = LV.ListItems(Index)
End Function
Public Function FullPath(ByVal Index As Long) As String
 FullPath = QualifyPath(mPath) & LV.ListItems(Index).Text
End Function
Public Sub SelectAll()
 If LV.ListItems.Count Then
  LVSelect mHWndLV, True
 End If
End Sub
Public Sub SelectNone()
 If LV.ListItems.Count Then
  LVSelect mHWndLV, False
 End If
End Sub
'Public Sub Clear() 'removed since FileListBox doesn't allow it
' LVClear mHWndLV
'End Sub
Public Sub Refresh()
 LoadFolder mPath
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,0
Public Property Get HorizontalLedger() As Boolean
Attribute HorizontalLedger.VB_MemberFlags = "400"
 HorizontalLedger = mHorizontalLedger
End Property

Public Property Let HorizontalLedger(ByVal New_HorizontalLedger As Boolean)
 If Ambient.UserMode = False Then Err.Raise 387
 mHorizontalLedger = New_HorizontalLedger
 If mHorizontalLedger Then
  mVerticalLedger = False
  HLedger
 Else
  Set LV.Picture = Nothing
 End If
 PropertyChanged "HorizontalLedger"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,0
Public Property Get VerticalLedger() As Boolean
Attribute VerticalLedger.VB_MemberFlags = "400"
 VerticalLedger = mVerticalLedger
End Property

Public Property Let VerticalLedger(ByVal New_VerticalLedger As Boolean)
 If Ambient.UserMode = False Then Err.Raise 387
 mVerticalLedger = New_VerticalLedger
 If mVerticalLedger Then
  mHorizontalLedger = False
  VLedger
 Else
  Set LV.Picture = Nothing
 End If
 PropertyChanged "VerticalLedger"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get LedgerColor1() As OLE_COLOR
 LedgerColor1 = mLedgerColor1
End Property

Public Property Let LedgerColor1(ByVal New_LedgerColor1 As OLE_COLOR)
 mLedgerColor1 = New_LedgerColor1
 If LV.View = lvwReport Then
  If mHorizontalLedger Then
   HLedger
  ElseIf mVerticalLedger Then
   VLedger
  End If
 End If
 PropertyChanged "LedgerColor1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get LedgerColor2() As OLE_COLOR
 LedgerColor2 = mLedgerColor2
End Property

Public Property Let LedgerColor2(ByVal New_LedgerColor2 As OLE_COLOR)
 mLedgerColor2 = New_LedgerColor2
 If LV.View = lvwReport Then
  If mHorizontalLedger Then
   HLedger
  ElseIf mVerticalLedger Then
   VLedger
  End If
 End If
 PropertyChanged "LedgerColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,Appearance
Public Property Get Appearance() As eApp
Attribute Appearance.VB_Description = "Returns/sets whether or not controls, Forms or an MDIForm are painted at run time with 3-D effects."
 Appearance = LV.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As eApp)
 LV.Appearance() = New_Appearance
 PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,BorderStyle
Public Property Get BorderStyle() As eBS
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
 BorderStyle = LV.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As eBS)
 LV.BorderStyle() = New_BorderStyle
 PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,HoverSelection
Public Property Get HoverSelection() As Boolean
Attribute HoverSelection.VB_Description = "Returns/sets whether hover selection is enabled."
Attribute HoverSelection.VB_ProcData.VB_Invoke_Property = "FileListPage"
 HoverSelection = LV.HoverSelection
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
 LV.HoverSelection() = New_HoverSelection
 PropertyChanged "HoverSelection"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=LV,LV,-1,HideColumnHeaders
Public Property Get HideColumnHeaders() As Boolean
    HideColumnHeaders = LV.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_HideColumnHeaders As Boolean)
    LV.HideColumnHeaders() = New_HideColumnHeaders
    If LV.HideColumnHeaders = False Then
     SortAscending
    End If
    PropertyChanged "HideColumnHeaders"
End Property

'==============Private Routines================
Private Sub SortAscending()
 LV.Sorted = True
 LV.SortKey = 0
 LV.SortOrder = lvwAscending
 LV.ColumnHeaders(1).Icon = "Asc"
End Sub

Private Sub RedoCols()
 Dim CH As ColumnHeader
 With LV.ColumnHeaders
  .Clear
  .Add , , "Name"
  If mShowSize Then
   .Add , "Size", "Size"
  End If
  If mShowType Then
   .Add , "Type", "Type"
  End If
  If mShowDate Then
   .Add , "Date", "Date Modified"
  End If
  For Each CH In LV.ColumnHeaders
   CH.Width = (LV.Width \ .Count) - 160
  Next
 End With
 If mHorizontalLedger Then
  HLedger
 ElseIf mVerticalLedger Then
  VLedger
 End If
End Sub
Private Sub LoadFolder(ByVal Path As String)
 Dim LVI As LVITEM
 Dim LI As ListItem
 Dim FI() As TFileData
 Dim FICnt As Long
 Dim CH As ColumnHeader
 Dim i As Long
 If Ambient.UserMode = False Then Exit Sub
 Screen.MousePointer = vbHourglass
 LVClear mHWndLV
 FilesinFolder Path, FI, FICnt, mPattern
 LVI.mask = LVIF_IMAGE
 For i = 1 To FICnt
  With FI(i)
   Set LI = LV.ListItems.Add(, , FileTitle(.TName))
   If mShowSize Then
    Set CH = LV.ColumnHeaders("Size")
    LI.SubItems(CH.Index - 1) = .TSize
   End If
   If mShowType Then
    Set CH = LV.ColumnHeaders("Type")
    LI.SubItems(CH.Index - 1) = .TType
   End If
   If mShowDate Then
    Set CH = LV.ColumnHeaders("Date")
    LI.SubItems(CH.Index - 1) = .TDate
   End If
   LI.Selected = False
   LVI.iItem = LI.Index - 1
   LVI.iImage = GetFileIconIndex(.TName, SHGFI_SMALLICON)
   Call ListView_SetItem(mHWndLV, LVI)
  End With
 Next
 LV.Refresh
 Screen.MousePointer = vbDefault
End Sub
'These two just draw the appropriately sized
'rows or cols on picLV then assign the pic to the LV
Private Sub HLedger()
 Dim i As Long
 Dim c As Long
 If LV.View <> lvwReport Then Exit Sub
 If LV.ListItems.Count = 0 Then Exit Sub
 c = mLedgerColor1
 With LV
  .PictureAlignment = lvwTile
  picLV.Cls
  picLV.Width = Screen.Width
  picLV.Height = .Height
  For i = 1 To .ListItems.Count
   picLV.Line (0, (i - 1) * .ListItems(1).Height)-(picLV.Width, i * .ListItems(1).Height), c, BF
   If c = mLedgerColor1 Then c = mLedgerColor2 Else c = mLedgerColor1
  Next
  picLV.Refresh
  Set .Picture = picLV.Image
 End With
End Sub

Private Sub VLedger()
 Dim i As Long
 Dim c As Long
 If LV.View <> lvwReport Then Exit Sub
 If LV.ListItems.Count = 0 Then Exit Sub
 c = mLedgerColor1
 With LV
  .PictureAlignment = lvwTile
  picLV.Cls
  picLV.Width = .Width
  picLV.Height = Screen.Height
  For i = 1 To .ColumnHeaders.Count
   picLV.Line ((i - 1) * .ColumnHeaders(i).Width, 0)-(i * .ColumnHeaders(i).Width, picLV.ScaleHeight), c, BF
   If c = mLedgerColor1 Then c = mLedgerColor2 Else c = mLedgerColor1
  Next
  picLV.Line (.ColumnHeaders(i - 1).Left + .ColumnHeaders(i - 1).Width, 0)-(picLV.ScaleWidth, picLV.ScaleHeight), c, BF
  picLV.Refresh
  Set .Picture = picLV.Image
 End With
End Sub
Private Sub Cleanup()
 Call UnSubClass(mHWndLV)
 Call ListView_SetImageList(mHWndLV, 0, LVSIL_SMALL)
 Call ListView_SetImageList(mHWndLV, 0, LVSIL_NORMAL)
End Sub
'================ListView Events======================
Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 Dim i As Long
 fDescending = Not fDescending
 Dim HdrIndex As Long
 HdrIndex = ColumnHeader.Index - 1
 LV.SortKey = HdrIndex

 Select Case HdrIndex
  Case 0 'name
   'Use default sorting to sort the items in the list
   LV.SortKey = HdrIndex
   LV.SortOrder = Abs(fDescending) '=Abs(Not LV.SortOrder = 1)
   LV.Sorted = True
  Case 1 'Use sort routine to sort the size col
   If mShowSize Then
    LV.Sorted = False
    SendMessage mHWndLV, LVM_SORTITEMS, mHWndLV, ByVal FARPROC(AddressOf CompareValues)
   End If
  Case 2 'type
   If mShowType Then
    LV.SortKey = HdrIndex
    LV.SortOrder = Abs(fDescending) '=Abs(Not LV.SortOrder = 1)
    LV.Sorted = True
   End If
  Case 3 'Use sort routine to sort by date
   If mShowDate Then
    LV.Sorted = False
    SendMessage mHWndLV, LVM_SORTITEMS, mHWndLV, ByVal FARPROC(AddressOf CompareDates)
   End If
 End Select
 'clear header icons
 For i = 1 To LV.ColumnHeaders.Count
  LV.ColumnHeaders(i).Icon = 0
 Next
 'assign the Asc/Desc icons to the headers (see imagelist)
 mDescend(HdrIndex) = Not mDescend(HdrIndex)
 If mDescend(HdrIndex) Then
  ColumnHeader.Icon = "Desc"
 Else
  ColumnHeader.Icon = "Asc"
 End If
End Sub
Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub LV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub LV_KeyUp(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub LV_ItemCheck(ByVal Item As ListItem)
 RaiseEvent ItemCheck(Item)
End Sub

'===================UserControl Routines==============
Private Sub UserControl_Initialize()
 mHWndLV = LV.hWnd
End Sub
Private Sub UserControl_InitProperties()
 mPath = vbNullString
 mPattern = "*.*"
 mShowSize = True
 mShowType = True
 mShowDate = True
 LV.View = lvwReport
 mHorizontalLedger = False
 mVerticalLedger = False
 mLedgerColor1 = m_def_LedgerColor1
 mLedgerColor2 = m_def_LedgerColor2
End Sub
Private Sub UserControl_Terminate()
 Cleanup
End Sub
Private Sub UserControl_Resize()
 LV.Move 0, 0, ScaleWidth, ScaleHeight
 RedoCols
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Dim i As Long
 SetSystemImageLists 'assign system imglists & subclass
 With PropBag
  mPath = .ReadProperty("Path", vbNullString)
  mPattern = .ReadProperty("Pattern", "*.*")
  mShowSize = .ReadProperty("ShowSize", True)
  mShowType = .ReadProperty("ShowType", True)
  mShowDate = .ReadProperty("ShowDate", True)
  LV.View = .ReadProperty("View", 3)
  LV.FullRowSelect = .ReadProperty("FullRowSelect", False)
  LV.MultiSelect = .ReadProperty("MultiSelect", False)
  LV.HideSelection = .ReadProperty("HideSelection", True)
  LV.GridLines = .ReadProperty("GridLines", False)
  LV.Checkboxes = .ReadProperty("Checkboxes", False)
  LV.HotTracking = .ReadProperty("HotTracking", False)
  mDescend(1) = True
  mDescend(3) = True
  mHorizontalLedger = .ReadProperty("HorizontalLedger", False)
  mVerticalLedger = .ReadProperty("VerticalLedger", False)
  mLedgerColor1 = .ReadProperty("LedgerColor1", m_def_LedgerColor1)
  mLedgerColor2 = .ReadProperty("LedgerColor2", m_def_LedgerColor2)
  LV.Appearance = .ReadProperty("Appearance", 1)
  LV.BorderStyle = .ReadProperty("BorderStyle", 1)
  LV.HoverSelection = .ReadProperty("HoverSelection", False)
  LV.HideColumnHeaders = .ReadProperty("HideColumnHeaders", False)
 End With
 RedoCols
 LoadFolder mPath
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  Call .WriteProperty("Path", mPath, vbNullString)
  Call .WriteProperty("Pattern", mPattern, "*.*")
  Call .WriteProperty("ShowSize", mShowSize, True)
  Call .WriteProperty("ShowType", mShowType, True)
  Call .WriteProperty("ShowDate", mShowDate, True)
  Call .WriteProperty("View", LV.View, 3)
  Call .WriteProperty("FullRowSelect", LV.FullRowSelect, False)
  Call .WriteProperty("MultiSelect", LV.MultiSelect, False)
  Call .WriteProperty("HideSelection", LV.HideSelection, True)
  Call .WriteProperty("GridLines", LV.GridLines, False)
  Call .WriteProperty("Checkboxes", LV.Checkboxes, False)
  Call .WriteProperty("HotTracking", LV.HotTracking, False)
  Call .WriteProperty("HorizontalLedger", mHorizontalLedger, False)
  Call .WriteProperty("VerticalLedger", mVerticalLedger, False)
  Call .WriteProperty("LedgerColor1", mLedgerColor1, m_def_LedgerColor1)
  Call .WriteProperty("LedgerColor2", mLedgerColor2, m_def_LedgerColor2)
  Call .WriteProperty("Appearance", LV.Appearance, 1)
  Call .WriteProperty("BorderStyle", LV.BorderStyle, 1)
  Call .WriteProperty("HoverSelection", LV.HoverSelection, False)
  Call .WriteProperty("HideColumnHeaders", LV.HideColumnHeaders, False)
 End With
End Sub
