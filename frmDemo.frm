VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDemo 
   Caption         =   "FileList Demo"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin FileListDemo.FileList FileList1 
      Height          =   3615
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6376
      FullRowSelect   =   -1  'True
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "Select None"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5880
      TabIndex        =   33
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Select All"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5040
      TabIndex        =   32
      Top             =   6120
      Width           =   855
   End
   Begin VB.Frame fraLedger 
      Caption         =   "Ledger"
      Height          =   615
      Left            =   0
      TabIndex        =   26
      Top             =   5880
      Width           =   3495
      Begin VB.CommandButton cmdC1 
         Caption         =   "C1"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdC2 
         Caption         =   "C2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdHLedger 
         Caption         =   "Horizizontal"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdVLedger 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdShowSel 
      Caption         =   "Show Selected"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame fraMisc 
      Caption         =   "ListView Props"
      Height          =   1455
      Left            =   3600
      TabIndex        =   15
      Top             =   4440
      Width           =   3375
      Begin VB.CheckBox chkMisc 
         Caption         =   "HideColumnHeaders"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   31
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkMisc 
         Caption         =   "HoverSelection"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkMisc 
         Caption         =   "HotTracking"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkMisc 
         Caption         =   "GridLines"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkMisc 
         Caption         =   "CheckBoxes"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkMisc 
         Caption         =   "MultiSelect"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkMisc 
         Caption         =   "HideSelection"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkMisc 
         Caption         =   "FullRowSelect"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdPatt 
      Caption         =   "OK"
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Top             =   3960
      Width           =   495
   End
   Begin VB.Frame fraView 
      Caption         =   "View"
      Height          =   1455
      Left            =   1680
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
      Begin VB.OptionButton optView 
         Caption         =   "Report"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optView 
         Caption         =   "List"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optView 
         Caption         =   "SmallIcon"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optView 
         Caption         =   "Icon"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtPatt 
      Height          =   285
      Left            =   3840
      TabIndex        =   7
      Text            =   "*.*"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "C:\Program Files\Microsoft Visual Studio\VB98"
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Frame fraCols 
      Caption         =   "Columns"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
      Begin VB.CheckBox chkCol 
         Caption         =   "Date Modified"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkCol 
         Caption         =   "Type"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCol 
         Caption         =   "Size"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6240
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Count"
      Height          =   255
      Left            =   5760
      TabIndex        =   24
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Pattern"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Path"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Loading As Boolean

Private Sub Form_Load()
 Loading = True
 With FileList1
  If Len(txtPath.Text) Then
   .Path = txtPath.Text
   Caption = "FileList Demo: " & txtPath.Text
  End If
  optView(.View).Value = True
  chkCol(0).Value = -CLng(.ShowSize)
  chkCol(1).Value = -CLng(.ShowType)
  chkCol(2).Value = -CLng(.ShowDate)
  chkMisc(0).Value = -CLng(.FullRowSelect)
  chkMisc(1).Value = -CLng(.HideSelection)
  chkMisc(2).Value = -CLng(.MultiSelect)
  chkMisc(3).Value = -CLng(.Checkboxes)
  chkMisc(4).Value = -CLng(.GridLines)
  chkMisc(5).Value = -CLng(.HotTracking)
  chkMisc(6).Value = -CLng(.HoverSelection)
  chkMisc(7).Value = -CLng(.HideColumnHeaders)
  txtPatt.Text = .Pattern
  cmdAll.Enabled = .MultiSelect
  cmdNone.Enabled = .MultiSelect
  cmdC1.BackColor = .LedgerColor1
  cmdC2.BackColor = .LedgerColor2
  lblCount.Caption = .Count & " Files"
 End With
 Loading = False
End Sub
Private Sub cmdAll_Click()
 FileList1.SelectAll
 FileList1.SetFocus
End Sub

Private Sub cmdNone_Click()
 FileList1.SelectNone
 FileList1.SetFocus
End Sub

Private Sub chkMisc_Click(Index As Integer)
 If Loading Then Exit Sub
 With FileList1
  Select Case Index
   Case 0
    .FullRowSelect = Not .FullRowSelect
   Case 1
    .HideSelection = Not .HideSelection
   Case 2
    .MultiSelect = Not .MultiSelect
   Case 3
    .Checkboxes = Not .Checkboxes
   Case 4
    .GridLines = Not .GridLines
   Case 5
    .HotTracking = Not .HotTracking
   Case 6
    .HoverSelection = Not .HoverSelection
   Case 7
    .HideColumnHeaders = Not .HideColumnHeaders
  End Select
  cmdAll.Enabled = .MultiSelect
  cmdNone.Enabled = .MultiSelect
 End With
End Sub

Private Sub cmdHLedger_Click()
 FileList1.HorizontalLedger = Not FileList1.HorizontalLedger
 cmdC1.Enabled = FileList1.HorizontalLedger
 cmdC2.Enabled = cmdC1.Enabled
 FileList1.SetFocus
End Sub
Private Sub cmdVLedger_Click()
 FileList1.VerticalLedger = Not FileList1.VerticalLedger
 cmdC1.Enabled = FileList1.VerticalLedger
 cmdC2.Enabled = cmdC1.Enabled
 FileList1.SetFocus
End Sub

Private Sub cmdShowSel_Click()
 Dim i As Long
 If FileList1.SelectCount Then
  For i = 1 To FileList1.Count
   If FileList1.Selected(i) Then
    Debug.Print FileList1.FullPath(i)
   End If
  Next
 End If
 FileList1.SetFocus
End Sub
Private Sub cmdC1_Click()
 With CD
  .Color = cmdC1.BackColor
  On Error Resume Next
  .ShowColor
  If Err = cdlCancel Then Exit Sub
  cmdC1.BackColor = .Color
  FileList1.LedgerColor1 = .Color
  FileList1.SetFocus
 End With
End Sub

Private Sub cmdC2_Click()
 With CD
  .Color = cmdC2.BackColor
  On Error Resume Next
  .ShowColor
  If Err = cdlCancel Then Exit Sub
  cmdC2.BackColor = .Color
  FileList1.LedgerColor2 = .Color
  FileList1.SetFocus
End With
End Sub


Private Sub FileList1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim LI As ListItem
 Set LI = FileList1.HitTest(x, y)
 If Not LI Is Nothing Then
  Debug.Print "MouseDown: " & LI.Text
 End If
End Sub

Private Sub FileList1_ItemClick(ByVal Item As MSComctlLib.ListItem)
 Debug.Print "ItemClick: " & Item.Text
End Sub

Private Sub chkCol_Click(Index As Integer)
 If Loading Then Exit Sub
 Select Case Index
  Case 0
   FileList1.ShowSize = Not FileList1.ShowSize
  Case 1
   FileList1.ShowType = Not FileList1.ShowType
  Case 2
   FileList1.ShowDate = Not FileList1.ShowDate
 End Select
End Sub
Private Sub optView_Click(Index As Integer)
 FileList1.View = Index
 EnableCols (Index = 3)
End Sub
Private Sub EnableCols(ByVal State As Boolean)
 Dim i As Long
 fraCols.Enabled = State
 For i = 0 To 2
  chkCol(i).Enabled = State
 Next
End Sub
Private Sub cmdPath_Click()
 Dim BP As String
 If (Len(txtPath.Text) = 0) Or (FolderExists(txtPath.Text) = False) Then
  Exit Sub
 End If
 BP = BrowseForFolderByPath(txtPath.Text, hWnd, "Select A Folder", False)
 If Len(BP) Then
  txtPath.Text = BP
  FileList1.Path = BP
 End If
 Caption = "FileList Demo: " & txtPath.Text
 lblCount.Caption = FileList1.Count & " Files"
End Sub

