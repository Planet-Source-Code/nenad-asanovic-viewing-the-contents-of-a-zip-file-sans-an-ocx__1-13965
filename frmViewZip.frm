VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewZip 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "The contents of the ZIP file"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   7726
            MinWidth        =   7727
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1032
            MinWidth        =   1024
            TextSave        =   "01:58"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1455
            MinWidth        =   1448
            TextSave        =   "31-12-00"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlZip 
      Left            =   2640
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.Zipit Zipit1 
      Left            =   4920
      Top             =   1680
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MSComctlLib.ListView lvwZip 
      Height          =   2470
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4366
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Packed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ratio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date/Time"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmViewZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'\B/----------------------------For the listview column width----------------------------
Public Enum LVSCWII_Styles
   LVSCWII_AUTOSIZE = -1
   LVSCWII_AUTOSIZE_USEHEADER = -2
End Enum
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
'\B/------------------------For the resizing of the controls------------------------
Dim frmHeight As Single, frmWidth As Single
'/E\------------------------For the resizing of the controls------------------------
'/E\----------------------------For the listview column width----------------------------
Private Sub Form_Load()
On Error Resume Next
Zipit1.FileName = App.Path & "\test.zip"
'\B/---------------------------------For the column width---------------------------------
   Call LVSetAllColWidths(lvwZip, LVSCWII_AUTOSIZE_USEHEADER)
'\B/------------------------For the resizing of the controls------------------------
    frmHeight = frmViewZip.Height
    frmWidth = frmViewZip.Width
'/E\------------------------For the resizing of the controls------------------------
End Sub
'Private Sub Form_Activate()
'lvwZip.SetFocus
'    SendKeys "^{END}", True
'    SendKeys "^{HOME}", True
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Unload Me
End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim RetVal
    RetVal = Shell("Start.exe " & "mailto:n63@bigfoot.com?Subject=Beautiful!_I'm_your_fan!", 0)
End Sub

Private Sub Zipit1_OnArchiveUpdate()
    'The archive has been updated so refresh the list
    Dim itmX As ListItem
    Dim r As Long
    Dim I As Long
    Dim Files As New ZipFileEntry

    'Get the number of files in the archive
    r = Zipit1.vFiles.Count

    'Show the amount of files in the archive
    frmViewZip.StatusBar1.Panels(1).Text = Format(r) & " file(s) in archive"
    
    'Clear the list
    lvwZip.ListItems.Clear
    
    'Loop through each file in the archive
    For I = 1 To r
        'Store file info in a variable for ease of use
        'because the intellisense will give help
        Set Files = Zipit1.vFiles.Item(I)
        With Files
            'Add a item to the list
            Set itmX = lvwZip.ListItems.Add(, , .FileName)
            'Add the info
            itmX.Tag = I
            itmX.SubItems(1) = .UncompressedSize
            itmX.SubItems(2) = .CompressedSize
            'Trap div by zero
            If .UncompressedSize <> 0 Then
                itmX.SubItems(3) = Format(CInt((1 - (.CompressedSize / .UncompressedSize)) * 100)) & "%"
            Else
                itmX.SubItems(3) = "0%"
            End If
            itmX.SubItems(4) = .FileDateTime
        End With
    Next I
End Sub
'\B/---------------------------------For the Column Width---------------------------------
Public Sub LVSetColWidth(lv As ListView, ByVal ColumnIndex As Long, ByVal Style As LVSCWII_Styles)
   With lv
      If .View = lvwReport Then
         If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
            Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
         End If
      End If
   End With
End Sub
Public Sub LVSetAllColWidths(lv As ListView, ByVal Style As LVSCWII_Styles)
   Dim ColumnIndex As Long
   With lv
      For ColumnIndex = 1 To .ColumnHeaders.Count
         LVSetColWidth lv, ColumnIndex, Style
      Next ColumnIndex
   End With
End Sub
'/E\---------------------------------For the Column Width---------------------------------
Private Sub Form_Resize()
On Error Resume Next
    lvwZip.Height = lvwZip.Height + (frmViewZip.Height - frmHeight)
    lvwZip.Width = lvwZip.Width + (frmViewZip.Width - frmWidth)
    frmWidth = frmViewZip.Width
    frmHeight = frmViewZip.Height
End Sub
