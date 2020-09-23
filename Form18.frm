VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form18 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Database Search"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9900
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form18.frx":0000
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include Subdirectories"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Search"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   2160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "In Folder"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Path As String

Private Sub Command1_Click()
Command1.Enabled = False
ListView1.ListItems.Clear

Screen.MousePointer = vbHourglass

FilesSearch Dir1.Path, "*.mdb"

Screen.MousePointer = vbDefault

Me.Caption = ListView1.ListItems.Count & " Database's found!"

Command1.Enabled = True

End Sub

Sub FilesSearch(DrivePath As String, Ext As String)

Dim XDir() As String

Dim TmpDir As String

Dim FFound As String

Dim DirCount As Integer

Dim X As Integer

Dim li As ListItem

DirCount = 0

ReDim XDir(0) As String

XDir(DirCount) = ""

If Right(DrivePath, 1) <> "\" Then

DrivePath = DrivePath & "\"

End If

'Enter here the code for showing the pat
' h being
'search. Example: Form1.label2 = DrivePa
' th
'Search for all directories and store in
' the
'XDir() variable

Me.Caption = DrivePath

DoEvents

TmpDir = Dir(DrivePath, vbDirectory)


Do While TmpDir <> ""


If TmpDir <> "." And TmpDir <> ".." Then


If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then

XDir(DirCount) = DrivePath & TmpDir & "\"

DirCount = DirCount + 1

ReDim Preserve XDir(DirCount) As String

End If

End If

TmpDir = Dir

Loop

'Searches for the files given by extensi

' on Ext

FFound = Dir(DrivePath & Ext)

Do Until FFound = ""

Set li = ListView1.ListItems.Add(, , FFound)

li.ListSubItems.Add , , DrivePath

li.ListSubItems.Add , , FileLen(DrivePath & FFound) & " Bytes"


FFound = Dir

Loop

If Check1.Value = 1 Then

For X = 0 To (UBound(XDir) - 1)

FilesSearch XDir(X), Ext

Next X
Else

End If

End Sub


Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
ListView1.ColumnHeaders(1).Width = ListView1.Width / 6
ListView1.ColumnHeaders(2).Width = ListView1.Width / 2
ListView1.ColumnHeaders(3).Width = ListView1.Width / 6
End Sub

Private Sub Form_Resize()

If Form18.Width < 3000 Then
Form18.Width = 5000
End If

If Form18.Height < 3000 Then
Form18.Height = 3000
End If

If Form18.WindowState <> 1 Then
Dir1.Height = Form18.Height - Drive1.Height - (Command1.Height * 3)
ListView1.Height = Form18.Height - (Command1.Height * 3)
ListView1.Width = Form18.Width - Dir1.Width - 200
ListView1.ColumnHeaders(1).Width = ListView1.Width / 6
ListView1.ColumnHeaders(2).Width = ListView1.Width / 2
ListView1.ColumnHeaders(3).Width = ListView1.Width / 6
Else
Dir1.Height = 0
ListView1.Height = 0
ListView1.ColumnHeaders(1).Width = 0
ListView1.ColumnHeaders(2).Width = 0
ListView1.ColumnHeaders(3).Width = 0
End If


End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'With ListView1

If ListView1.SortKey = ColumnHeader.Index - 1 Then

If ListView1.SortOrder = lvwAscending Then

ListView1.SortOrder = lvwDescending

Else

ListView1.SortOrder = lvwAscending

End If

Else

ListView1.SortOrder = lvwAscending

ListView1.SortKey = ColumnHeader.Index - 1

End If

ListView1.Sorted = True

'End With

End Sub

Private Sub ListView1_DblClick()

Const adSchemaTables = 20

Screen.MousePointer = vbHourglass

Path = ListView1.SelectedItem.ListSubItems(1).Text & ListView1.SelectedItem.Text

Set Db = CreateObject("AdoDb.Connection")

Db.Open "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & Path


Set Tables = Db.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))

If Not Tables.EOF Then

While Not Tables.EOF

Form14.List2.AddItem Tables("Table_Name")

Tables.MoveNext

Wend

Else

MsgBox ("This database does not have any tables!")

End If

Screen.MousePointer = vbDefault

Form16.Path = Path

Unload Me

End Sub
