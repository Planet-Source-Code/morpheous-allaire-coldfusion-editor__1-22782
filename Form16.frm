VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Database Wizard"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2310
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   2310
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   0
      Pattern         =   "*.mdb"
      TabIndex        =   1
      Top             =   2400
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Path As String

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Screen.MousePointer = vbHourglass
Const adSchemaTables = 20
Const adSchemaColumns = 4

Path = File1.Path & "\" & File1.List(File1.ListIndex)

Set Db = CreateObject("AdoDb.Connection")

Db.Open "Provider=Microsoft.Jet.OleDb.4.0; Data Source=" & Path


Set Tables = Db.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))

'Set Columns = Db.OpenSchema(adSchemaColumns, Array(Empty, Empty, Column))

If Not Tables.EOF Then

While Not Tables.EOF

Form14.List2.AddItem Tables("Table_Name")

'While Not Columns.EOF

'Form14.List3.AddItem Columns("Column_Name")

'Columns.MoveNext

'Wend

Tables.MoveNext

Wend

Else

MsgBox ("This database does not have any tables!")

End If

Screen.MousePointer = vbDefault
Unload Me
End Sub

Private Sub Form_Load()
Dir1.Path = "C:\"
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Form16.Height < 4000 Then
Form16.Height = 5385
End If
If Form16.Width < 2000 Then
Form16.Width = 2400
End If

Drive1.Width = Form16.Width - 100
Dir1.Width = Form16.Width - 100
File1.Width = Form16.Width - 100
File1.Height = Form16.Height - Dir1.Height - Drive1.Height
End Sub
