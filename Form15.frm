VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "External Browsers"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find ..."
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   1560
      Pattern         =   "*.brw"
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1

Private Sub Command1_Click()

If List1.ListCount = -1 Then
Exit Sub
Else

If List1.List(List1.ListIndex) = "" Then

MsgBox ("Please choose an item from the list to open!")
Exit Sub

Else

If List1.List(List1.ListIndex) = "IEXPLORER" Then
Form1.SaveTmp
ShellExecute hwnd, "open", List1.List(List1.ListIndex), "c:\e.tmp", vbNullString, conSwNormal

Else

SaveFN
ShellExecute hwnd, "open", List1.List(List1.ListIndex), "file:///C|/e.html", vbNullString, conSwNormal
End If

Unload Me
End If

End If
End Sub

Public Sub SaveFN()
Dim FileNum As Integer, Buffer As String
Dim FileName As String
Buffer = "<Html><Body>" & Form1.RichTextBox1.Text & "</Body></Html>" ' the opening and closing body and html tags are for files that may not begin with these tags like [!html:header!] & [!html:footer!] or [!system:verifyuser!] 'Both netscape and IE will parse it as a text file and not html especially netscape
FileName = "c:\" & "e.html"
FileNum = FreeFile
Open FileName For Output As FileNum
Print #FileNum, Buffer
Close FileNum
FileName = ""
End Sub

Private Sub Command2_Click()
Dim Path As String
Dim FileName As String
Dim R As String
Dim Name As String
Dim Executable As String

Name = Space(50)

With CommonDialog1
.Filter = "Applications (*.exe)|*.exe"
.ShowOpen

If .FileName = "" Then
Exit Sub
Else
Text1.Text = .FileName
Executable = Right(.FileName, 12)
List1.AddItem (Left(Executable, Len(Executable) - 4))
Path = App.Path & "\" & Executable & ".brw"
WritePrivateProfileString Executable, "Browser", Text1.Text, Path
End If

End With
File1.Refresh
List1.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If List1.ListCount < 1 Then
Exit Sub
Else

If List1.List(List1.ListIndex) = "" Then
MsgBox ("Please choose an item from the list to delete!")
Exit Sub
Else

Dim RetVal As String

RetVal = MsgBox("Are you sure you would like to delete " & List1.List(List1.ListIndex) & " from the list?", vbInformation + vbOKCancel)

If RetVal = vbOK Then

Kill App.Path & "\" & List1.List(List1.ListIndex) & ".exe.brw"

Unload Me
Form15.Show
Else

Exit Sub

End If

End If

End If
End Sub

Private Sub Form_Load()
'This is not a misprint
Form15Load  'This is not a misprint
'This is not a misprint
End Sub

Public Function Form15Load() 'may use for other global items
File1.Path = App.Path

Dim I As Integer

For I = 0 To File1.ListCount - 1

Dim X As String

Dim li As ListItem

X = Left(File1.List(I), Len(File1.List(I)) - 8)

List1.AddItem X

Next I

File1.Refresh

End Function


