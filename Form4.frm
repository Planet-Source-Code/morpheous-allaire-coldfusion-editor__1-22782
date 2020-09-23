VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Tag Chooser"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   6720
      Pattern         =   "*.tag"
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add a Tag"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "Form4.frx":08CA
      Left            =   2760
      List            =   "Form4.frx":08F5
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "Form4.frx":0931
      Left            =   120
      List            =   "Form4.frx":0933
      TabIndex        =   1
      Top             =   2640
      Width           =   8055
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "Form4.frx":0935
      Left            =   120
      List            =   "Form4.frx":0937
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox List5 
      Height          =   2010
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "User Defined Tags:"
      Height          =   195
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Opening,Closing & Comment:"
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   2040
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Merkatum:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Html:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Menu MnuForm4A 
      Caption         =   "MnuForm4A"
      Visible         =   0   'False
      Begin VB.Menu MnuEditTag 
         Caption         =   "&Edit This Tag"
      End
      Begin VB.Menu MnuDeleteTag 
         Caption         =   "&Delete This Tag"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

Form17.Show

End Sub

Private Sub Form_Load()
Dim I As Integer
Dim T As String

File1.Refresh

List5.Refresh

File1.Path = App.Path

For I = 0 To File1.ListCount - 1

T = Left(File1.List(I), Len(File1.List(I)) - 4)

List5.AddItem (T)

Next I

End Sub

Private Sub List1_DblClick()

Form1.RichTextBox1.SelText = Form1.RichTextBox1.SelText & "<" & List1.List(List1.ListIndex) & ">" & "</" & List1.List(List1.ListIndex) & ">"

End Sub
Private Sub List3_DblClick()

Form1.RichTextBox1.SelText = Form1.RichTextBox1.SelText & List3.List(List3.ListIndex)

End Sub

Private Sub List4_DblClick()

Form1.RichTextBox1.SelText = Form1.RichTextBox1.SelText & List4.List(List4.ListIndex)

End Sub

Private Sub List5_DblClick()
Dim StrFileName
Dim SzReturn As String
Dim SName As String

SName = Space(50)

StrFileName = App.Path & "\" & File1.List(List5.ListIndex)

GetPrivateProfileString "Tag", "Tag", SzReturn, SName, Len(SName), StrFileName

Form1.RichTextBox1.SelText = SName

End Sub

Private Sub List5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then

PopupMenu MnuForm4A

End If

End Sub

Private Sub MnuDeleteTag_Click()

Dim RetVal As String

RetVal = MsgBox("You are about to remove the tag labeled as " & List5.List(List5.ListIndex), vbOKCancel + vbInformation, App.ProductName)

If RetVal = vbOK Then

Kill App.Path & "\" & List5.List(List5.ListIndex) & ".tag"

List5.RemoveItem (List5.ListIndex)

File1.Refresh

List5.Refresh

Else

Exit Sub

End If

End Sub

Private Sub MnuEditTag_Click()
Dim StrFileName
Dim SzReturn As String
Dim Tag As String

Tag = Space(200)

StrFileName = App.Path & "\" & List5.List(List5.ListIndex) & ".tag"


GetPrivateProfileString "Tag", "Tag", SzReturn, Tag, Len(Tag), StrFileName

Form17.Text1.Text = List5.List(List5.ListIndex)

Form17.Text2.Text = Tag

Form17.Show

End Sub
