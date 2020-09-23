VERSION 5.00
Begin VB.Form Form17 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tags"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form17.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      MaxLength       =   300
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel:"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save:"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tag:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox ("Please enter a name for this tag!")
Else
SaveTag
Unload Me
Unload Form4
Form4.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Function SaveTag()
On Error Resume Next
Dim Path As String
Dim Tag As String
Path = App.Path & "\" & Text1.Text & ".tag"
WritePrivateProfileString "Tag", "Tag", Text2.Text, Path
Form4.List5.AddItem (Text1.Text)
End Function
