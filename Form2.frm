VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Find"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Find What:"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5640
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Enabled         =   0   'False
         Height          =   255
         Left            =   5640
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Boolean
Private Sub Command1_Click()
Dim TextFound As Integer
Form1.RichTextBox1.SelStart = 0
Form1.RichTextBox1.Find (Text1.Text)
Form1.RichTextBox1.SetFocus
TextFound = Form1.RichTextBox1.Find(Text1.Text)
Command1.Enabled = False
Command2.Enabled = True
If TextFound = -1 Then
Form2.Move 0, 0
MsgBox ("End of Document" & vbCrLf & "Text Not Found"), vbInformation, App.ProductName
End If
End Sub

Private Sub Command2_Click()
Form1.RichTextBox1.SetFocus
Form1.RichTextBox1.Find (Text1.Text), Form1.RichTextBox1.SelStart + 1
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
b = SetWinPos(1, Form2.hwnd)
Text1.Text = Form1.RichTextBox1.SelText
End Sub

Private Sub Text1_Change()

If Len(Form1.RichTextBox1.Text) = 0 Then
If Len(Text1.Text) = 0 Then
Command1.Enabled = False
Else
Command1.Enabled = True
Command2.Enabled = False
End If
Command1.Enabled = False
Else
Command1.Enabled = True
Command2.Enabled = False
End If

End Sub
