VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   Caption         =   "Replace"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox TxtReplaceWith 
      Height          =   1455
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form3.frx":08CA
   End
   Begin RichTextLib.RichTextBox TxtReplaceWhat 
      Height          =   1455
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form3.frx":09AC
   End
   Begin VB.CommandButton CmdReplace 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   6960
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdReplaceAll 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Find:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Replace:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   645
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim b As Boolean

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdReplace_Click()
Dim Result As Long
Result = Form1.RichTextBox1.Find(TxtReplaceWhat.Text, 0)
If Result = -1 Then
Form3.Move 0, 0
MsgBox ("No Results Found"), vbOKOnly, App.ProductName
Else
Form1.RichTextBox1.SelText = TxtReplaceWith.Text
Form3.Move 0, 0
MsgBox ("Operation Complete!"), vbOKOnly, App.ProductName
End If
End Sub

Private Sub CmdReplaceAll_Click()
Dim IntCount As Integer
Dim LngPos As Long
IntCount = 0
LngPos = 0
With Form1.RichTextBox1
Do
If .Find(TxtReplaceWhat.Text, LngPos) = -1 Then
If IntCount < 0 Then
MsgBox ("No Results Found!"), vbOKOnly, App.ProductName
End If
Exit Do
Else
LngPos = .SelStart + .SelLength
IntCount = IntCount + 1
.SelText = TxtReplaceWith.Text
End If
Loop
Form3.Move 0, 0
MsgBox (IntCount & " " & "Items Replaced"), vbInformation + vbOKOnly, App.ProductName
End With
End Sub

Private Sub Form_Load()
b = SetWinPos(1, Form3.hWnd)
TxtReplaceWhat.Text = Form1.RichTextBox1.SelText
TxtReplaceWhat.SelStart = 0
TxtReplaceWhat.SelLength = Len(TxtReplaceWhat.Text)
End Sub

Private Sub TxtReplaceWhat_Change()
If Len(Form1.RichTextBox1.Text) = 0 Then
If Len(TxtReplaceWhat) = 0 Then
cmdReplace.Enabled = False
CmdReplaceAll.Enabled = False
Else
cmdReplace.Enabled = True
CmdReplaceAll.Enabled = True
End If
cmdReplace.Enabled = False
CmdReplaceAll.Enabled = False
Else
cmdReplace.Enabled = True
CmdReplaceAll.Enabled = True
End If
End Sub

Private Sub TxtReplaceWith_GotFocus()
TxtReplaceWith.SelStart = 0
TxtReplaceWith.SelLength = Len(TxtReplaceWith.Text)
End Sub
