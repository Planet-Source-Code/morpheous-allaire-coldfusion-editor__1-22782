VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form14 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sql Builder"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Search"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form14.frx":08CA
      Left            =   6360
      List            =   "Form14.frx":08D7
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   360
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Insert Comma"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   4680
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find DB"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   3240
      TabIndex        =   7
      Top             =   3000
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Insert as VTrader Tag"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   3720
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Insert as Auctioneer Tag"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form14.frx":08F1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1620
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   3015
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "Form14.frx":099F
      Left            =   6360
      List            =   "Form14.frx":0A00
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sql:"
      Height          =   195
      Left            =   6360
      TabIndex        =   15
      Top             =   720
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Scope:"
      Height          =   195
      Left            =   6360
      TabIndex        =   14
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Columns:"
      Height          =   195
      Left            =   3240
      TabIndex        =   9
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tables:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   525
   End
   Begin VB.Menu MnuTables 
      Caption         =   "Tables"
      Visible         =   0   'False
      Begin VB.Menu MnuGetColumns 
         Caption         =   "Get Columns"
      End
      Begin VB.Menu MnuEnterAsStatement 
         Caption         =   "Enter as statement"
      End
   End
   Begin VB.Menu MnuColumns 
      Caption         =   "Columns"
      Visible         =   0   'False
      Begin VB.Menu MnuClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const adSchemaColumns = 4

Private Sub Combo1_Click()
If Combo1.List(Combo1.ListIndex) = "Close" Then
If Option1.Value = True Then
Form1.RichTextBox1.SelText = "<#Query:Close#>"
Unload Me
Else
Form1.RichTextBox1.SelText = "[!Query:Close!]"
Unload Me
End If
Else
' Do Nothing as the user wants to build a custom sql statement
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim Quote As String
Dim DblQuote As String
Quote = """"
DblQuote = Quote & Quote
If RichTextBox1.Text <> "" Then
If Option1.Value = True Then
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
Form1.RichTextBox1.SelText = "<#" & LCase(RichTextBox1.Text) & " #>"
Else
If Combo1.Text <> "" Then
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
Form1.RichTextBox1.SelText = "[!Query:" & Combo1.Text & " Name=" & DblQuote & " SQL =" & Quote & LCase(RichTextBox1.Text) & Quote & "!]"
Else
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
Form1.RichTextBox1.SelText = "[!Query:Open" & " Name=" & DblQuote & " SQL =" & Quote & LCase(RichTextBox1.Text) & Quote & "!]"
End If
'
End If
'
Else
'
Exit Sub
'
End If
End Sub

Private Sub Command3_Click()
Form14.List2.Clear
Form14.List3.Clear
Form16.Show
End Sub

Private Sub Command4_Click()
RichTextBox1.Text = ""
End Sub

Private Sub Command5_Click()
List2.Clear
List3.Clear
Form18.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form18
End Sub

Private Sub List1_DblClick()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List1.List(List1.ListIndex))
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
List1_DblClick
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
If List2.ListCount > 0 Then
PopupMenu MnuTables
End If
End If
End Sub

Private Sub List3_DblClick()
If Check1.Value = 0 Then
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List3.List(List3.ListIndex))
Else
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List3.List(List3.ListIndex)) & ", "
End If
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
If List3.ListCount > 0 Then
PopupMenu MnuColumns
End If
End If
End Sub

Private Sub MnuClear_Click()
List3.Clear
End Sub

Private Sub MnuEnterAsStatement_Click()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List2.List(List2.ListIndex))
End Sub

Private Sub MnuGetColumns_Click()
Dim Db As Database
Dim Rst As Recordset
Dim I As Integer
Screen.MousePointer = vbHourglass
List3.Clear
Set Db = OpenDatabase(Form16.Path)

With Db

Set Rst = .OpenRecordset(List2.List(List2.ListIndex))

For I = 0 To Rst.Fields.Count - 1
List3.AddItem Rst.Fields(I).SourceField
Next I
Set Db = Nothing
End With
Screen.MousePointer = vbDefault
End Sub


