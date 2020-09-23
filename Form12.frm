VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form12 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Settings..."
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Show Cursor Outline"
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "Form12.frx":08CA
      Left            =   1800
      List            =   "Form12.frx":08CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   " Font"
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "FontSize"
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form12.frx":08CE
      Left            =   600
      List            =   "Form12.frx":08F6
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "ForeColor"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "BackColor"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Start with template"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Form12.frx":0929
      Left            =   120
      List            =   "Form12.frx":0939
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form12.frx":0963
      Left            =   2760
      List            =   "Form12.frx":0973
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form12.frx":0990
      Left            =   600
      List            =   "Form12.frx":09A0
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "About Auto Save...."
      Height          =   1335
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "Form12.frx":09BD
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto Save Current File?"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7646
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public T As Boolean
Public C As Boolean

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
Form1.AutoSave
Form1.Timer1.Enabled = True
Form1.MnuAutoSave.Enabled = True

If Check1.Value = 1 Then
Form1.StatusBar1.Panels(3).Text = "Autosave is on"
End If

Else
Form1.Timer1.Enabled = False
Form1.MnuAutoSave.Enabled = False
Form1.StatusBar1.Panels(3).Text = "Autosave is " & Check1.Value

If Check1.Value = 0 Then
Form1.StatusBar1.Panels(3).Text = "Autosave is off"
End If

End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then 'checked
T = True
Command1.Enabled = True
Text2.Locked = False
Else
T = False
Command1.Enabled = False
Text2.Locked = True
SetString
End If
End Sub

Private Sub Check3_Click()
On Error Resume Next
Dim Path As String
Path = App.Path & "\Satellite.ini"
WritePrivateProfileString "ShowCursorOutline", "Yes,No", Check3.Value, Path
SetCursor
End Sub

Private Sub Combo1_Click()
Form1.RichTextBox1.BackColor = Combo3.List(Combo1.ListIndex)
Dim Path As String
Dim Name As String

Path = App.Path & "\Satellite.ini"

WritePrivateProfileString "BackColor", "Color", Combo1.ListIndex, Path
End Sub

Private Sub Combo2_Click()
Form1.RichTextBox1.SelStart = 0
Form1.RichTextBox1.SelLength = Len(Form1.RichTextBox1.Text)
Form1.RichTextBox1.SelColor = Combo3.List(Combo2.ListIndex)

Dim Path As String
Dim Name As String

Path = App.Path & "\Satellite.ini"

WritePrivateProfileString "ForeColor", "Color", Combo2.ListIndex, Path
End Sub

Private Sub Combo4_Click()
Form1.RichTextBox1.SelStart = 0
Form1.RichTextBox1.SelLength = Len(Form1.RichTextBox1.Text)
Form1.RichTextBox1.SelFontSize = Combo4.List(Combo4.ListIndex)
Dim Path As String
Dim Name As String

Path = App.Path & "\Satellite.ini"

WritePrivateProfileString "FontSize", "Size", Combo4.List(Combo4.ListIndex), Path
Form1.RichTextBox1.SelStart = 0
End Sub

Private Sub Combo5_Click()
Form1.RichTextBox1.SelStart = 0
Form1.RichTextBox1.SelLength = Len(Form1.RichTextBox1.Text)
Form1.RichTextBox1.SelFontName = Combo5.List(Combo5.ListIndex)
Dim Path As String
Dim Name As String

Path = App.Path & "\Satellite.ini"

WritePrivateProfileString "FontName", "Font", Combo5.List(Combo5.ListIndex), Path
Form1.RichTextBox1.SelStart = 0
End Sub

Private Sub Command1_Click()
On Error Resume Next

Dim Path As String
Dim Name As String
Path = App.Path & "\Satellite.ini"
Name = Space(50)
CommonDialog1.Filter = "All Files (*.*)|*.*"
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then
Exit Sub
Else
Text2.Text = CommonDialog1.FileName
WritePrivateProfileString "Template", "Template", Text2.Text, Path
WritePrivateProfileString "StartWithTemplate", "Yes,No", Check2.Value, Path
Form1.GetGlobals
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim II As Integer
For II = 0 To Screen.FontCount
Combo5.AddItem Screen.Fonts(II)
Next II
SetCursor
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs(2).Selected = True Then
Check1.Visible = True
Frame1.Visible = True
Check2.Visible = True
Text2.Visible = True
Command1.Visible = True
Else
Check1.Visible = False
Frame1.Visible = False
Check2.Visible = False
Text2.Visible = False
Command1.Visible = False
End If

If TabStrip1.Tabs(3).Selected = True Then
Combo1.Visible = True
Combo2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Combo4.Visible = True
Text6.Visible = True
Combo5.Visible = True
Check3.Visible = True
Else
Combo1.Visible = False
Combo2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Combo4.Visible = False
Text6.Visible = False
Combo5.Visible = False
Check3.Visible = False
End If
End Sub

Public Function SetString()
On Error Resume Next
Dim Path As String
Dim Name As String
Path = App.Path & "\Satellite.ini"
Text2.Text = ""
WritePrivateProfileString "Template", "Template", "", Path
WritePrivateProfileString "StartWithTemplate", "Yes,No", "0", Path
End Function

Public Function SetCursor()
Dim StrFileName
Dim SzReturn As String
Dim Cursor As String
Cursor = Space(5)

StrFileName = App.Path & "\Satellite.ini"
GetPrivateProfileString "ShowCursorOutline", "Yes,No", SzReturn, Cursor, Len(Cursor), StrFileName
Check3.Value = Cursor
If Cursor = 1 Then
C = True
Form1.Frame1.Visible = True
Form1.Frame2.Visible = True
Else
C = False
Form1.Frame1.Visible = False
Form1.Frame2.Visible = False
End If

End Function
