VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Configure Ftp Server"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   2655
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
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "Form5.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Host Name:"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   840
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim StrFileName
Dim SiteName As String
Dim SzReturn As String, SName As String
SName = Space(50)
StrFileName = App.Path & "\" & Text1.Text & ".ftp"
If Text1.Text <> "" Then
If Text2.Text <> "" Then
WritePrivateProfileString "Profile", "SiteName", Text1.Text, StrFileName
WritePrivateProfileString "Profile", "Address", Text2.Text, StrFileName
WritePrivateProfileString "Profile", "UserId", Text3.Text, StrFileName
WritePrivateProfileString "Profile", "PassWord", Text4.Text, StrFileName
MsgBox "This Profile has been saved!", vbInformation, App.ProductName
Form1.List2.Clear
Form1.List2.AddItem ("Remote Servers")
Form1.File2.Refresh
Form1.List2.Refresh
Form1.GetGlobals
Unload Me
Else
MsgBox ("Please enter your information!"), , App.ProductName
End If
Else
MsgBox ("Please enter your information!"), , App.ProductName
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_Change()
Command1.Enabled = True
End Sub
