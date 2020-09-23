VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Filter"
   ClientHeight    =   540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2400
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form10.frx":08CA
      Left            =   120
      List            =   "Form10.frx":08EF
      TabIndex        =   0
      Text            =   "*.*"
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.File1.Pattern = Combo1.Text
Unload Me
End Sub

