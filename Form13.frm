VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Satellite"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   25
      Left            =   7200
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6720
      Top             =   1080
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   930
      Left            =   120
      Picture         =   "Form13.frx":0000
      ScaleHeight     =   870
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   0
      Width           =   7515
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6240
      Top             =   1080
   End
   Begin VB.Image Image1 
      Height          =   15
      Left            =   0
      Picture         =   "Form13.frx":170D
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   7680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   645
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Label1.Caption = "Loading Please Wait..."
Me.Caption = App.ProductName & " v" & App.Major & " ." & App.Minor
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form1.Show
Unload Me
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Screen.MousePointer = vbDefault
End Sub

Private Sub Timer2_Timer()
Label1.Caption = "Thank you for using Satellite.."
End Sub

Private Sub Timer3_Timer()
If Image1.Left > Form1.Width Then
Image1.Left = Form1.Left - (Image1.Width * 1)
Image1.Left = Image1.Left + 200
Else
Image1.Left = Image1.Left + 200
End If
End Sub
