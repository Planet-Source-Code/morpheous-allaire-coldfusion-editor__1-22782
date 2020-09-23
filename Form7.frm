VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form7 
   Caption         =   "Open Url"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox TxtUrl 
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
      Left            =   240
      TabIndex        =   1
      Text            =   "http://"
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton CmdGet 
      Caption         =   "&Get"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1560
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Inet1.Cancel
Screen.MousePointer = vbDefault
Unload Me
End Sub

Private Sub CmdGet_Click()
Saved = False
If IsNetConnectOnline = False Then
MsgBox "You are not connected to the internet!", vbOKOnly, App.ProductName
Unload Me
Else
If TxtUrl.Text = "http://" Then
MsgBox ("This is a malformed url!" & vbCrLf & "Operation has been canceled"), , App.ProductName
Unload Me
Else
Screen.MousePointer = vbHourglass
Form1.RichTextBox1.Text = Inet1.OpenURL(TxtUrl.Text)
Form1.RichTextBox2.Text = Inet1.OpenURL(TxtUrl.Text)
Form1.Caption = App.ProductName & " - [" & TxtUrl.Text & " ]"
Form1.StatusBar1.Panels(4).Text = Form1.Caption
Screen.MousePointer = vbDefault
Unload Me
End If
End If
End Sub

Private Sub TxtUrl_Change()
CmdGet.Enabled = True
End Sub

Private Sub TxtUrl_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
CmdGet_Click
End If
End Sub

