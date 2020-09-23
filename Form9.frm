VERSION 5.00
Begin VB.Form Form9 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data transfer"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Download"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Bytes Read"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   ":"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const sReadBuffer = 1024
Private hFile As Long
Public Saved As Boolean

Private Sub Command1_Click()
If hFile <> 0 Then
InternetCloseHandle hFile
MsgBox ("Opertion Canceled."), vbInformation, App.ProductName
Unload Me
End If
End Sub

Private Sub Download()
Dim sBuffer As String
Dim FileData As String
Dim Ret As Long, sAllBytes As Long, z As Long
Dim I As Integer, FF As Integer
Dim Kam As String, Ode As String
Dim FS As Long, StartT As Long, T As Long, Cnt As Long, p As Long
Dim spRate As Single
Dim SentBytes As Long
Ode = Klic & List1.List(I)
Kam = App.Path & "\" & List1.List(I)
hFile = FtpOpenFile(server, Ode, GENERIC_READ, FTP_TRANSFER_TYPE_BINARY, 0)
If hFile = 0 Then
MsgBox ("Can't open file path!"), vbExclamation, App.ProductName & hFile
Exit Sub
End If
sBuffer = Space(sReadBuffer)
FileData = ""
SentBytes = 0
StartT = GetTickCount
Do
InternetReadFile hFile, sBuffer, sReadBuffer, Ret

If Ret <> sReadBuffer Then
sBuffer = Left$(sBuffer, Ret)
End If
FileData = FileData + sBuffer
SentBytes = SentBytes + Ret
sAllBytes = sAllBytes + Ret
If SentBytes <> 0 Then
T = GetTickCount - StartT
If T <> 0 Then
Label1.Caption = ":" & sAllBytes
End If
End If
Loop Until Ret <> sReadBuffer
FF = FreeFile
Open Kam For Binary As #FF
Put #FF, , FileData
Close #FF
p = T / 1000
InternetCloseHandle hFile
Form1.RichTextBox1.LoadFile (App.Path & "\" & List1.List(I))
Form1.RichTextBox2.LoadFile (App.Path & "\" & List1.List(I))
Form1.Caption = App.ProductName & " - " & Adresa & "\" & Form1.List3.List(Form1.List3.ListIndex)
Form1.StatusBar1.Panels(4).Text = Form1.Caption
Form1.List4.Clear
Form1.List4.AddItem (List1.List(I))
Saved = False
Unload Me
End Sub

Private Sub Command3_Click()
Screen.MousePointer = vbHourglass
Download
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
List1.Clear
End Sub
