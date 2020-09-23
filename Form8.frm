VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Connect to..."
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4500
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Create Log"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
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
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   1
      Top             =   120
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
      Picture         =   "Form8.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Host Name:"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FtpService As Integer
Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Dim Service As Long
If Len(Text2.Text) <= 6 Then
MsgBox ("Wrong adress!"), vbOKOnly, App.ProductName
Exit Sub
End If

Adresa = Text2.Text
ID = Text3.Text
Pass = Text4.Text
Klic = ""
session = InternetOpen(Text1.Text, INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
If session <> 0 Then
Text5.SelText = Text5.SelText & Date & ", " & Time & " *** " & UCase(Text1.Text) & " ***" & vbCrLf & Time & " > Connecting to: " & Adresa & "..." & vbCrLf
Text5.SelText = Text5.SelText & Time & " > Need UserId and PassWord" & vbCrLf
server = InternetConnect(session, Adresa, Port, ID, Pass, INTERNET_SERVICE_FTP, Service, &H0)

If server = 0 Then
MsgBox ("Connection to server failed!"), vbExclamation, App.ProductName
Text5.SelText = Text5.SelText & Time & " > Connection to server failed." & vbCrLf
InternetCloseHandle session
Unload Me
Screen.MousePointer = vbDefault
Exit Sub

Else
Text5.SelText = Text5.SelText & Time & " > UserId Ok" & vbCrLf & Time & " > PassWord Ok" & vbCrLf
Text5.SelText = Text5.SelText & Time & " > Connected to service, looking for host." & vbCrLf
adr = Space(260)
FtpGetCurrentDirectory server, adr, Len(adr)
adr = Left(adr, InStr(1, adr, Chr(0)) - 1)
adr = adr & IIf((Right(adr, 1) = "/"), "*.*", "/*.*")
Text5.SelText = Text5.SelText & Time & " > Connected to server." & vbCrLf
Klic = "/"
Form1.List1.BackColor = vbWhite
Form1.List1.ForeColor = vbBlack
Form1.List3.BackColor = vbWhite
Form1.List3.ForeColor = vbBlack
List
Form1.Caption = App.ProductName & " -" & " ftp://" & Adresa
End If
Else
MsgBox ("Connection to service failed!"), vbExclamation, App.ProductName
Text5.SelText = Text5.SelText & "Connection to service failed." & vbCrLf
InternetCloseHandle session
Exit Sub
End If
Screen.MousePointer = vbDefault
If Check1.Value = 1 Then
CreateLog
Else
'Nada
End If
Saved = True
Form1.StatusBar1.Panels(4).Text = Form1.Caption
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim StrFileName
Dim SzReturn As String
Dim SName As String
Dim AName As String
Dim User As String
Dim p As String
SName = Space(50)
AName = Space(50)
User = Space(50)
p = Space(50)
StrFileName = App.Path & "\" & Form1.List2.List(Form1.List2.ListIndex)
GetPrivateProfileString "Profile", "SiteName", SzReturn, SName, Len(SName), StrFileName
GetPrivateProfileString "Profile", "Address", SzReturn, AName, Len(AName), StrFileName
GetPrivateProfileString "Profile", "UserId", SzReturn, User, Len(User), StrFileName
GetPrivateProfileString "Profile", "PassWord", SzReturn, p, Len(p), StrFileName
Form8.Text1.Text = SName
Form8.Text2.Text = AName
Form8.Text3.Text = User
Form8.Text4.Text = p
End Sub

Public Sub List()
Dim hFile As Long, udtWFD As WIN32_FIND_DATA
Dim strFile As String
Dim Img As Integer, R As Integer
Dim l&
Dim sTime As SYSTEMTIME, lTime As FILETIME

If session = 0 Or server = 0 Then
MsgBox ("You are not connected to any server"), vbInformation, App.ProductName
Exit Sub
End If
Text5.SelText = Time & "  > Sending request..., wait."
Form1.List1.Clear
Text5.SelText = Time & " > Transfering data..." & vbCrLf
Text5.SelText = Time & " > Opening folder: " & Chr(34) & adr & Chr(34) & vbCrLf
hFile = FtpFindFirstFile(server, adr, udtWFD, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
If hFile Then
Form1.List1.AddItem ("/")
Do
strFile = Left(udtWFD.cFileName, InStr(1, udtWFD.cFileName, Chr(0)) - 1)
If Len(strFile) > 0 Then
If udtWFD.dwFileAttributes And vbDirectory Then
Form1.List1.AddItem (Klic & strFile & "/")
Else
Form1.List3.AddItem (strFile)
End If
End If
Loop While InternetFindNextFile(hFile, udtWFD)
End If
InternetCloseHandle hFile
Text5.SelText = Time & " > Data transfer completed succesfully." & vbCrLf
End Sub

Private Sub CreateLog()
Dim FileNum As Integer, Buffer As String
Buffer = Text5.Text
FileName = App.Path & "\" & Text1.Text & ".log"
FileNum = FreeFile
Open FileName For Append As FileNum
Print #FileNum, Buffer
Close FileNum
MsgBox ("A log file has been created and saved as " & FileName), vbOKOnly + vbInformation, App.ProductName
FileName = ""
End Sub
