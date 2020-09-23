VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "About Satellite"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&System Info...."
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "This product licensed to:"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   1740
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0&
Private Const READ_CONTROL = &H20000
Private Const KEY_NOTIFY = &H10
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private SysInfoPath As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Sub Command1_Click()
Unload Me
End Sub
Public Property Get UserName() As Variant
Dim sBuffer As String
Dim lSize As Long
sBuffer = Space$(255)
lSize = Len(sBuffer)
Call GetUserName(sBuffer, lSize)
UserName = Left$(sBuffer, lSize)
End Property

Public Property Get ThreadID() As Variant
ThreadID = GetCurrentThreadId
End Property
Public Property Get ProcessID() As Variant
ProcessID = GetCurrentProcessId
End Property
Private Function FindSysInfoPath() As String
Dim buf As String
Dim buf_len As Long
Dim info_key As Long
Dim value_type As Long
Dim key_size As Long
If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
"software\microsoft\shared tools\msinfo", _
0, KEY_READ, info_key) = ERROR_SUCCESS _
Then
buf = Space$(256)
buf_len = Len(buf)
If RegQueryValueEx(info_key, "Path", _
0, value_type, buf, buf_len) _
= ERROR_SUCCESS _
Then
FindSysInfoPath = Left$(buf, buf_len)
End If
End If
RegCloseKey info_key
End Function

Private Sub Command2_Click()
SysInfoPath = FindSysInfoPath
Shell SysInfoPath, vbNormalFocus
Unload Me
End Sub

Private Sub Form_Load()
Label2.Caption = "Satellite Editor" & vbCrLf & "Version " & App.Major & "." & App.Minor & "." & App.Revision
Text1.Text = UserName
End Sub

