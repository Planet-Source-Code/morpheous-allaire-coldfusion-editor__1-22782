Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_UNDO = &HC7

'API Constants
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const LB_ITEMFROMPOINT = &H1A9

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Declare Function CreateProcess Lib "kernel32" _
   Alias "CreateProcessA" _
  (ByVal lpAppName As String, _
   ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, _
   ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, _
   ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, _
   ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long
     
Public Declare Function CloseHandle Lib "kernel32" _
  (ByVal hObject As Long) As Long

Public Declare Function FindExecutable Lib "shell32.dll" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Public Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal nSize As Long, _
   ByVal lpBuffer As String) As Long
   Public Const CREATE_NEW_CONSOLE As Long = &H10
Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const INFINITE As Long = -1
Public Const STARTF_USESHOWWINDOW As Long = &H1
Public Const SW_SHOWNORMAL As Long = 1

'Public Const MAX_PATH As Long = 260
Public Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Public Const ERROR_FILE_NOT_FOUND As Long = 2
Public Const ERROR_PATH_NOT_FOUND As Long = 3
Public Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Public Const ERROR_BAD_FORMAT As Long = 11

Public Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Public Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type

Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) _
As Long
Public A As POINTAPI
Public Type POINTAPI
X As Long
Y As Long
End Type

Public Function SetWinPos(iPos As Integer, lHWnd As Long) As Boolean
    Dim lWinPos As Long 'A variable to hold the
                        'the value of API window
                        'position constant
    Dim l As Long
    
    'Use a SELECT CASE to set the value of the of
    'the API Window constant
    
    Select Case iPos
        'The window is to set to it regular position
        Case 0
            lWinPos = HWND_NOTOPMOST
        'Set the window always on top
        Case 1
            lWinPos = HWND_TOPMOST
        'You have a bad value, leave the function
        Case Else
            Exit Function
    End Select
    
    'Run the API SetWindowPos function
    If SetWindowPos(lHWnd, lWinPos, 0, 0, 0, 0, SWP_NOMOVE _
                                    + SWP_NOSIZE) Then
        'If the function is greater than 0 (FALSE) then
        'the operation was successful. Return a True for
        'to indicate such.
        SetWinPos = True
    End If
End Function
