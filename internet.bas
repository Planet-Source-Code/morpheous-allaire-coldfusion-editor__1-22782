Attribute VB_Name = "isconnected"
Option Explicit
Public Declare Function InternetGetConnectedState _
Lib "wininet.dll" (ByRef lpdwFlags As Long, _
ByVal dwReserved As Long) As Long
       
Public Function IsNetConnectOnline() As Boolean
IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
End Function
    

