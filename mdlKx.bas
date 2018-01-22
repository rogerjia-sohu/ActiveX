Attribute VB_Name = "mdlKx"
Option Explicit

Declare Function AboutKernelX Lib "KernelX.dll" (ByVal hWnd As Long) As Long
Declare Function ActiveWindow Lib "KernelX.dll" (ByVal WinTitleh As String) As Long
Declare Function AlphaBlendWindow Lib "KernelX.dll" (ByVal hWnd As Long, Optional ByVal Alpha As Byte) As Long
Declare Function BugReport Lib "KernelX.dll" () As Long
Declare Function CDROM Lib "KernelX.dll" (ByVal Eject As Boolean) As Long
Declare Function CloseApp Lib "KernelX.dll" (ByVal ClassName As String, ByVal WndTitle As String) As Long

'_ConsoleRead@8 'not availabel in VB
'_ConsoleWrite@4 'not availabel in VB

Declare Function kShell Lib "KernelX.dll" (ByVal AppName As String) As Long


Declare Function Delay Lib "KernelX.dll" ( _
    ByVal ms As Long, _
    ByVal SleepMode As Boolean, _
    ByVal Alterable As Boolean) As Long

Declare Function GetScreenSize Lib "KernelX.dll" () As Long
Declare Function HiWord Lib "KernelX.dll" (ByVal dwVal As Long) As Long
Declare Function LoWord Lib "KernelX.dll" (ByVal dwVal As Long) As Long

Declare Function IsNT Lib "KernelX.dll" () As Long
Declare Function OSBuild Lib "KernelX.dll" () As Long
Declare Function OSMajor Lib "KernelX.dll" () As Long
Declare Function OSMinor Lib "KernelX.dll" () As Long
Declare Function PlayAudio Lib "KernelX.dll" () As Long
Declare Function RunScreenSaver Lib "KernelX.dll" () As Long
Declare Function ShutDown Lib "KernelX.dll" () As Long
Declare Function StopPlay Lib "KernelX.dll" () As Long

Declare Function WriteResDataToFile Lib "KernelX.dll" ( _
    ByVal ResID As String, _
    ByVal FileName As String, _
    ByVal cbSize As Long, _
    ByVal Execute As Boolean, _
    ByVal DeleteAfterUse As Boolean) As Long


Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


