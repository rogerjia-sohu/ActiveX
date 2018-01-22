VERSION 5.00
Begin VB.UserControl Kx 
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ClipBehavior    =   0  '无
   ClipControls    =   0   'False
   Enabled         =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Kx.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   510
End
Attribute VB_Name = "Kx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const MEM_LOAD As Long = 1
Const TOTAL_PHYS  As Long = 2
Const AVAIL_PHYS  As Long = 3
Const TOTAL_PAGE  As Long = 4
Const AVAIL_PAGE  As Long = 5
Const TOTAL_VIRT  As Long = 6
Const AVAIL_VIRT  As Long = 7
'缺省属性值:
Const m_def_CurrentUser = "CurrentUserName"
Const m_def_MemLoad = 0
Const m_def_MemTotalPhys = 0
Const m_def_MemAvailPhys = 0
Const m_def_MemTotalPageFile = 0
Const m_def_MemAvailPageFile = 0
Const m_def_MemTotalVirtual = 0
Const m_def_MemAvailVirtual = 0
Const m_def_IsNT = 0
Const m_def_OSMajor = 0
Const m_def_OSMinor = 0
Const m_def_OSBuild = 0
'属性变量:
Dim m_CurrentUser As String
Dim m_MemLoad As Long
Dim m_MemTotalPhys As Long
Dim m_MemAvailPhys As Long
Dim m_MemTotalPageFile As Long
Dim m_MemAvailPageFile As Long
Dim m_MemTotalVirtual As Long
Dim m_MemAvailVirtual As Long
Dim m_IsNT As Long
Dim m_OSMajor As Long
Dim m_OSMinor As Long
Dim m_OSBuild As Long


'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function About(Optional ByVal hWnd As Long = 0) As Long
    About = AboutKernelX(hWnd)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function ActiveWindow(ByVal Title As String) As Long
    ActiveWindow = mdlKx.ActiveWindow(Title)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function AlphaBlendWindow(ByVal hWnd As Long, Optional ByVal Alpha As Byte = 80) As Long
    AlphaBlendWindow = mdlKx.AlphaBlendWindow(hWnd, Alpha)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function CDROM(ByVal Eject As Boolean) As Long
    CDROM = mdlKx.CDROM(Eject)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function Delay(ByVal ms As Long, Optional ByVal SleepMode As Boolean = True, Optional ByVal Alterable As Boolean = False) As Long
    Delay = mdlKx.Delay(ms, SleepMode, Alterable)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function GetScreenSize() As Long
    GetScreenSize = mdlKx.GetScreenSize
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function RunScreenSaver() As Long
    RunScreenSaver = mdlKx.RunScreenSaver
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function ShutDown() As Long
    ShutDown = mdlKx.ShutDown
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function WriteResDataToFile(ByVal ResID As String, ByVal FileName As String, ByVal cbSize As Long, ByVal Execute As Boolean, ByVal DeleteAfterUse As Boolean) As Long

End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get IsNT() As Long
Attribute IsNT.VB_MemberFlags = "400"
    m_IsNT = mdlKx.IsNT
    IsNT = m_IsNT
End Property

Public Property Let IsNT(ByVal New_IsNT As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_IsNT = New_IsNT
    PropertyChanged "IsNT"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get OSMajor() As Long
Attribute OSMajor.VB_MemberFlags = "400"
    m_OSMajor = mdlKx.OSMajor
    OSMajor = m_OSMajor
End Property

Public Property Let OSMajor(ByVal New_OSMajor As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_OSMajor = New_OSMajor
    PropertyChanged "OSMajor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get OSMinor() As Long
Attribute OSMinor.VB_MemberFlags = "400"
    m_OSMinor = mdlKx.OSMinor
    OSMinor = m_OSMinor
End Property

Public Property Let OSMinor(ByVal New_OSMinor As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_OSMinor = New_OSMinor
    PropertyChanged "OSMinor"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get OSBuild() As Long
Attribute OSBuild.VB_MemberFlags = "400"
    m_OSBuild = mdlKx.OSBuild
    If m_OSBuild = 67766446 Then m_OSBuild = 2222
    OSBuild = m_OSBuild
End Property

Public Property Let OSBuild(ByVal New_OSBuild As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_OSBuild = New_OSBuild
    PropertyChanged "OSBuild"
End Property


'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_IsNT = m_def_IsNT
    m_OSMajor = m_def_OSMajor
    m_OSMinor = m_def_OSMinor
    m_OSBuild = m_def_OSBuild
    m_MemLoad = m_def_MemLoad
    m_MemTotalPhys = m_def_MemTotalPhys
    m_MemAvailPhys = m_def_MemAvailPhys
    m_MemTotalPageFile = m_def_MemTotalPageFile
    m_MemAvailPageFile = m_def_MemAvailPageFile
    m_MemTotalVirtual = m_def_MemTotalVirtual
    m_MemAvailVirtual = m_def_MemAvailVirtual
   m_CurrentUser = m_def_CurrentUser
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_IsNT = PropBag.ReadProperty("IsNT", m_def_IsNT)
    m_OSMajor = PropBag.ReadProperty("OSMajor", m_def_OSMajor)
    m_OSMinor = PropBag.ReadProperty("OSMinor", m_def_OSMinor)
    m_OSBuild = PropBag.ReadProperty("OSBuild", m_def_OSBuild)
    m_MemLoad = PropBag.ReadProperty("MemLoad", m_def_MemLoad)
    m_MemTotalPhys = PropBag.ReadProperty("MemTotalPhys", m_def_MemTotalPhys)
    m_MemAvailPhys = PropBag.ReadProperty("MemAvailPhys", m_def_MemAvailPhys)
    m_MemTotalPageFile = PropBag.ReadProperty("MemTotalPageFile", m_def_MemTotalPageFile)
    m_MemAvailPageFile = PropBag.ReadProperty("MemAvailPageFile", m_def_MemAvailPageFile)
    m_MemTotalVirtual = PropBag.ReadProperty("MemTotalVirtual", m_def_MemTotalVirtual)
    m_MemAvailVirtual = PropBag.ReadProperty("MemAvailVirtual", m_def_MemAvailVirtual)
'    Set m_Mem = PropBag.ReadProperty("Mem", Nothing)
'   m_CurrentUser = PropBag.ReadProperty("CurrentUser", m_def_CurrentUser)
   m_CurrentUser = PropBag.ReadProperty("CurrentUser", m_def_CurrentUser)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("IsNT", m_IsNT, m_def_IsNT)
    Call PropBag.WriteProperty("OSMajor", m_OSMajor, m_def_OSMajor)
    Call PropBag.WriteProperty("OSMinor", m_OSMinor, m_def_OSMinor)
    Call PropBag.WriteProperty("OSBuild", m_OSBuild, m_def_OSBuild)
    Call PropBag.WriteProperty("MemLoad", m_MemLoad, m_def_MemLoad)
    Call PropBag.WriteProperty("MemTotalPhys", m_MemTotalPhys, m_def_MemTotalPhys)
    Call PropBag.WriteProperty("MemAvailPhys", m_MemAvailPhys, m_def_MemAvailPhys)
    Call PropBag.WriteProperty("MemTotalPageFile", m_MemTotalPageFile, m_def_MemTotalPageFile)
    Call PropBag.WriteProperty("MemAvailPageFile", m_MemAvailPageFile, m_def_MemAvailPageFile)
    Call PropBag.WriteProperty("MemTotalVirtual", m_MemTotalVirtual, m_def_MemTotalVirtual)
    Call PropBag.WriteProperty("MemAvailVirtual", m_MemAvailVirtual, m_def_MemAvailVirtual)
'    Call PropBag.WriteProperty("Mem", m_Mem, Nothing)
'   Call PropBag.WriteProperty("CurrentUser", m_CurrentUser, m_def_CurrentUser)
   Call PropBag.WriteProperty("CurrentUser", m_CurrentUser, m_def_CurrentUser)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function HiWord(ByVal dwVal As Long) As Long
    HiWord = mdlKx.HiWord(dwVal)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function LoWord(ByVal dwVal As Long) As Long
    LoWord = mdlKx.LoWord(dwVal)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function CloseApp(ByVal ClassName As String, ByVal WndTitle As String) As Long
    CloseApp = mdlKx.CloseApp(ClassName, WndTitle)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function kShell(ByVal AppName As String) As Long
    kShell = mdlKx.kShell(AppName)
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function StopPlay() As Long
    StopPlay = mdlKx.StopPlay
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function PlayAudio() As Long
    PlayAudio = mdlKx.PlayAudio
End Function

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemLoad() As Long
Attribute MemLoad.VB_MemberFlags = "400"
    m_MemLoad = mdlMem.GlobalMemInfo(MEM_LOAD)
    MemLoad = m_MemLoad
End Property

Public Property Let MemLoad(ByVal New_MemLoad As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemLoad = New_MemLoad
    PropertyChanged "MemLoad"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemTotalPhys() As Long
Attribute MemTotalPhys.VB_MemberFlags = "400"
    m_MemTotalPhys = mdlMem.GlobalMemInfo(TOTAL_PHYS)
    MemTotalPhys = m_MemTotalPhys
End Property

Public Property Let MemTotalPhys(ByVal New_MemTotalPhys As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemTotalPhys = New_MemTotalPhys
    PropertyChanged "MemTotalPhys"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemAvailPhys() As Long
Attribute MemAvailPhys.VB_MemberFlags = "400"
    m_MemAvailPhys = mdlMem.GlobalMemInfo(AVAIL_PHYS)
    MemAvailPhys = m_MemAvailPhys
End Property

Public Property Let MemAvailPhys(ByVal New_MemAvailPhys As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemAvailPhys = New_MemAvailPhys
    PropertyChanged "MemAvailPhys"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemTotalPageFile() As Long
Attribute MemTotalPageFile.VB_MemberFlags = "400"
    m_MemTotalPageFile = mdlMem.GlobalMemInfo(TOTAL_PAGE)
    MemTotalPageFile = m_MemTotalPageFile
End Property

Public Property Let MemTotalPageFile(ByVal New_MemTotalPageFile As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemTotalPageFile = New_MemTotalPageFile
    PropertyChanged "MemTotalPageFile"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemAvailPageFile() As Long
Attribute MemAvailPageFile.VB_MemberFlags = "400"
    m_MemAvailPageFile = mdlMem.GlobalMemInfo(AVAIL_PAGE)
    MemAvailPageFile = m_MemAvailPageFile
End Property

Public Property Let MemAvailPageFile(ByVal New_MemAvailPageFile As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemAvailPageFile = New_MemAvailPageFile
    PropertyChanged "MemAvailPageFile"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemTotalVirtual() As Long
Attribute MemTotalVirtual.VB_MemberFlags = "400"
    m_MemTotalVirtual = mdlMem.GlobalMemInfo(TOTAL_VIRT)
    MemTotalVirtual = m_MemTotalVirtual
End Property

Public Property Let MemTotalVirtual(ByVal New_MemTotalVirtual As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemTotalVirtual = New_MemTotalVirtual
    PropertyChanged "MemTotalVirtual"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemAvailVirtual() As Long
Attribute MemAvailVirtual.VB_MemberFlags = "400"
    m_MemAvailVirtual = mdlMem.GlobalMemInfo(AVAIL_VIRT)
    MemAvailVirtual = m_MemAvailVirtual
End Property

Public Property Let MemAvailVirtual(ByVal New_MemAvailVirtual As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemAvailVirtual = New_MemAvailVirtual
    PropertyChanged "MemAvailVirtual"
End Property
'注意！不要删除或修改下列被注释的行！
'MemberInfo=8
Public Function BugReport() As Long
   BugReport = mdlKx.BugReport
End Function
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,1,2,CurrentUserName
Public Property Get CurrentUser() As String
Attribute CurrentUser.VB_MemberFlags = "400"
   CurrentUser = mdlKx.GetUserName(m_CurrentUser, 256)
   CurrentUser = m_CurrentUser
End Property

Public Property Let CurrentUser(ByVal New_CurrentUser As String)
   If Ambient.UserMode = False Then Err.Raise 387
   If Ambient.UserMode Then Err.Raise 382
   m_CurrentUser = New_CurrentUser
   PropertyChanged "CurrentUser"
End Property

