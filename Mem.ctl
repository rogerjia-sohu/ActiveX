VERSION 5.00
Begin VB.UserControl Mem 
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   480
End
Attribute VB_Name = "Mem"
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
Const m_def_MemoryLoad = 0
Const m_def_TotalPhys = 0
Const m_def_AvailPhys = 0
Const m_def_TotalPageFile = 0
Const m_def_AvailPageFile = 0
Const m_def_TotalVirtual = 0
Const m_def_AvailVirtual = 0
'属性变量:
Dim m_MemoryLoad As Long
Dim m_TotalPhys As Long
Dim m_AvailPhys As Long
Dim m_TotalPageFile As Long
Dim m_AvailPageFile As Long
Dim m_TotalVirtual As Long
Dim m_AvailVirtual As Long



'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get MemoryLoad() As Long
Attribute MemoryLoad.VB_MemberFlags = "400"
    m_MemoryLoad = mdlMem.GlobalMemInfo(MEM_LOAD)
    MemoryLoad = m_MemoryLoad
End Property

Public Property Let MemoryLoad(ByVal New_MemoryLoad As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_MemoryLoad = New_MemoryLoad
    PropertyChanged "MemoryLoad"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get TotalPhys() As Long
Attribute TotalPhys.VB_MemberFlags = "400"
    m_TotalPhys = mdlMem.GlobalMemInfo(TOTAL_PHYS)
    TotalPhys = m_TotalPhys
End Property

Public Property Let TotalPhys(ByVal New_TotalPhys As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_TotalPhys = New_TotalPhys
    PropertyChanged "TotalPhys"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get AvailPhys() As Long
Attribute AvailPhys.VB_MemberFlags = "400"
    m_AvailPhys = mdlMem.GlobalMemInfo(AVAIL_PHYS)
    AvailPhys = m_AvailPhys
End Property

Public Property Let AvailPhys(ByVal New_AvailPhys As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_AvailPhys = New_AvailPhys
    PropertyChanged "AvailPhys"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get TotalPageFile() As Long
Attribute TotalPageFile.VB_MemberFlags = "400"
    m_TotalPageFile = mdlMem.GlobalMemInfo(TOTAL_PAGE)
    TotalPageFile = m_TotalPageFile
End Property

Public Property Let TotalPageFile(ByVal New_TotalPageFile As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_TotalPageFile = New_TotalPageFile
    PropertyChanged "TotalPageFile"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get AvailPageFile() As Long
Attribute AvailPageFile.VB_MemberFlags = "400"
    m_AvailPageFile = mdlMem.GlobalMemInfo(AVAIL_PAGE)
    AvailPageFile = m_AvailPageFile
End Property

Public Property Let AvailPageFile(ByVal New_AvailPageFile As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_AvailPageFile = New_AvailPageFile
    PropertyChanged "AvailPageFile"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get TotalVirtual() As Long
Attribute TotalVirtual.VB_MemberFlags = "400"
    m_TotalVirtual = mdlMem.GlobalMemInfo(TOTAL_VIRT)
    TotalVirtual = m_TotalVirtual
End Property

Public Property Let TotalVirtual(ByVal New_TotalVirtual As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_TotalVirtual = New_TotalVirtual
    PropertyChanged "TotalVirtual"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,1,2,0
Public Property Get AvailVirtual() As Long
Attribute AvailVirtual.VB_MemberFlags = "400"
    m_AvailVirtual = mdlMem.GlobalMemInfo(AVAIL_VIRT)
    AvailVirtual = m_AvailVirtual
End Property

Public Property Let AvailVirtual(ByVal New_AvailVirtual As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_AvailVirtual = New_AvailVirtual
    PropertyChanged "AvailVirtual"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_MemoryLoad = m_def_MemoryLoad
    m_TotalPhys = m_def_TotalPhys
    m_AvailPhys = m_def_AvailPhys
    m_TotalPageFile = m_def_TotalPageFile
    m_AvailPageFile = m_def_AvailPageFile
    m_TotalVirtual = m_def_TotalVirtual
    m_AvailVirtual = m_def_AvailVirtual
    
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_MemoryLoad = PropBag.ReadProperty("MemoryLoad", m_def_MemoryLoad)
    m_TotalPhys = PropBag.ReadProperty("TotalPhys", m_def_TotalPhys)
    m_AvailPhys = PropBag.ReadProperty("AvailPhys", m_def_AvailPhys)
    m_TotalPageFile = PropBag.ReadProperty("TotalPageFile", m_def_TotalPageFile)
    m_AvailPageFile = PropBag.ReadProperty("AvailPageFile", m_def_AvailPageFile)
    m_TotalVirtual = PropBag.ReadProperty("TotalVirtual", m_def_TotalVirtual)
    m_AvailVirtual = PropBag.ReadProperty("AvailVirtual", m_def_AvailVirtual)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MemoryLoad", m_MemoryLoad, m_def_MemoryLoad)
    Call PropBag.WriteProperty("TotalPhys", m_TotalPhys, m_def_TotalPhys)
    Call PropBag.WriteProperty("AvailPhys", m_AvailPhys, m_def_AvailPhys)
    Call PropBag.WriteProperty("TotalPageFile", m_TotalPageFile, m_def_TotalPageFile)
    Call PropBag.WriteProperty("AvailPageFile", m_AvailPageFile, m_def_AvailPageFile)
    Call PropBag.WriteProperty("TotalVirtual", m_TotalVirtual, m_def_TotalVirtual)
    Call PropBag.WriteProperty("AvailVirtual", m_AvailVirtual, m_def_AvailVirtual)
End Sub

