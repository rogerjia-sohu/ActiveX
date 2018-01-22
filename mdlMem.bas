Attribute VB_Name = "mdlMem"
Option Explicit
Const MEM_LOAD As Long = 1
Const TOTAL_PHYS  As Long = 2
Const AVAIL_PHYS  As Long = 3
Const TOTAL_PAGE  As Long = 4
Const AVAIL_PAGE  As Long = 5
Const TOTAL_VIRT  As Long = 6
Const AVAIL_VIRT  As Long = 7

Declare Function GlobalMemInfo Lib "KernelX.dll" (InfoType As Long) As Long

