VERSION 5.00
Object = "{CF9ED78E-94F5-11D6-8731-DF4DA72C6754}#49.0#0"; "KernelX.ocx"
Begin VB.Form frmKxHost 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   3675
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Current User is..."
      Top             =   1200
      Width           =   2295
   End
   Begin KernelX.Kx Kx1 
      Left            =   1560
      Top             =   3720
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton Command9 
      Caption         =   "BugReport"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "StopPlay"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PlayCD"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "kShell"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Screen Saver"
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
      Begin VB.CommandButton Command2 
         Caption         =   "Execute"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Eject CD"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "hide/show"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1000
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Don't press ""ShutDown"""
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ShutDown"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get it"
      Height          =   320
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Press ""Get it"""
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmKxHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GlobalMemInfo Lib "KernelX.dll" (InfoType As Long) As Long
'Const MEM_LOAD As Long = 1
'Const TOTAL_PHYS  As Long = 2
'Const AVAIL_PHYS  As Long = 3
'Const TOTAL_PAGE  As Long = 4
'Const AVAIL_PAGE  As Long = 5
'Const TOTAL_VIRT  As Long = 6
'Const AVAIL_VIRT  As Long = 7


Private Sub Command1_Click()
Text1.Text = "Screen Size: " & _
                 Kx1.HiWord(Kx1.GetScreenSize) & " x " & _
                 Kx1.LoWord(Kx1.GetScreenSize) & " " & _
                 "pixel"
    
Dim winver As String
winver = "WinVer: Win"
winver = winver & IIf(Kx1.IsNT, "NT", IIf(Kx1.OSMinor > 1, "98", "95")) & " "
    
    Text2 = winver & _
          Kx1.OSMajor & "." & _
          Kx1.OSMinor & "." & _
          Kx1.OSBuild

Text3 = Left$(Text3, Len(Text3) - 3) & Space$(1) & Kx1.CurrentUser

Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    Kx1.CDROM Check1.Value
End Sub

Private Sub Command3_Click()
    If MsgBox("shutdown?  Now works on Win98/me/nt/2k/xp!", vbYesNo + vbQuestion, "confirm") = vbYes Then Kx1.ShutDown
End Sub

Private Sub Command4_Click()
    Me.WindowState = vbMinimized
    Me.Visible = Not Me.Visible
    Kx1.Delay 1500
    Me.Visible = Not Me.Visible
    Me.WindowState = vbNormal
End Sub

Private Sub Command5_Click()
    Kx1.RunScreenSaver
End Sub

Private Sub Command6_Click()
    Kx1.kShell "calc"
    MsgBox "kShell Ended!"
End Sub


Private Sub Command7_Click()
    Kx1.PlayAudio
End Sub

Private Sub Command8_Click()
    Kx1.StopPlay
End Sub

Private Sub Command9_Click()
    Kx1.BugReport
End Sub
