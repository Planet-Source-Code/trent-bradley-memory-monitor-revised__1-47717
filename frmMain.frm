VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Memory Monitor"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5160
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar pgbVirtMem 
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pgbPhysMem 
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblVirtUsed 
      Caption         =   "Free Virtual Memory: "
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblPhysUsed 
      Caption         =   "Free Physical Memory: "
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblAvailVirtual 
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblTotalVirtual 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblAvailPage 
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblTotalPage 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblAvailPhys 
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label lblTotalPhys 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
    End With
    memInfo.dwLength = Len(memInfo)
    Call GlobalMemoryStatus(memInfo)
    pgbPhysMem.Min = 0
    pgbPhysMem.Max = memInfo.dwTotalPhys
    pgbVirtMem.Min = 0
    pgbVirtMem.Max = memInfo.dwTotalVirtual
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result As Long
    Dim msg As Long
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_LBUTTONUP
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_LBUTTONDBLCLK
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_RBUTTONUP
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
    End Select
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Shell_NotifyIcon NIM_ADD, nid
    ElseIf Me.WindowState = vbNormal Then
        Me.Show
        Shell_NotifyIcon NIM_DELETE, nid
        frmMain.Width = 7320
        frmMain.Height = 2160
    End If
End Sub
Private Sub Form_Terminate()
    Shell_NotifyIcon NIM_DELETE, nid
    End
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
    End
End Sub
Private Sub Timer1_Timer()
    Dim PhysUsed
    Dim VirtUsed
    Call GlobalMemoryStatus(memInfo)
    If memInfo.dwAvailPhys = 0 Then
        MsgBox ("Your system is out of memory. It is advised that you close one or more applications to free up memory or your computer may crash.")
        PhysUsed = memInfo.dwTotalPhys - memInfo.dwAvailPhys
        pgbPhysMem.Value = PhysUsed
        lblPhysUsed.Caption = "Physical Memory Usage: " & Format(PhysUsed / memInfo.dwTotalPhys, "0.00%")
        VirtUsed = memInfo.dwTotalVirtual - memInfo.dwAvailVirtual
        pgbVirtMem.Value = VirtUsed
        lblVirtUsed.Caption = "Virtual Memory Usage: " & Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
        lblTotalPhys.Caption = "Total physical memory (RAM): " & memInfo.dwTotalPhys / 1024 & " KB"
        lblAvailPhys.Caption = "Free physical memory (RAM): " & memInfo.dwAvailPhys / 1024 & " KB"
        lblTotalPage.Caption = "Total KB in current paging file: " & memInfo.dwTotalPageFile / 1024
        lblAvailPage.Caption = "Free KB in current paging file: " & memInfo.dwAvailPageFile / 1024
        lblTotalVirtual.Caption = "Total virtual memory: " & memInfo.dwTotalVirtual / 1024 & " KB"
        lblAvailVirtual.Caption = "Free virtual memory: " & memInfo.dwAvailVirtual / 1024 & " KB"
        nid.szTip = "Physical Memory Usage: " & Format(PhysUsed / memInfo.dwTotalPhys, "0.00%") & " - " & "Virtual Memory Usage: " & Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
    Else
        PhysUsed = memInfo.dwTotalPhys - memInfo.dwAvailPhys
        pgbPhysMem.Value = PhysUsed
        lblPhysUsed.Caption = "Physical Memory Usage: " & Format(PhysUsed / memInfo.dwTotalPhys, "0.00%")
        VirtUsed = memInfo.dwTotalVirtual - memInfo.dwAvailVirtual
        pgbVirtMem.Value = VirtUsed
        lblVirtUsed.Caption = "Virtual Memory Usage: " & Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
        lblTotalPhys.Caption = "Total physical memory (RAM): " & memInfo.dwTotalPhys / 1024 & " KB"
        lblAvailPhys.Caption = "Free physical memory (RAM): " & memInfo.dwAvailPhys / 1024 & " KB"
        lblTotalPage.Caption = "Total KB in current paging file: " & memInfo.dwTotalPageFile / 1024
        lblAvailPage.Caption = "Free KB in current paging file: " & memInfo.dwAvailPageFile / 1024
        lblTotalVirtual.Caption = "Total virtual memory: " & memInfo.dwTotalVirtual / 1024 & " KB"
        lblAvailVirtual.Caption = "Free virtual memory: " & memInfo.dwAvailVirtual / 1024 & " KB"
        nid.szTip = "Physical Memory Usage: " & Format(PhysUsed / memInfo.dwTotalPhys, "0.00%") & " - " & "Virtual Memory Usage: " & Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
    End If
End Sub
