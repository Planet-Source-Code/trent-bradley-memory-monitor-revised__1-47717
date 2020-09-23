Attribute VB_Name = "Memory"
Option Explicit
Public Type MemoryStatus
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MemoryStatus)
Public memInfo As MemoryStatus
