Attribute VB_Name = "modAPI"
Option Explicit
'API function to obtain amount of total and free RAM
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'Type for calling GlobalMemoryStatus
Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

'Variable for calling GlobalMemoryStatus
Global MemoryInfo As MEMORYSTATUS

'Variables to calculate RAM in MB's
Global intTotal, intFree As Integer

'Variables used in ShowRAMInfo
Global strTotal, strFree As String
