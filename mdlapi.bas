Attribute VB_Name = "mdlapi"
'Author : Erwan L.
'email:erwan.l@free.fr

Option Explicit


Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

    Public Declare Function GetVersionEx Lib "kernel32" _
Alias "GetVersionExA" _
(lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Public Declare Sub CopyMemory Lib "kernel32" _
'        Alias "RtlMoveMemory" (ByRef hpvDest As Any, _
'        ByVal hpvSource As Long, ByVal cbCopy As Long)
        
       Declare Sub CopyMemory_any Lib "kernel32" _
        Alias "RtlMoveMemory" (ByRef hpvDest As Any, _
          ByRef hpvSource As Any, ByVal cbCopy As Long)

