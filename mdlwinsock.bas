Attribute VB_Name = "mdlwinsock"
'Author : Erwan L.
'email:erwan.l@free.fr

Option Explicit

Public cnt As Long

Public Const WM_USER = &H400
Public Const WINSOCKMSG = WM_USER + 1

Public Const SIO_RCVALL = &H98000001
Public Const SO_RCVTIMEO = &H1006
Public Const AF_INET = 2
Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1
Public Const FD_READ = &H1&
Public Const FD_WRITE = &H2&
Public Const FD_CONNECT = &H10&
Public Const FD_CLOSE = &H20&
Public Const PF_INET = 2
Public Const SOCK_STREAM = 1
Public Const SOCK_RAW = 3
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_IP = 0
Public Const GWL_WNDPROC = (-4)

Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1
Public Const INADDR_NONE = &HFFFF
Public Const SOL_SOCKET = &HFFFF&
Public Const SO_LINGER = &H80&
Public Const hostent_size = 16
Public Const sockaddr_size = 16
Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

Type ipheader
    ip_verlen As Byte
    ip_tos As Byte
    ip_totallength As Integer
    ip_id As Integer
    ip_offset As Integer
    ip_ttl As Byte
    ip_protocol As Byte
    ip_checksum As Integer
    ip_srcaddr As Long
    ip_destaddr As Long
End Type

Type tcpheader
    src_portno As Integer
    dst_portno As Integer
    Sequenceno As Long
    Acknowledgeno As Long
    DataOffset As Byte
    flag As Byte
    Windows As Integer
    checksum As Integer
    UrgentPointer As Integer
End Type


Type udpheader
    src_portno As Integer
    dst_portno As Integer
    udp_length As Integer
    udp_checksum As Integer
End Type

Private Const SIO_GET_INTERFACE_LIST = &H4004747F

  Type sockaddr_gen
    AddressIn As sockaddr
    filler(0 To 7) As Byte
  End Type
  
  Type INTERFACE_INFO
iiFlags As Long     ' Interface flags
iiAddress As sockaddr_gen     ' Interface address
iiBroadcastAddress As sockaddr_gen     ' Broadcast address
iiNetmask As sockaddr_gen     ' Network mask
  End Type

  
  Type aINTERFACE_INFO
   interfaceinfo(0 To 7) As INTERFACE_INFO
  End Type

Public Declare Function bind Lib "wsock32.dll" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
Public Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function Send Lib "wsock32.dll" Alias "send" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Public Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Public Declare Function Connect Lib "wsock32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Public Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, _
ByVal dwIoControlCode As Long, _
lpvInBuffer As Any, _
ByVal cbInBuffer As Long, _
lpvOutBuffer As Any, _
ByVal cbOutBuffer As Long, _
lpcbBytesReturned As Long, _
lpOverlapped As Long, _
lpCompletionRoutine As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
          

Public saZero As sockaddr
Public WSAStartedUp As Boolean, Obj As TextBox
Public PrevProc As Long, lSocket As Long
'subclassing function
Public Sub HookForm(F As Form)
    PrevProc = SetWindowLong(F.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHookForm(F As Form)
    If PrevProc <> 0 Then
        SetWindowLong F.hwnd, GWL_WNDPROC, PrevProc
        PrevProc = 0
    End If
End Sub
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Debug.Print uMsg
    If uMsg = WINSOCKMSG Then
        ProcessMessage wParam, lParam
    Else
    'If cGetInputState() <> 0 Then
        
        WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
        
    'End If
    End If
End Function
Sub display1(readbuffer() As Byte)
'texbox -> scrollbars=3 & multiline=true
On Error GoTo errhand
Dim ip_header As ipheader
Dim tcp_header As tcpheader
Dim udp_header As udpheader
Dim lvitem
CopyMemory_any ip_header, readbuffer(0), Len(ip_header)
                   cnt = cnt + 1
                    Obj.Parent.Label2 = cnt
                    
                    Obj.Text = Obj.Text & ntohs(ip_header.ip_totallength) & " bytes" & vbCrLf
                    Obj.Text = Obj.Text & getascip(ip_header.ip_srcaddr) & "->" & getascip(ip_header.ip_destaddr) & vbCrLf
                    
                    
                    'icmp
                    If ip_header.ip_protocol = 1 Then
                        Obj.Text = Obj.Text & "ICMP" & vbCrLf
                    End If
                    'tcp
                    If ip_header.ip_protocol = 6 Then
                        Obj.Text = Obj.Text & "TCP" & vbCrLf
                        CopyMemory_any tcp_header, readbuffer(0 + 20), Len(tcp_header)
                        Obj.Text = Obj.Text & "(src port) " & ntohs(tcp_header.src_portno) & vbCrLf
                        Obj.Text = Obj.Text & "(dst port) " & ntohs(tcp_header.dst_portno) & vbCrLf
                    End If
                    'udp
                    If ip_header.ip_protocol = 17 Then
                        Obj.Text = Obj.Text & "UDP" & vbCrLf
                        CopyMemory_any udp_header, readbuffer(0 + 20), Len(udp_header)
                        Obj.Text = Obj.Text & "(src port) " & ntohs(udp_header.src_portno) & vbCrLf
                        Obj.Text = Obj.Text & "(dst port) " & ntohs(udp_header.dst_portno) & vbCrLf
                    End If
                    Obj.Text = Obj.Text & String(34, "*") & vbCrLf
Exit Sub
errhand:
MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Sub

Sub display2(readbuffer() As Byte)
On Error GoTo errhand
Dim ip_header As ipheader
Dim tcp_header As tcpheader
Dim udp_header As udpheader
Dim lvitem
CopyMemory_any ip_header, readbuffer(0), Len(ip_header)
                   cnt = cnt + 1
                    Form1.Label2 = cnt
                    
                    Set lvitem = Form1.listview1.object.ListItems.Add(, , getascip(ip_header.ip_srcaddr))
                    lvitem.Tag = readbuffer()
                    lvitem.subitems(2) = getascip(ip_header.ip_destaddr)
                    lvitem.subitems(5) = ntohs(ip_header.ip_totallength)
                    'icmp
                    If ip_header.ip_protocol = 1 Then
                        lvitem.subitems(4) = "ICMP"
                    End If
                    'tcp
                    If ip_header.ip_protocol = 6 Then
                        CopyMemory_any tcp_header, readbuffer(0 + 20), Len(tcp_header)
                        lvitem.subitems(4) = "TCP"
                        lvitem.subitems(1) = ntohs(tcp_header.src_portno)
                        lvitem.subitems(3) = ntohs(tcp_header.dst_portno)
                    End If
                    'udp
                    If ip_header.ip_protocol = 17 Then
                        CopyMemory_any udp_header, readbuffer(0 + 20), Len(udp_header)
                        lvitem.subitems(4) = "UDP"
                        lvitem.subitems(1) = ntohs(udp_header.src_portno)
                        lvitem.subitems(3) = ntohs(udp_header.dst_portno)
                    End If
Exit Sub
errhand:
MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
                    
End Sub
'our Winsock-message handler
Public Sub ProcessMessage(ByVal lFromSocket As Long, ByVal lParam As Long)
    On Error GoTo errhand
    Dim x As Long, strCommand As String
    Dim readbuffer(0 To 1499) As Byte
    Dim ip_header As ipheader
    
    Select Case lParam
        Case FD_CONNECT 'we are connected
            Debug.Print "FD_CONNECT"
        Case FD_WRITE 'we can write to our connection
            Debug.Print "FD_WRITE"
        Case FD_READ 'we have data waiting to be processed
            'start reading the data
            'Debug.Print "FD_READ"
        
        
            Do
                x = recv(lFromSocket, readbuffer(0), 1500, 0)
                If x > 0 Then
                
                display2 readbuffer()
                End If
                If x <> 1500 Then Exit Do
            
            
            Loop
        Case FD_CLOSE 'the connection is closed
            Debug.Print "FD_CLOSE"
    End Select
    Exit Sub
errhand:
    MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
    AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Sub
'the following functions are standard WinSock functions
'from the wsksock.bas-file
Public Function StartWinsock(sDescription As String) As Boolean
    Dim StartupData As WSADataType
    If Not WSAStartedUp Then
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function
Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub
'the nice part
Function ConnectSock(ByVal Host$, ByVal Port&, ByVal HWndToMsg&, ByVal Async%) As Long
    Dim s&, SelectOps&, Dummy&
    Dim RCVTIMEO As Long
    Dim sockin As sockaddr
    Dim ret As Long

    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        MsgBox "INVALID_SOCKET"
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(Host$)

    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        MsgBox "INVALID_SOCKET"
        Exit Function
    End If


    s = socket(AF_INET, SOCK_RAW, IPPROTO_IP)
    If s < 0 Then
        ConnectSock = INVALID_SOCKET
        MsgBox "INVALID_SOCKET"
        Exit Function
    End If


RCVTIMEO = 5000
ret = setsockopt(s, SOL_SOCKET, SO_RCVTIMEO, (RCVTIMEO), 4)
If ret <> 0 Then
    MsgBox "setsockopt failed"
    If s > 0 Then Dummy = closesocket(s)
    Exit Function
End If

'we could check if setsockopt did ok...
'Dim v As Long
'ret = getsockopt(s, SOL_SOCKET, &H1006, v, 4)
'Debug.Print v

ret = bind(s, sockin, Len(sockin))
If ret <> 0 Then
     If s > 0 Then Dummy = closesocket(s)
     MsgBox "bind failed"
     Exit Function
End If


Dim lngInBuffer As Long
Dim lngBytesReturned As Long
Dim lngOutBuffer As Long

lngInBuffer = 1
ret = WSAIoctl(s, SIO_RCVALL, lngInBuffer, Len(lngInBuffer), _
lngOutBuffer, Len(lngOutBuffer), lngBytesReturned, ByVal 0, ByVal 0)
If ret <> 0 Then
    If s > 0 Then Dummy = closesocket(s)
    MsgBox "WSAIoctl failed"
    Exit Function
End If
        
SelectOps = FD_READ 'Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
ret = WSAAsyncSelect(s, HWndToMsg, WINSOCKMSG, ByVal SelectOps)
If ret <> 0 Then
    If s > 0 Then Dummy = closesocket(s)
    ConnectSock = INVALID_SOCKET
    MsgBox "INVALID_SOCKET"
    Exit Function
End If

ConnectSock = s
Exit Function
errhand:
MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Function
Function GetHostByNameAlias(ByVal hostname$) As Long
    On Error Resume Next
    Dim phe&
    Dim heDestHost As HostEnt
    Dim addrList&
    Dim retIP&
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function
Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function

 

Public Function wsck_enum_interfaces(ByRef str() As String) As Long
Dim lngSocketDescriptor   As Long
Dim lngInBuffer           As Long
Dim lngBytesReturned      As Long
Dim lngWin32apiResultCode As Long
Dim mudtWSAData As WSADataType
Dim desc As String

Call StartWinsock(desc)

lngSocketDescriptor = socket(AF_INET, SOCK_STREAM, 0)

If lngSocketDescriptor Then
    If lngWin32apiResultCode Then
    wsck_enum_interfaces = Err.LastDllError
    Exit Function
End If

End If

Dim buffer As aINTERFACE_INFO

lngWin32apiResultCode = _
WSAIoctl(lngSocketDescriptor, SIO_GET_INTERFACE_LIST, _
ByVal 0, ByVal 0, _
buffer, 1024, lngBytesReturned, ByVal 0, ByVal 0)

If lngWin32apiResultCode Then
    wsck_enum_interfaces = Err.LastDllError
    Exit Function
End If


Dim NumInterfaces As Integer
NumInterfaces = CInt(lngBytesReturned / 76)
Dim i As Integer
For i = 0 To NumInterfaces - 1
    ReDim Preserve str(i)
    
    str(i) = getascip(buffer.interfaceinfo(i).iiAddress.AddressIn.sin_addr) & ";" & getascip(buffer.interfaceinfo(i).iiNetmask.AddressIn.sin_addr)
Next i


lngWin32apiResultCode = closesocket(lngSocketDescriptor)
End Function


Public Function IsWindowsNT5() As Boolean
IsWindowsNT5 = False
Dim res As Long
'
    Dim typOSInfo As OSVERSIONINFO
     
    typOSInfo.dwOSVersionInfoSize = Len(typOSInfo)
    res = GetVersionEx(typOSInfo)
    If typOSInfo.dwMajorVersion >= 5 Then IsWindowsNT5 = True
    
    End Function
    
Public Sub AddLog(ByVal strTexte As String)
    Dim intFreefile As Integer
    intFreefile = FreeFile
    Open App.Path & "\sniffer.log" For Append As #intFreefile
        Print #intFreefile, strTexte
    Close #intFreefile
End Sub

