VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "IP Promiscuous Sniffer By Erwan L. V2"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txt_display 
      Height          =   2895
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   6000
      Width           =   8535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Interface (IP)"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author : Erwan L.
'email:erwan.l@free.fr
Public WithEvents listview1  As VBControlExtender
Attribute listview1.VB_VarHelpID = -1

Public Sub ListView1_ObjectEvent(info As EventInfo)
'Debug.Print info.Name
Select Case info.Name
Case "Click":
Dim bbytes() As Byte
bbytes = listview1.object.selecteditem.Tag
If UBound(bbytes) <= 0 Then Exit Sub
    display_packet bbytes
    ' listview1.object.selecteditem ip
    ' listview1.object.selecteditem.subitems(1) mac
    ' listview1.object.selecteditem.subitems(2) index
Case Else:
End Select
End Sub

Private Sub Command1_Click()
On Error GoTo errhand
cnt = 0
If IsWindowsNT5 = False Then
    MsgBox "nt5 or above only!!!"
    Exit Sub
End If
    Dim sSave As String
    Me.AutoRedraw = True
    'Set Obj = Me.Text1
    'Start subclassing
    HookForm Me
    'create a new winsock session
    StartWinsock sSave
    'show the winsock version on this form
    If InStr(1, sSave, Chr$(0)) > 0 Then sSave = Left$(sSave, InStr(1, sSave, Chr$(0)) - 1)
    'Me.Print sSave
    'connect
    lSocket = ConnectSock(Combo1.Text, 7000, Me.hwnd, False)

Command1.Enabled = False
Command2.Enabled = True
Exit Sub
errhand:
    MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError, vbCritical, "wsck_displayadapterinfo"
    AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Sub

Private Sub Command2_Click()
On Error GoTo errhand
Call WSAAsyncSelect(lSocket, Me.hwnd, ByVal 1025, 0)
'close our connection
    closesocket lSocket
    'end winsock session
    EndWinsock
    'stop subclassing
    UnHookForm Me
    
    Command1.Enabled = True
Command2.Enabled = False
Exit Sub
errhand:
    MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError, vbCritical, "wsck_displayadapterinfo"
    AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Sub


Sub make_ctl()
'
On Error Resume Next
Licenses.Add "MSComctlLib.listviewctrl"
Err.Clear
Set Me.listview1 = Me.Controls.Add("MSComctlLib.listviewctrl", "listview1", Me)
If Err <> 0 Then
    MsgBox Err.Description & vbCrLf & "cant use MSCOMCTL.OCX : MSComctlLib.listviewctrl"
    Unload Me
    End If
On Error GoTo 0
On Error GoTo errhand
Me.listview1.Top = Me.ScaleTop + 10
Me.listview1.Width = Me.ScaleWidth - 50
Me.listview1.Height = 250
Me.listview1.Left = 10
Me.listview1.Visible = True
'
listview1.object.ColumnHeaders.Add , , "IP SRC"
listview1.object.ColumnHeaders.Add 2, , "Port SRC"
listview1.object.ColumnHeaders.Add 3, , "IP Dest"
listview1.object.ColumnHeaders.Add 4, , "Port SRC"
listview1.object.ColumnHeaders.Add 5, , "Prot."
listview1.object.ColumnHeaders.Add 6, , "Length"
listview1.object.View = 3 'lvwreport
listview1.object.fullrowselect = True


Exit Sub
errhand:
 MousePointer = vbDefault
   MsgBox Err.Description, vbCritical, Me.Caption
   AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Sub
Private Sub Form_Load()
On Error GoTo errhand
make_ctl
Command1.Enabled = True
Command2.Enabled = False
displayadapterinfo
Txt_display.FontSize = 8
Txt_display.FontName = "Courier"
Exit Sub
errhand:
    MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError, vbCritical, "wsck_displayadapterinfo"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errhand
Call WSAAsyncSelect(lSocket, Me.hwnd, ByVal 1025, 0)
    'close our connection
    closesocket lSocket
    'end winsock session
    EndWinsock
    'stop subclassing
    UnHookForm Me
Exit Sub
errhand:
    MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError, vbCritical, "wsck_displayadapterinfo"
    AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Sub

Sub displayadapterinfo()
On Error GoTo errhand
Combo1.Clear
Dim str() As String
Call wsck_enum_interfaces(str)
Dim i As Integer
Dim j As Integer
Dim v As Variant
For i = 0 To UBound(str)
    v = Split(str(i), ";")
    Combo1.AddItem v(0)
Next i
Combo1.ListIndex = 0
Exit Sub
errhand:
    MsgBox Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError, vbCritical, "wsck_displayadapterinfo"
    AddLog Err.Description & vbCrLf & Err.Number & vbCrLf & Err.LastDllError
End Sub

Private Sub Text1_DblClick()
Text1 = ""
End Sub

Sub display_packet(bbytes() As Byte)
Txt_display = ""
Dim s As String
Dim sh As String
Dim st As String
Dim i As Integer
Dim c As Integer
For i = 0 To listview1.object.selecteditem.subitems(5) - 1 ' UBound(bbytes)
c = c + 1

s = Hex(bbytes(i))
If Len(s) = 1 Then s = "0" & s
sh = sh & " " & s

s = Chr(bbytes(i))
If Asc(s) < 32 Then s = "."
st = st & s



If c = 16 Then
    c = 0
    Txt_display = Txt_display & " " & sh & " " & st & vbCrLf
    sh = ""
    st = ""
End If

Next i

If sh <> "" Then
    Txt_display = Txt_display & " " & sh & Space(48 - Len(sh)) & " " & st & vbCrLf
    sh = ""
    st = ""
End If

End Sub

'Author : Erwan L.
'email:erwan.l@free.fr
