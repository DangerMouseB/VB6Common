VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTP_AdvSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIF_ALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Private Const WM_USER = &H400

Private Const NIN_BALLOONSHOW = WM_USER + 2
Private Const NIN_BALLOONHIDE = WM_USER + 3
Private Const NIN_BALLOONTIMEOUT = WM_USER + 4
Private Const NIN_BALLOONUSERCLICK = WM_USER + 5

Private Const NOTIFYICONDATA_V1_SIZE = 88

Public Enum NIIF
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_NOSOUND = &H10
End Enum

Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private myVersion As Long
Private myTipSize As Integer
Private myIsCreated As Boolean
Private WithEvents myOwnerForm As Form
Attribute myOwnerForm.VB_VarHelpID = -1
Private myNID As NOTIFYICONDATA

Private Declare Function DllGetVersion Lib "shell32" (dwVersion As DLLVERSIONINFO) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Event LButtonDown()
Public Event LButtonUp()
Public Event LButtonDblClk()
Public Event RButtonDown()
Public Event RButtonUp()
Public Event RButtonDblClk()

Public Event BalloonShow()
Public Event BalloonHide()
Public Event BalloonTimeout()
Public Event BalloonUserClick()

Private Sub class_Initialize()
    Dim dvi As DLLVERSIONINFO
    myIsCreated = False
    dvi.cbSize = Len(dvi)
    On Error Resume Next
    DllGetVersion dvi
    On Error GoTo 0
    myVersion = dvi.dwMajorVersion
    myTipSize = IIf(myVersion < 5, 63, 127)
End Sub

Private Sub class_Terminate()
    destroy
End Sub



Property Get shellVersion() As Long
    shellVersion = myVersion
End Property

Property Let icon(ByVal icon As StdPicture)
    If Not myIsCreated Then Exit Property
    myNID.hIcon = icon
    Shell_NotifyIcon NIM_MODIFY, myNID
End Property

Property Get icon() As StdPicture
    If Not myIsCreated Then Exit Property
    icon = myNID.hIcon
End Property

Property Let tooltip(ByVal tooltip As String)
    If Not myIsCreated Then Exit Property
    myNID.szTip = Left$(tooltip, myTipSize) & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, myNID
End Property

Property Get tooltip() As String
    If Not myIsCreated Then Exit Property
    tooltip = Left$(myNID.szTip, InStr(myNID.szTip, vbNullChar) - 1)
End Property

Sub create(mouseEventsForm As Form, trayIcon As StdPicture, caption As String)
    If myIsCreated Then Exit Sub
    If myVersion >= 5 Then
        myNID.cbSize = Len(myNID)
        myNID.uFlags = NIF_ALL Or NIF_INFO
    Else
        myNID.cbSize = NOTIFYICONDATA_V1_SIZE
        myNID.uFlags = NIF_ALL
    End If
    myNID.hWnd = mouseEventsForm.hWnd
    myNID.uID = 1
    myNID.uCallbackMessage = WM_MOUSEMOVE
    myNID.hIcon = trayIcon
    myNID.szTip = caption & vbNullChar
    myIsCreated = CBool(Shell_NotifyIcon(NIM_ADD, myNID))
    If myIsCreated Then Set myOwnerForm = mouseEventsForm
End Sub

Sub destroy()
    If Not myIsCreated Then Exit Sub
    myIsCreated = Not CBool(Shell_NotifyIcon(NIM_DELETE, myNID))
    Set myOwnerForm = Nothing
End Sub

Sub showBalloon(ByVal Text As String, Optional ByVal Title As String = "", Optional ByVal Flags As NIIF = NIIF_NONE Or NIIF_NOSOUND)
    If Not (myIsCreated And myVersion >= 5) Then Exit Sub
    myNID.szInfo = Left$(Text, 255) & vbNullChar
    myNID.szInfoTitle = Left$(Title, 63) & vbNullChar
    myNID.dwInfoFlags = Flags
    Shell_NotifyIcon NIM_MODIFY, myNID
End Sub

Sub hideBalloon()
    If Not (myIsCreated And myVersion >= 5) Then Exit Sub
    myNID.szInfo = vbNullChar
    Shell_NotifyIcon NIM_MODIFY, myNID
End Sub

Private Sub myOwnerForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case x / Screen.TwipsPerPixelX
        Case WM_LBUTTONDOWN:
            RaiseEvent LButtonDown
        Case WM_LBUTTONUP:
            RaiseEvent LButtonUp
        Case WM_LBUTTONDBLCLK:
            RaiseEvent LButtonDblClk
        Case WM_RBUTTONDOWN:
            RaiseEvent RButtonDown
        Case WM_RBUTTONUP:
            RaiseEvent RButtonUp
        Case WM_RBUTTONDBLCLK:
            RaiseEvent RButtonDblClk
        Case NIN_BALLOONSHOW:
            RaiseEvent BalloonShow
        Case NIN_BALLOONHIDE:
            RaiseEvent BalloonHide
        Case NIN_BALLOONTIMEOUT:
            RaiseEvent BalloonTimeout
        Case NIN_BALLOONUSERCLICK:
            RaiseEvent BalloonUserClick
    End Select
End Sub
