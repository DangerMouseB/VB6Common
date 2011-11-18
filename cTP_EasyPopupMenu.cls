VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTP_EasyPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' EasyPopupMenu
' Shahin Gasanov
' www.gasanov.net

Option Explicit

Public Enum MF
    MF_STRING = &H0
    MF_GRAYED = &H1
    MF_CHECKED = &H8
    MF_SEPARATOR = &H800
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const MF_POPUP = &H10
Private Const MF_BYPOSITION = &H400

Private Const TPM_RETURNCMD = &H100

Private Declare Function apiCreatePopupMenu Lib "user32" Alias "CreatePopupMenu" () As Long
Private Declare Function apiAppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function apiGetSubMenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function apiCheckMenuRadioItem Lib "user32" Alias "CheckMenuRadioItem" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare Function apiTrackPopupMenu Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Any) As Long
Private Declare Function apiDestroyMenu Lib "user32" Alias "DestroyMenu" (ByVal hMenu As Long) As Long
Private Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long


Private myhMenu As Long
Private myhSubMenu As Long


Private Sub class_Initialize()
    createMenu
End Sub

Private Sub class_Terminate()
    deleteMenu
End Sub

Sub createMenu()
    If myhMenu = 0 Then myhMenu = apiCreatePopupMenu
End Sub

Sub addMenuItem(itemCaption As String, itemType As MF, itemNumber As Long)
    apiAppendMenu myhMenu, itemType, itemNumber, itemCaption
End Sub

Sub createSubmenu(itemCaption As String)
    myhSubMenu = apiCreatePopupMenu()
    apiAppendMenu myhMenu, MF_POPUP, myhSubMenu, itemCaption
End Sub

Sub addSubmenuItem(itemCaption As String, itemType As MF, itemNumber As Long)
    apiAppendMenu myhSubMenu, itemType, itemNumber, itemCaption
End Sub

Sub checkRadioItem(FirstItem As Long, LastItem As Long, RadioItem As Long)
    apiCheckMenuRadioItem myhSubMenu, FirstItem, LastItem, RadioItem, MF_BYPOSITION
End Sub

Function trackMenu(hWnd As Long) As Long
    Dim pa As POINTAPI
    apiGetCursorPos pa
    trackMenu = apiTrackPopupMenu(myhMenu, TPM_RETURNCMD, pa.x, pa.y, 0, hWnd, 0&)
End Function

Sub deleteMenu()
    If myhMenu > 0 Then If apiDestroyMenu(myhMenu) Then myhMenu = 0
End Sub
