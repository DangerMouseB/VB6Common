Attribute VB_Name = "mDLL_Main"
'*************************************************************************************************************************************************************************************************************************************************
'
' This module is from the vbAdvance product (c) 2001 - 2007 Young Dynamic Software. All rights reserved.
'
' vbAdvance is now unsupported freeware - see http://vb.mvps.org/tools/vbAdvance/
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit

Public Const INITClassName As String = "cDLL_DummyInit"

Function DllMain(ByVal hinstDLL As Long, ByVal fdwReason As Long, ByVal lpvReserved As Long) As Long
    Const DLL_PROCESS_ATTACH As Long = 1
    If fdwReason = DLL_PROCESS_ATTACH Then
        initVBRuntime hinstDLL
        DllMain = 1
        apiMessageBoxA 0, "MAIN - DLL_PROCESS_ATTACH", "My Box", vbOKOnly
    End If
    If fdwReason = DLL_PROCESS_DETACH Then
        apiMessageBoxA 0, "MAIN - DLL_PROCESS_DETACH", "My Box", vbOKOnly
    End If
End Function


