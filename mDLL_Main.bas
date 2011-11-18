Attribute VB_Name = "mDLL_Main"
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


