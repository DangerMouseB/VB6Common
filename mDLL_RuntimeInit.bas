Attribute VB_Name = "mDLL_RuntimeInit"
'*************************************************************************************************************************************************************************************************************************************************
'
' This module is from the vbAdvance product (c) 2001 - 2007 Young Dynamic Software. All rights reserved.
'
' vbAdvance is now unsupported freeware - see http://vb.mvps.org/tools/vbAdvance/
'
'*************************************************************************************************************************************************************************************************************************************************

' Warnings  :   DO NOT MODIFY THIS CODE UNLESS YOU KNOW EXACTLY WHAT YOU ARE DOING.
' Requires vbAdvance.tlb (vbAdvance Type Library) Needed only at compile-time.
' Requires cDummyInit.cls

' don't use API calls defined in the nWin32 modules during initialisation as it can crash on some machines!!!! wierd!!! use the type lib calls instead??!!!
' CopyMemory instead of apiCopyMemory
' OutputDebugString instead of apiOutputDebugStringA

Option Explicit


'==============================================================================
'FUNCTION DELEGATOR CODE from Feb.2000 VBPJ article "Call Function Pointers" by Matthew Curland - http://www.powervb.com

Private Const cDelegateASM As Currency = -368956918007638.6215@

Private Type DelegatorVTables
    VTable(7) As Long
End Type

Private Type FunctionDelegator
    pVTable As Long
    pfn As Long
End Type

Private myDelegateASM As Currency
Private myVTables As DelegatorVTables
Private mypVTableOKQI As Long
Private mypVTableFailQI As Long

'END FUNCTION DELEGATOR CODE
'==============================================================================

Private myInitObject As cDLL_DummyInit                                    ' Object reference which keeps runtime alive:

Global g_IsInitialised As Boolean

Sub initDLL()
    initVBRuntime GetModuleHandle(DLL_NAME)
    g_IsInitialised = True
End Sub

Sub initVBRuntime(ByVal hMod As Long)
    Dim sFile As String, lLen As Long, lRet As Long, i As Long, lpTypeLib As Long, TLI As ITypeLib, lppTypeInfo As Long, TI As ITypeInfo, sName As String, pAttr As Long, TA As TYPEATTR, IID_ClassFactory As VBGUID
    Dim IID_IUnknown As VBGUID, pGetClass As Long, pCall As ICallDLLGetClassObject, FD As FunctionDelegator, pICF As IClassFactory, pUnk As IUnknown

    'Make sure parent process is not VB IDE:
    If GetModuleHandle("VBA6.DLL") <> 0 Then Exit Sub
    If GetModuleHandle("VBA5.DLL") <> 0 Then Exit Sub
    sFile = Space$(260)
    lLen = Len(sFile)
    lRet = GetModuleFileName(hMod, sFile, lLen)
    If lRet Then
        sFile = Left$(sFile, lLen - 1)
        lpTypeLib = LoadTypeLibEx(sFile, REGKIND_NONE)
        CopyMemory TLI, lpTypeLib, 4
        For i = 0 To TLI.GetTypeInfoCount - 1
            If TLI.GetTypeInfoType(i) = TKIND_COCLASS Then
                lppTypeInfo = TLI.GetTypeInfo(i)
                CopyMemory TI, lppTypeInfo, 4
                TI.GetDocumentation DISPID_UNKNOWN, sName, "", 0, ""
                If lstrcmp(sName, INITClassName) = 0 Then
                    pAttr = TI.GetTypeAttr
                    CopyMemory TA, ByVal pAttr, Len(TA)
                    TI.ReleaseTypeAttr pAttr
                    If TA.wTypeFlags Then Exit For
                End If
            End If
        Next i
        IID_ClassFactory.Data1 = 1
        IID_ClassFactory.Data4(0) = &HC0
        IID_ClassFactory.Data4(7) = &H46
        IID_IUnknown.Data4(0) = &HC0
        IID_IUnknown.Data4(7) = &H46
        pGetClass = GetProcAddress(hMod, "DllGetClassObject")
        If pGetClass Then
            CopyMemory pCall, initializeDelegator(FD, pGetClass), 4
            lRet = pCall.call(TA.IID, IID_ClassFactory, pICF)
            If lRet <> CLASS_E_CLASSNOTAVAILABLE Then
                lRet = pICF.CreateInstance(0&, IID_IUnknown, pUnk)
                If lRet = S_OK Then
                    Set myInitObject = pUnk
                    myInitObject.initVBCall
                    CopyMemory pCall, 0&, 4
                    Set pICF = Nothing
                    Set pUnk = Nothing
                End If
            End If
        End If
    End If

End Sub


'==============================================================================
'FUNCTION DELEGATOR CODE from Feb.2000 VBPJ article "Call Function Pointers" by Matthew Curland - http://www.powervb.com

Private Function initializeDelegator(delegator As FunctionDelegator, Optional ByVal pfn As Long) As IUnknown
    If mypVTableOKQI = 0 Then initializeVTables
    delegator.pVTable = mypVTableOKQI
    delegator.pfn = pfn
    CopyMemory initializeDelegator, VarPtr(delegator), 4
End Function

Private Sub initializeVTables()
    Dim pAddRefRelease As Long
    myVTables.VTable(0) = FuncAddr(AddressOf queryInterfaceOK)
    myVTables.VTable(4) = FuncAddr(AddressOf queryInterfaceFail)
    pAddRefRelease = FuncAddr(AddressOf AddRefRelease)
    myVTables.VTable(1) = pAddRefRelease
    myVTables.VTable(5) = pAddRefRelease
    myVTables.VTable(2) = pAddRefRelease
    myVTables.VTable(6) = pAddRefRelease
    myDelegateASM = cDelegateASM
    myVTables.VTable(3) = VarPtr(myDelegateASM)
    myVTables.VTable(7) = myVTables.VTable(3)
    mypVTableOKQI = VarPtr(myVTables.VTable(0))
    mypVTableFailQI = VarPtr(myVTables.VTable(4))
End Sub

Private Function queryInterfaceOK(This As FunctionDelegator, riid As Long, pvObj As Long) As Long
    pvObj = VarPtr(This)
    This.pVTable = mypVTableFailQI
End Function

Private Function AddRefRelease(ByVal This As Long) As Long
End Function

Private Function queryInterfaceFail(ByVal This As Long, riid As Long, pvObj As Long) As Long
    pvObj = 0
    queryInterfaceFail = E_NOINTERFACE
End Function

Private Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function

'END FUNCTION DELEGATOR CODE
'==============================================================================
