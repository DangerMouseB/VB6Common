Attribute VB_Name = "mDB_SafeArray"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit

Function getSafeArrayDetailsFromByteArray(anArray() As Byte, oSA As SAFEARRAY) As HRESULT
    Dim ptr As Long
    If Not IsArray(anArray) Then getSafeArrayDetailsFromByteArray = CHRESULT(E_INVALIDARG): Exit Function
    ptr = apiVarPtrArray(anArray)
    apiCopyMemory ptr, ByVal ptr, 4                      ' don't check for ptr = 0 because that would mean VB has broken
    apiCopyMemory oSA.cDims, ByVal ptr, 16        ' The fixed part of the SAFEARRAY structure is 16 bytes.
    If oSA.cDims > 0 Then
        ReDim oSA.rgSABound(1 To oSA.cDims)
        apiCopyMemory oSA.rgSABound(1), ByVal ptr + 16&, oSA.cDims * Len(oSA.rgSABound(1))
    End If
End Function

Function getSafeArrayDetails(anArray As Variant, oSA As SAFEARRAY) As HRESULT
    Dim ptr As Long
    ptr = getSafeArrayPointer(anArray)
    If ptr = 0 Then getSafeArrayDetails = CHRESULT(E_INVALIDARG): Exit Function
    apiCopyMemory oSA.cDims, ByVal ptr, 16         ' The fixed part of the SAFEARRAY structure is 16 bytes.
    If oSA.cDims > 0 Then
        ReDim oSA.rgSABound(1 To oSA.cDims)
        apiCopyMemory oSA.rgSABound(1), ByVal ptr + 16&, oSA.cDims * Len(oSA.rgSABound(1))
    End If
End Function

Function getSafeArrayPointer(anArray As Variant) As Long
    Dim ptr As Long, vType As Integer
    If Not IsArray(anArray) Then Exit Function
    apiCopyMemory vType, anArray, 2                                                         ' Get the VARTYPE value from the first 2 bytes of the VARIANT structure
    apiCopyMemory ptr, ByVal VarPtr(anArray) + 8, 4                                    ' Get the pointer to the array descriptor (SAFEARRAY structure)   NOTE: A Variant's descriptor, padding & union take up 8 bytes.
    If (vType And VT_BYREF) <> 0 Then apiCopyMemory ptr, ByVal ptr, 4        ' Test if lp is a pointer or a pointer to a pointer and if so get real pointer to the array descriptor (SAFEARRAY structure)
    getSafeArrayPointer = ptr
End Function

Function redimPreserve(a As Variant, nDimensions As Long, x1 As Long, x2 As Long, y1 As Long, y2 As Long, z1 As Long, z2 As Long) As HRESULT
    Dim ASA As SAFEARRAY, retVal As HRESULT, currentSize As Long, redimmedSize As Long, i As Long, b As Variant, pASA As Long, pBSA As Long, temp As Long, varType As Long
    
    ' check the parameters
    If getSafeArrayDetails(a, ASA).HRESULT <> S_OK Then redimPreserve = CHRESULT(E_INVALIDARG): Exit Function
    Select Case True
        Case nDimensions < 1, nDimensions > 3
            redimPreserve = CHRESULT(E_WRONG_NUMBER_OF_DIMENSIONS): Exit Function
        Case ASA.cDims < 1, ASA.cDims > 3
            redimPreserve = CHRESULT(E_INVALIDARG): Exit Function
        Case x2 < x1, y2 < y1, z2 < z1
            redimPreserve = CHRESULT(E_INVALIDARG): Exit Function
    End Select
    currentSize = ASA.rgSABound(1).cElements
    For i = 2 To ASA.cDims
        currentSize = currentSize * ASA.rgSABound(i).cElements
    Next
    If nDimensions = 1 Then redimmedSize = (x2 - x1 + 1)
    If nDimensions = 2 Then redimmedSize = (x2 - x1 + 1) * (y2 - y1 + 1)
    If nDimensions = 3 Then redimmedSize = (x2 - x1 + 1) * (y2 - y1 + 1) * (z2 - z1 + 1)
    If currentSize <> redimmedSize Then redimPreserve = CHRESULT(E_INVALIDARG): Exit Function
    
    pASA = getSafeArrayPointer(a)
    If nDimensions <> ASA.cDims Then
        ' handle the change in numberOfDimensions
    
        ' create a blank array of the right number of dimensions and element size
        redimPreserve = apiSafeArrayGetVartype(pASA, varType)
        If redimPreserve.HRESULT <> S_OK Then Exit Function     ' how I wish I had C syntax sometimes
        If _
                varType <> vbByte And _
                varType <> vbInteger And _
                varType <> vbLong And _
                varType <> vbSingle And _
                varType <> vbDouble And _
                varType <> vbBoolean And _
                varType <> vbDate And _
                varType <> vbCurrency And _
                varType <> vbVariant Then
            redimPreserve = CHRESULT(E_UNSUPPORT_TYPE_FOR_CHANGE_OF_DIMENSIONS)
            Exit Function
        End If
'        apiMessageBoxA 0, "vartype: " & varType, "A", vbCritical
        redimPreserve = createOneElementArray(b, nDimensions, varType)
        If redimPreserve.HRESULT <> S_OK Then Exit Function
        
        pBSA = getSafeArrayPointer(b)
        
        ' make A point to B'a data and vice versa
        apiCopyMemory temp, ByVal pBSA + 12, 4
        apiCopyMemory ByVal pBSA + 12, ByVal pASA + 12, 4
        apiCopyMemory ByVal pASA + 12, temp, 4
        
'        apiMessageBoxA 0, "vartype: " & varType, "B", vbCritical

        ' now can safely Redim A
        redimPreserve = createOneElementArray(a, nDimensions, varType)
        If redimPreserve.HRESULT <> S_OK Then Exit Function     ' how I wish I had C syntax sometimes
        
'        apiMessageBoxA 0, "vartype: " & varType, "C", vbCritical
        
        pASA = getSafeArrayPointer(a)
        
        ' switch back again
        apiCopyMemory temp, ByVal pBSA + 12, 4
        apiCopyMemory ByVal pBSA + 12, ByVal pASA + 12, 4
        apiCopyMemory ByVal pASA + 12, temp, 4
        
'        apiMessageBoxA 0, "vartype: " & varType, "D", vbCritical
        
        Erase b ' could be left to do on the stack?
        
'        apiMessageBoxA 0, "vartype: " & varType, "E", vbCritical

    End If

    ' do the redim - note that there was now change in dimension we can just write to the SAFEARRAYBOUND directly
    i = 16 - 4
    Select Case nDimensions
        Case 1
            i = i + 4: apiCopyMemory ByVal pASA + i, x2 - x1 + 1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, x1, 4
        Case 2
            i = i + 4: apiCopyMemory ByVal pASA + i, y2 - y1 + 1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, y1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, x2 - x1 + 1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, x1, 4
        Case 3
            i = i + 4: apiCopyMemory ByVal pASA + i, z2 - z1 + 1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, z1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, y2 - y1 + 1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, y1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, x2 - x1 + 1, 4
            i = i + 4: apiCopyMemory ByVal pASA + i, x1, 4
    End Select
End Function

Private Function createOneElementArray(oArray As Variant, nDimensions As Long, varType As Long) As HRESULT
    createOneElementArray = CHRESULT(S_OK)
    Select Case varType
        Case vbByte
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Byte
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Byte
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Byte
        Case vbInteger
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Integer
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Integer
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Integer
        Case vbLong
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Long
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Long
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Long
        Case vbSingle
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Single
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Single
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Single
        Case vbDouble
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Double
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Double
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Double
        Case vbBoolean
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Boolean
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Boolean
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Boolean
        Case vbDate
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Date
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Date
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Date
        Case vbCurrency
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Currency
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Currency
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Currency
        Case vbString
            If nDimensions = 1 Then ReDim oArray(1 To 1) As String
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As String
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As String
        Case vbVariant
            If nDimensions = 1 Then ReDim oArray(1 To 1) As Variant
            If nDimensions = 2 Then ReDim oArray(1 To 1, 1 To 1) As Variant
            If nDimensions = 3 Then ReDim oArray(1 To 1, 1 To 1, 1 To 1) As Variant
        Case Else
            createOneElementArray = CHRESULT(E_NOTIMPL)
    End Select
End Function

Function redimTranspose2D(anArray As Variant) As HRESULT
End Function

Function createDoubleArrayMap(oMap() As Double, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    Dim ptrMapSA As Long, retVal As HRESULT
    
    ' get VB to create the oMap's SAFEARRY with the appropiate number of dimensions
    Select Case nDimensions
        Case 1
            ReDim oMap(1 To 1)
        Case 2
            ReDim oMap(1 To 1, 1 To 1)
        Case 3
            ReDim oMap(1 To 1, 1 To 1, 1 To 1)
        Case Else
            createDoubleArrayMap = CHRESULT(E_WRONG_NUMBER_OF_DIMENSIONS): Exit Function     ' invalid number of dimensions
    End Select
    
    ' get the pointer to oMap's SAFEARRAY
    ptrMapSA = apiVarPtrArray(oMap)
    If ptrMapSA = 0 Then createDoubleArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    apiCopyMemory ptrMapSA, ByVal ptrMapSA, 4
    If ptrMapSA = 0 Then createDoubleArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    
    ' release the memory of the contents of oMap
    If apiSafeArrayDestroyData(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createDoubleArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' copy the dimension variables
    apiCopyMemory ByVal ptrMapSA + 16 + 0, i2 - i1 + 1, 4
    apiCopyMemory ByVal ptrMapSA + 16 + 4, i1, 4
    If nDimensions >= 2 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 8, j2 - j1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 12, j1, 4
    End If
    If nDimensions >= 3 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 16, k2 - k1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 20, k1, 4
    End If
    
    ' lock the array
    If apiSafeArrayLock(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createDoubleArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' set oMap to point to the same data in ptr
    apiCopyMemory ByVal ptrMapSA + 12, ptr, 4
   
End Function

Function createSingleArrayMap(oMap() As Single, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    Dim ptrMapSA As Long, retVal As HRESULT
    
    ' get VB to create the oMap's SAFEARRY with the appropiate number of dimensions
    Select Case nDimensions
        Case 1
            ReDim oMap(1 To 1)
        Case 2
            ReDim oMap(1 To 1, 1 To 1)
        Case 3
            ReDim oMap(1 To 1, 1 To 1, 1 To 1)
        Case Else
            createSingleArrayMap = CHRESULT(E_WRONG_NUMBER_OF_DIMENSIONS): Exit Function     ' invalid number of dimensions
    End Select
    
    ' get the pointer to oMap's SAFEARRAY
    ptrMapSA = apiVarPtrArray(oMap)
    If ptrMapSA = 0 Then createSingleArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    apiCopyMemory ptrMapSA, ByVal ptrMapSA, 4
    If ptrMapSA = 0 Then createSingleArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    
    ' release the memory of the contents of oMap
    If apiSafeArrayDestroyData(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createSingleArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' copy the dimension variables
    apiCopyMemory ByVal ptrMapSA + 16 + 0, i2 - i1 + 1, 4
    apiCopyMemory ByVal ptrMapSA + 16 + 4, i1, 4
    If nDimensions >= 2 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 8, j2 - j1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 12, j1, 4
    End If
    If nDimensions >= 3 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 16, k2 - k1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 20, k1, 4
    End If
    
    ' lock the array
    If apiSafeArrayLock(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createSingleArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' set oMap to point to the same data in ptr
    apiCopyMemory ByVal ptrMapSA + 12, ptr, 4
   
End Function

Function createDateArrayMap(oMap() As Date, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    Dim ptrMapSA As Long, retVal As HRESULT
    
    ' get VB to create the oMap's SAFEARRY with the appropiate number of dimensions
    Select Case nDimensions
        Case 1
            ReDim oMap(1 To 1)
        Case 2
            ReDim oMap(1 To 1, 1 To 1)
        Case 3
            ReDim oMap(1 To 1, 1 To 1, 1 To 1)
        Case Else
            createDateArrayMap = CHRESULT(E_WRONG_NUMBER_OF_DIMENSIONS): Exit Function     ' invalid number of dimensions
    End Select
    
    ' get the pointer to oMap's SAFEARRAY
    ptrMapSA = apiVarPtrArray(oMap)
    If ptrMapSA = 0 Then createDateArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    apiCopyMemory ptrMapSA, ByVal ptrMapSA, 4
    If ptrMapSA = 0 Then createDateArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    
    ' release the memory of the contents of oMap
    If apiSafeArrayDestroyData(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createDateArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' copy the dimension variables
    apiCopyMemory ByVal ptrMapSA + 16 + 0, i2 - i1 + 1, 4
    apiCopyMemory ByVal ptrMapSA + 16 + 4, i1, 4
    If nDimensions >= 2 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 8, j2 - j1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 12, j1, 4
    End If
    If nDimensions >= 3 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 16, k2 - k1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 20, k1, 4
    End If
    
    ' lock the array
    If apiSafeArrayLock(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createDateArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' set oMap to point to the same data in ptr
    apiCopyMemory ByVal ptrMapSA + 12, ptr, 4
   
End Function

Function createLongArrayMap(oMap() As Long, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    Dim ptrMapSA As Long, retVal As HRESULT
    
    ' get VB to create the oMap's SAFEARRY with the appropiate number of dimensions
    Select Case nDimensions
        Case 1
            ReDim oMap(1 To 1)
        Case 2
            ReDim oMap(1 To 1, 1 To 1)
        Case 3
            ReDim oMap(1 To 1, 1 To 1, 1 To 1)
        Case Else
            createLongArrayMap = CHRESULT(E_WRONG_NUMBER_OF_DIMENSIONS): Exit Function     ' invalid number of dimensions
    End Select
    
    ' get the pointer to oMap's SAFEARRAY
    ptrMapSA = apiVarPtrArray(oMap)
    If ptrMapSA = 0 Then createLongArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    apiCopyMemory ptrMapSA, ByVal ptrMapSA, 4
    If ptrMapSA = 0 Then createLongArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    
    ' release the memory of the contents of oMap
    If apiSafeArrayDestroyData(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createLongArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' copy the dimension variables
    apiCopyMemory ByVal ptrMapSA + 16 + 0, i2 - i1 + 1, 4
    apiCopyMemory ByVal ptrMapSA + 16 + 4, i1, 4
    If nDimensions >= 2 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 8, j2 - j1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 12, j1, 4
    End If
    If nDimensions >= 3 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 16, k2 - k1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 20, k1, 4
    End If
    
    ' lock the array
    If apiSafeArrayLock(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createLongArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' set oMap to point to the same data in ptr
    apiCopyMemory ByVal ptrMapSA + 12, ptr, 4
   
End Function

Function createIntegerArrayMap(oMap() As Integer, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    Dim ptrMapSA As Long, retVal As HRESULT
    
    ' get VB to create the oMap's SAFEARRY with the appropiate number of dimensions
    Select Case nDimensions
        Case 1
            ReDim oMap(1 To 1)
        Case 2
            ReDim oMap(1 To 1, 1 To 1)
        Case 3
            ReDim oMap(1 To 1, 1 To 1, 1 To 1)
        Case Else
            createIntegerArrayMap = CHRESULT(E_WRONG_NUMBER_OF_DIMENSIONS): Exit Function     ' invalid number of dimensions
    End Select
    
    ' get the pointer to oMap's SAFEARRAY
    ptrMapSA = apiVarPtrArray(oMap)
    If ptrMapSA = 0 Then createIntegerArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    apiCopyMemory ptrMapSA, ByVal ptrMapSA, 4
    If ptrMapSA = 0 Then createIntegerArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    
    ' release the memory of the contents of oMap
    If apiSafeArrayDestroyData(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createIntegerArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' copy the dimension variables
    apiCopyMemory ByVal ptrMapSA + 16 + 0, i2 - i1 + 1, 4
    apiCopyMemory ByVal ptrMapSA + 16 + 4, i1, 4
    If nDimensions >= 2 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 8, j2 - j1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 12, j1, 4
    End If
    If nDimensions >= 3 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 16, k2 - k1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 20, k1, 4
    End If
    
    ' lock the array
    If apiSafeArrayLock(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createIntegerArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' set oMap to point to the same data in ptr
    apiCopyMemory ByVal ptrMapSA + 12, ptr, 4
   
End Function

Function createByteArrayMap(oMap() As Byte, ptr As Long, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long) As HRESULT
    Dim ptrMapSA As Long, retVal As HRESULT
    
    ' get VB to create the oMap's SAFEARRY with the appropiate number of dimensions
    Select Case nDimensions
        Case 1
            ReDim oMap(1 To 1)
        Case 2
            ReDim oMap(1 To 1, 1 To 1)
        Case 3
            ReDim oMap(1 To 1, 1 To 1, 1 To 1)
        Case Else
            createByteArrayMap = CHRESULT(E_WRONG_NUMBER_OF_DIMENSIONS): Exit Function     ' invalid number of dimensions
    End Select
    
    ' get the pointer to oMap's SAFEARRAY
    ptrMapSA = apiVarPtrArray(oMap)
    If ptrMapSA = 0 Then createByteArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    apiCopyMemory ptrMapSA, ByVal ptrMapSA, 4
    If ptrMapSA = 0 Then createByteArrayMap = CHRESULT(E_UNEXPECTED): Exit Function
    
    ' release the memory of the contents of oMap
    If apiSafeArrayDestroyData(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createByteArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' copy the dimension variables
    apiCopyMemory ByVal ptrMapSA + 16 + 0, i2 - i1 + 1, 4
    apiCopyMemory ByVal ptrMapSA + 16 + 4, i1, 4
    If nDimensions >= 2 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 8, j2 - j1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 12, j1, 4
    End If
    If nDimensions >= 3 Then
        apiCopyMemory ByVal ptrMapSA + 16 + 16, k2 - k1 + 1, 4
        apiCopyMemory ByVal ptrMapSA + 16 + 20, k1, 4
    End If
    
    ' lock the array
    If apiSafeArrayLock(ptrMapSA).HRESULT <> S_OK Then
        Erase oMap   ' we had a problem so let VB clean it up
        createByteArrayMap = CHRESULT(E_UNEXPECTED)
        Exit Function
    End If
    
    ' set oMap to point to the same data in ptr
    apiCopyMemory ByVal ptrMapSA + 12, ptr, 4
   
End Function

Function releaseArrayMap(map As Variant) As HRESULT
    Dim ptrMapSA As Long, cLocks As Long, nDimensions As Integer, retVal As HRESULT
    
    ' get the pointer to map's SAFEARRAY
    ptrMapSA = getSafeArrayPointer(map)
    If ptrMapSA = 0 Then releaseArrayMap = CHRESULT(E_INVALIDARG): Exit Function
    
    apiCopyMemory cLocks, ByVal ptrMapSA + 8, 4
    Select Case cLocks
        Case 0
            Exit Function
        Case 1
            ' unlock the array
            retVal = apiSafeArrayUnlock(ptrMapSA)
            If retVal.HRESULT <> S_OK Then
                releaseArrayMap = CHRESULT(E_UNEXPECTED)
                Exit Function
            End If
            
            ' make small size
            apiCopyMemory nDimensions, ByVal ptrMapSA, 2
            apiCopyMemory ByVal ptrMapSA + 16 + 0, 1&, 4
            apiCopyMemory ByVal ptrMapSA + 16 + 4, 1&, 4
            If nDimensions >= 2 Then
                apiCopyMemory ByVal ptrMapSA + 16 + 8, 1&, 4
                apiCopyMemory ByVal ptrMapSA + 16 + 12, 1&, 4
            End If
            If nDimensions >= 3 Then
                apiCopyMemory ByVal ptrMapSA + 16 + 16, 1&, 4
                apiCopyMemory ByVal ptrMapSA + 16 + 20, 1&, 4
            End If
            
            ' null out the pointer to the mapped data
            apiCopyMemory ByVal ptrMapSA + 12, 0&, 4
            
            ' allocate some memory so that the SAFEARRAY is in a valid state
            If apiSafeArrayAllocData(ptrMapSA).HRESULT <> S_OK Then
                releaseArrayMap = CHRESULT(E_OUTOFMEMORY)
                Exit Function
            End If

            ' let VB null it out
            Erase map
            Exit Function
        Case Else
            releaseArrayMap = CHRESULT(E_TOO_MANY_LOCKS): Exit Function     ' too many locks
    End Select
End Function


