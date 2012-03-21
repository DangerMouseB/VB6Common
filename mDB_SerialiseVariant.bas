Attribute VB_Name = "mDB_SerialiseVariant"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit
Option Private Module

' error reporting
Private Const MODULE_NAME As String = "mDB_SerialiseVariant"
Private Const MODULE_VERSION As String = "0.0.0.1"

Private Const TYPE_LENGTH As Long = 2
Private Const SIZE_LENGTH As Long = 4
Private Const NULL_LENGTH As Long = 2
Private Const UNICODE_LENGTH As Long = 2
Private Const NDIMENSIONS_LENGTH As Long = 2            ' although the safe array only has up to 60 dimensions, probably could relax this and just use a byte - very slightly smaller files
Private Const DIMENSION_LENGTH As Long = 8
Private Const CELEMENTS_LENGTH As Long = 4
Private Const LLBOUND_LENGTH As Long = 4
Private Const BYTE_LENGTH As Long = 1
Private Const INTEGER_LENGTH As Long = 2
Private Const LONG_LENGTH As Long = 4
Private Const SINGLE_LENGTH As Long = 4
Private Const DOUBLE_LENGTH As Long = 8
Private Const BOOLEAN_LENGTH As Long = 1                   ' boolean can be stored as a byte, even though the variant standard says 2 bytes
Private Const DATE_LENGTH As Long = 8
Private Const CURRENCY_LENGTH As Long = 8

' fixed length types are stored thus:       <typeID>, [data]                                                                        (size is implied in the type so not needed)
' variable length types are stored thus:   <typeID>, <size>, [data]
' arrays are stored thus:                        <typeID>, <size>, [nDims, nDims x (cElements, lLBound), data] - if size = 0 then blank array (do we allow this?)

' <type> is a UINT2
' <size> is a UINT4

' could possibly handle vbUserDefinedType - can get the public type name using TypeName(aVariant) and might need to restrict to well known types e.g. LONGLONG etc
' because I can't access a VB variable pointer, I have to construct arrays using ReDim, so I'm going to limit arrays up to three dimensions as an implementation rather than a protocol restriction


Function DBLengthOfVariantAsBytes(aVariant As Variant) As Long
    Dim SA As SAFEARRAY, i As Long, j As Long, k As Long, dataLength As Long
    
    Const METHOD_NAME As String = "DBLengthOfVariantAsBytes"
    
    Select Case varType(aVariant)
    
        ' fixed length data types
        Case (vbEmpty)
            DBLengthOfVariantAsBytes = TYPE_LENGTH
        Case (vbByte)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + BYTE_LENGTH
        Case (vbInteger)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + INTEGER_LENGTH
        Case (vbLong)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + LONG_LENGTH
        Case (vbSingle)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + SINGLE_LENGTH
        Case (vbDouble)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + DOUBLE_LENGTH
        Case (vbBoolean)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + BOOLEAN_LENGTH
        Case (vbDate)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + DATE_LENGTH
        Case (vbCurrency)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + CURRENCY_LENGTH
            
        ' variable length data types
        Case (vbString)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant) + NULL_LENGTH
            
        ' array of fixed length data types
        Case (vbByte Or vbArray), (vbInteger Or vbArray), (vbLong Or vbArray), (vbSingle Or vbArray), (vbDouble Or vbArray), (vbBoolean Or vbArray), (vbDate Or vbArray), (vbCurrency Or vbArray)
            DBGetSafeArrayDetails aVariant, SA
            If SA.cDims = 0 Then
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH
            Else
                dataLength = SA.cbElements
                If SA.cDims > 3 Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for String()"
                For i = 1& To SA.cDims
                    dataLength = dataLength * SA.rgSABound(i).cElements
                Next
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH + DIMENSION_LENGTH * SA.cDims + dataLength
            End If
            
        ' array of strings
        Case (vbString Or vbArray)
            DBGetSafeArrayDetails aVariant, SA
            If SA.cDims = 0 Then
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH
            Else
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            dataLength = dataLength + SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant(i)) + NULL_LENGTH
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                dataLength = dataLength + SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant(i, j)) + NULL_LENGTH
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    dataLength = dataLength + SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant(i, j, k)) + NULL_LENGTH
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle more than 3 dimensions"
                End Select
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH + DIMENSION_LENGTH * SA.cDims + dataLength
            End If
        
        ' array of variants
        Case (vbVariant Or vbArray)
            DBGetSafeArrayDetails aVariant, SA
            If SA.cDims = 0 Then
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH
            Else
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            dataLength = dataLength + DBLengthOfVariantAsBytes(aVariant(i))
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                dataLength = dataLength + DBLengthOfVariantAsBytes(aVariant(i, j))
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    dataLength = dataLength + DBLengthOfVariantAsBytes(aVariant(i, j, k))
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle more than 3 dimensions"
                End Select
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH + DIMENSION_LENGTH * SA.cDims + dataLength
            End If
            
        Case Else
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle data type " & TypeName(aVariant)
            
    End Select
    
End Function


Sub DBVariantAsBytes(aVariant As Variant, buffer() As Byte, indexOfBufferEndPlusOne As Long, indexOfNext As Long)
    Dim varTypeID As Integer, SA As SAFEARRAY, i As Long, j As Long, k As Long, dataLength As Long
    
    Const METHOD_NAME As String = "DBVariantAsBytes"
    
    varTypeID = varType(aVariant)
    
    Select Case varTypeID
    
        ' fixed length data types
        Case (vbEmpty)
            If indexOfNext + TYPE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
        Case (vbByte)
            If indexOfNext + TYPE_LENGTH + BYTE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, BYTE_LENGTH: indexOfNext = indexOfNext + BYTE_LENGTH
        Case (vbInteger)
            If indexOfNext + TYPE_LENGTH + INTEGER_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, INTEGER_LENGTH: indexOfNext = indexOfNext + INTEGER_LENGTH
        Case (vbLong)
            If indexOfNext + TYPE_LENGTH + LONG_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, LONG_LENGTH: indexOfNext = indexOfNext + LONG_LENGTH
        Case (vbSingle)
            If indexOfNext + TYPE_LENGTH + SINGLE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, SINGLE_LENGTH: indexOfNext = indexOfNext + SINGLE_LENGTH
        Case (vbDouble)
            If indexOfNext + TYPE_LENGTH + DOUBLE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, DOUBLE_LENGTH: indexOfNext = indexOfNext + DOUBLE_LENGTH
        Case (vbBoolean)
            If indexOfNext + TYPE_LENGTH + BOOLEAN_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"        ' TRUE -> 255, FALSE -> 0
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, BOOLEAN_LENGTH: indexOfNext = indexOfNext + BOOLEAN_LENGTH
        Case (vbDate)
            If indexOfNext + TYPE_LENGTH + DATE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, DATE_LENGTH: indexOfNext = indexOfNext + DATE_LENGTH
        Case (vbCurrency)
            If indexOfNext + TYPE_LENGTH + CURRENCY_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(aVariant) + 8&, CURRENCY_LENGTH: indexOfNext = indexOfNext + CURRENCY_LENGTH
            
        ' variable length data types
        Case (vbString)
            dataLength = SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant) + NULL_LENGTH
            If indexOfNext + TYPE_LENGTH + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal StrPtr(aVariant) - 4, dataLength: indexOfNext = indexOfNext + dataLength
            
        ' array of fixed length data types
        Case (vbByte Or vbArray), (vbInteger Or vbArray), (vbLong Or vbArray), (vbSingle Or vbArray), (vbDouble Or vbArray), (vbBoolean Or vbArray), (vbDate Or vbArray), (vbCurrency Or vbArray)
            DBGetSafeArrayDetails aVariant, SA
            If SA.cDims = 0 Then
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Else
                dataLength = SA.cbElements
                If SA.cDims > 3 Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for String()"
                For i = 1& To SA.cDims
                    dataLength = dataLength * SA.rgSABound(i).cElements
                Next
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH + DIMENSION_LENGTH * SA.cDims + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
                For i = 1& To SA.cDims
                    apiCopyMemory buffer(indexOfNext), SA.rgSABound(i).cElements, DIMENSION_LENGTH: indexOfNext = indexOfNext + DIMENSION_LENGTH            ' copys both cElements and lLbound
                Next
                apiCopyMemory buffer(indexOfNext), ByVal SA.pvData, dataLength: indexOfNext = indexOfNext + dataLength
            End If

        ' array of strings
        Case (vbString Or vbArray)
            DBGetSafeArrayDetails aVariant, SA
            If SA.cDims = 0 Then
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Else
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH + DIMENSION_LENGTH * SA.cDims > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
                For i = 1& To SA.cDims
                    apiCopyMemory buffer(indexOfNext), SA.rgSABound(i).cElements, DIMENSION_LENGTH: indexOfNext = indexOfNext + DIMENSION_LENGTH            ' copys both cElements and lLbound
                Next
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            dataLength = SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant(i)) + NULL_LENGTH
                            If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            apiCopyMemory buffer(indexOfNext), ByVal StrPtr(aVariant(i)) - 4, dataLength: indexOfNext = indexOfNext + dataLength
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                dataLength = SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant(i, j)) + NULL_LENGTH
                                If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                apiCopyMemory buffer(indexOfNext), ByVal StrPtr(aVariant(i, j)) - 4, dataLength: indexOfNext = indexOfNext + dataLength
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    dataLength = SIZE_LENGTH + UNICODE_LENGTH * Len(aVariant(i, j, k)) + NULL_LENGTH
                                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                    apiCopyMemory buffer(indexOfNext), ByVal StrPtr(aVariant(i, j, k)) - 4, dataLength: indexOfNext = indexOfNext + dataLength
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for String()"
                End Select
            End If
            
        ' array of variants
        Case (vbVariant Or vbArray)
            DBGetSafeArrayDetails aVariant, SA
            If SA.cDims = 0 Then
                If indexOfNext + TYPE_LENGTH + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Else
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH + DIMENSION_LENGTH * SA.cDims > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), varTypeID, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
                For i = 1& To SA.cDims
                    apiCopyMemory buffer(indexOfNext), SA.rgSABound(i).cElements, DIMENSION_LENGTH: indexOfNext = indexOfNext + DIMENSION_LENGTH            ' copys both cElements and lLbound
                Next
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            DBVariantAsBytes aVariant(i), buffer, indexOfBufferEndPlusOne, indexOfNext
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                DBVariantAsBytes aVariant(i, j), buffer, indexOfBufferEndPlusOne, indexOfNext
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    DBVariantAsBytes aVariant(i, j, k), buffer, indexOfBufferEndPlusOne, indexOfNext
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for Variant()"
                End Select
            End If
            
        Case Else
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise data type " & TypeName(aVariant)
            
    End Select
    
End Sub


Function DBBytesAsVariant(buffer() As Byte, indexOfBufferEndPlusOne As Long, indexOfNext As Long) As Variant
    Dim variantType As Integer, i As Long, j As Long, k As Long, dataLength As Long, SA As SAFEARRAY
    Dim aByte As Byte, anInteger As Integer, aLong As Long, aSingle As Single, aDouble As Double, aBoolean As Boolean, aDate As Date, aCurrency As Currency, aString As String, anArray As Variant
    Dim arrayType As Integer, nDimensions As Long, I1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long
    
    Const METHOD_NAME As String = "DBBytesAsVariant"
    
    If indexOfNext + TYPE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
    apiCopyMemory variantType, buffer(indexOfNext), TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
    
    Select Case variantType
    
        ' fixed length data types
        Case (vbEmpty)
            DBBytesAsVariant = Empty
        Case (vbByte)
            If indexOfNext + BYTE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aByte), buffer(indexOfNext), BYTE_LENGTH: indexOfNext = indexOfNext + BYTE_LENGTH
            DBBytesAsVariant = aByte
        Case (vbInteger)
            If indexOfNext + INTEGER_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(anInteger), buffer(indexOfNext), INTEGER_LENGTH: indexOfNext = indexOfNext + INTEGER_LENGTH
            DBBytesAsVariant = anInteger
        Case (vbLong)
            If indexOfNext + LONG_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aLong), buffer(indexOfNext), LONG_LENGTH: indexOfNext = indexOfNext + LONG_LENGTH
            DBBytesAsVariant = aLong
        Case (vbSingle)
            If indexOfNext + SINGLE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aSingle), buffer(indexOfNext), SINGLE_LENGTH: indexOfNext = indexOfNext + SINGLE_LENGTH
            DBBytesAsVariant = aSingle
        Case (vbDouble)
            If indexOfNext + DOUBLE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aDouble), buffer(indexOfNext), DOUBLE_LENGTH: indexOfNext = indexOfNext + DOUBLE_LENGTH
            DBBytesAsVariant = aDouble
        Case (vbBoolean)
            If indexOfNext + BOOLEAN_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"         ' TRUE -> 255, FALSE -> 0
            apiCopyMemory ByVal VarPtr(aByte), buffer(indexOfNext), BYTE_LENGTH: indexOfNext = indexOfNext + BOOLEAN_LENGTH
            If aByte = 0 Then DBBytesAsVariant = False Else DBBytesAsVariant = True
        Case (vbDate)
            If indexOfNext + DATE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aDate), buffer(indexOfNext), DATE_LENGTH: indexOfNext = indexOfNext + DATE_LENGTH
            DBBytesAsVariant = aDate
        Case (vbCurrency)
            If indexOfNext + CURRENCY_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aCurrency), buffer(indexOfNext), CURRENCY_LENGTH: indexOfNext = indexOfNext + CURRENCY_LENGTH
            DBBytesAsVariant = aCurrency
            
        ' variable length data types
        Case (vbString)
            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
            If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            aString = String(dataLength / 2, " ")
            apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + NULL_LENGTH
            DBBytesAsVariant = aString
            
        ' array of fixed length data types
        Case (vbByte Or vbArray), (vbInteger Or vbArray), (vbLong Or vbArray), (vbSingle Or vbArray), (vbDouble Or vbArray), (vbBoolean Or vbArray), (vbDate Or vbArray), (vbCurrency Or vbArray)
            arrayType = Not (vbArray) And variantType
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                    DBCreateNewArrayOfType anArray, arrayType, 1, 1
                    Erase anArray
                Case 1
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewArrayOfType anArray, arrayType, I1, i2
                    DBGetSafeArrayDetails anArray, SA
                    dataLength = (i2 - I1 + 1) * SA.cbElements
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory ByVal SA.pvData, buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength
                Case 2
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewArrayOfType anArray, arrayType, I1, i2, j1, j2
                    DBGetSafeArrayDetails anArray, SA
                    dataLength = (i2 - I1 + 1) * (j2 - j1 + 1) * SA.cbElements
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory ByVal SA.pvData, buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength
                Case 3
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory k1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewArrayOfType anArray, arrayType, I1, i2, j1, j2, k1, k2
                    DBGetSafeArrayDetails anArray, SA
                    dataLength = (i2 - I1 + 1) * (j2 - j1 + 1) * (k2 - k1 + 1) * SA.cbElements
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory ByVal SA.pvData, buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for " & nameFromVBType(variantType)
            End Select
            DBBytesAsVariant = anArray

        ' array of strings
        Case (vbString Or vbArray)
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                    DBCreateNewArrayOfType anArray, vbString, 1, 1
                    Erase anArray
                Case 1
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewArrayOfType anArray, vbString, I1, i2
                    For i = I1 To i2
                        If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                        If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        aString = String(dataLength / 2, " ")
                        apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + NULL_LENGTH
                        anArray(i) = aString
                    Next
                Case 2
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewArrayOfType anArray, vbString, I1, i2, j1, j2
                    For i = I1 To i2
                        For j = j1 To j2
                            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                            If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            aString = String(dataLength / 2, " ")
                            apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + NULL_LENGTH
                            anArray(i, j) = aString
                        Next
                    Next
                Case 3
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory k1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewArrayOfType anArray, vbString, I1, i2, j1, j2, k1, k2
                    For i = I1 To i2
                        For j = j1 To j2
                            For k = k1 To k2
                                If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                                If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                aString = String(dataLength / 2, " ")
                                apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + NULL_LENGTH
                                anArray(i, j, k) = aString
                            Next
                        Next
                    Next
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for String()"
            End Select
            DBBytesAsVariant = anArray

        ' array of variants
        Case (vbVariant Or vbArray)
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                    DBCreateNewVariantArray anArray, 1, 1
                    Erase anArray
                Case 1
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewVariantArray anArray, I1, i2
                    For i = I1 To i2
                        anArray(i) = DBBytesAsVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                    Next
                Case 2
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewVariantArray anArray, I1, i2, j1, j2
                    For i = I1 To i2
                        For j = j1 To j2
                            anArray(i, j) = DBBytesAsVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                        Next
                    Next
                Case 3
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory k1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    DBCreateNewVariantArray anArray, I1, i2, j1, j2, k1, k2
                    For i = I1 To i2
                        For j = j1 To j2
                            For k = k1 To k2
                                anArray(i, j, k) = DBBytesAsVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                            Next
                        Next
                    Next
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for Variant()"
            End Select
            DBBytesAsVariant = anArray

        Case Else
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise data type " & nameFromVBType(variantType)
            
    End Select
    
End Function


Function DBVerifyStructureOfSerialisedVariant(buffer() As Byte, indexOfBufferEndPlusOne As Long, indexOfNext As Long) As Boolean
    Dim variantType As Integer, i As Long, j As Long, k As Long, dataLength As Long, nullCheck As Long
    Dim arrayType As Integer, nDimensions As Long, I1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long
    
    Const METHOD_NAME As String = "DBVerifyStructureOfSerialisedVariant"
    
    If indexOfNext + TYPE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
    apiCopyMemory variantType, buffer(indexOfNext), TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
    
    Select Case variantType
    
        ' fixed length data types
        Case (vbEmpty)
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbByte)
            If indexOfNext + BYTE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + BYTE_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbInteger)
            If indexOfNext + INTEGER_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + INTEGER_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbLong)
            If indexOfNext + LONG_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + LONG_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbSingle)
            If indexOfNext + SINGLE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + SINGLE_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbDouble)
            If indexOfNext + DOUBLE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + DOUBLE_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbBoolean)
            If indexOfNext + BOOLEAN_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"         ' TRUE -> 255, FALSE -> 0
            indexOfNext = indexOfNext + BOOLEAN_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbDate)
            If indexOfNext + DATE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + DATE_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case (vbCurrency)
            If indexOfNext + CURRENCY_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + CURRENCY_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
            
        ' variable length data types
        Case (vbString)
            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
            If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + dataLength
            apiCopyMemory nullCheck, buffer(indexOfNext), NULL_LENGTH: indexOfNext = indexOfNext + NULL_LENGTH
            DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
            
        ' array of fixed length data types
        Case (vbByte Or vbArray), (vbInteger Or vbArray), (vbLong Or vbArray), (vbSingle Or vbArray), (vbDouble Or vbArray), (vbBoolean Or vbArray), (vbDate Or vbArray), (vbCurrency Or vbArray)
            arrayType = Not (vbArray) And variantType
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                Case 1
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    dataLength = (i2 - I1 + 1) * lengthOfVbType(arrayType)
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    indexOfNext = indexOfNext + dataLength
                Case 2
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    dataLength = (i2 - I1 + 1) * (j2 - j1 + 1) * lengthOfVbType(arrayType)
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    indexOfNext = indexOfNext + dataLength
                Case 3
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory k1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    dataLength = (i2 - I1 + 1) * (j2 - j1 + 1) * (k2 - k1 + 1) * lengthOfVbType(arrayType)
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    indexOfNext = indexOfNext + dataLength
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for " & nameFromVBType(variantType)
            End Select
            DBVerifyStructureOfSerialisedVariant = True

        ' array of strings
        Case (vbString Or vbArray)
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                Case 1
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    For i = I1 To i2
                        If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                        If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        indexOfNext = indexOfNext + dataLength
                        apiCopyMemory nullCheck, buffer(indexOfNext), NULL_LENGTH: indexOfNext = indexOfNext + NULL_LENGTH
                        DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
                    Next
                Case 2
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    For i = I1 To i2
                        For j = j1 To j2
                            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                            If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            indexOfNext = indexOfNext + dataLength
                            apiCopyMemory nullCheck, buffer(indexOfNext), NULL_LENGTH: indexOfNext = indexOfNext + NULL_LENGTH
                            DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
                        Next
                    Next
                Case 3
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory k1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    For i = I1 To i2
                        For j = j1 To j2
                            For k = k1 To k2
                                If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                                If indexOfNext + dataLength + NULL_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                indexOfNext = indexOfNext + dataLength
                                apiCopyMemory nullCheck, buffer(indexOfNext), NULL_LENGTH: indexOfNext = indexOfNext + NULL_LENGTH
                                DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
                            Next
                        Next
                    Next
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for String()"
            End Select
            DBVerifyStructureOfSerialisedVariant = True

        ' array of variants
        Case (vbVariant Or vbArray)
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                Case 1
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    For i = I1 To i2
                        DBVerifyStructureOfSerialisedVariant = DBVerifyStructureOfSerialisedVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                        If DBVerifyStructureOfSerialisedVariant = False Then Exit Function
                    Next
                Case 2
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    For i = I1 To i2
                        For j = j1 To j2
                            DBVerifyStructureOfSerialisedVariant = DBVerifyStructureOfSerialisedVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                             If DBVerifyStructureOfSerialisedVariant = False Then Exit Function
                        Next
                    Next
                Case 3
                    If indexOfNext + DIMENSION_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory k1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory j1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + CELEMENTS_LENGTH
                    apiCopyMemory I1, buffer(indexOfNext), CELEMENTS_LENGTH: indexOfNext = indexOfNext + LLBOUND_LENGTH
                    i2 = I1 + i2 - 1
                    For i = I1 To i2
                        For j = j1 To j2
                            For k = k1 To k2
                                DBVerifyStructureOfSerialisedVariant = DBVerifyStructureOfSerialisedVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                                 If DBVerifyStructureOfSerialisedVariant = False Then Exit Function
                            Next
                        Next
                    Next
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for Variant()"
            End Select
            
        Case Else
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise data type " & nameFromVBType(variantType)
            
    End Select
    
End Function

Private Function nameFromVBType(vbType As Integer) As String
    Select Case Not (vbArray) And vbType
        Case vbByte
            nameFromVBType = "Byte"
        Case vbInteger
            nameFromVBType = "Integer"
        Case vbLong
            nameFromVBType = "Long"
        Case vbSingle
            nameFromVBType = "Single"
        Case vbDouble
            nameFromVBType = "Double"
        Case vbBoolean
            nameFromVBType = "Boolean"
        Case vbDate
            nameFromVBType = "Date"
        Case vbCurrency
            nameFromVBType = "Currency"
        Case vbString
            nameFromVBType = "String"
        Case vbVariant
            nameFromVBType = "Variant"
        Case Else
    End Select
    If vbArray And vbType Then nameFromVBType = nameFromVBType & "()"
End Function

Private Function lengthOfVbType(vbType As Integer) As Long
    Select Case Not (vbArray) And vbType
        Case vbByte
            lengthOfVbType = 1
        Case vbInteger
            lengthOfVbType = 2
        Case vbLong
            lengthOfVbType = 4
        Case vbSingle
            lengthOfVbType = 4
        Case vbDouble
            lengthOfVbType = 8
        Case vbBoolean
            lengthOfVbType = 2
        Case vbDate
            lengthOfVbType = 8
        Case vbCurrency
            lengthOfVbType = 8
        Case Else
            lengthOfVbType = -1
    End Select
End Function


Sub varTest1()
    Dim c As String, d As Variant, buffer() As Byte, buf2() As Byte, ptr1 As Long, ptr2 As Long, ptr3 As Long, ptr4 As Long, ptr5 As Long
    c = "Hello"
    d = c
    buffer = getBuffer(d)
    apiCopyMemory ptr1, buffer(9), 4
    apiCopyMemory ptr2, ByVal ptr1, 4
    ptr3 = VarPtr(c)
    apiCopyMemory ptr4, ByVal ptr3, 4
    ptr5 = StrPtr(c)
    ReDim buf2(1 To Len(c) * 2 + NULL_LENGTH + SIZE_LENGTH)
    apiCopyMemory buf2(1), ByVal StrPtr(c) - 4, Len(c) * 2 + NULL_LENGTH + SIZE_LENGTH
    Debug.Print ptr1, ptr2, ptr3, ptr4, VarPtr(d), StrPtr(d)
    Stop
    
End Sub

Sub varTest2()
    Dim a() As Byte, b(9 To 10) As Integer, c(7 To 8) As Long, d(5 To 6) As Single, e(-1 To 0) As Double, f(1 To 10) As Boolean, g(1 To 10) As Date, h(1 To 10) As Currency, i(1 To 2) As String, j As Variant, length As Long
    Dim x As Long, buffer() As Byte, result As Variant
    ReDim a(1 To 10) As Byte
    
    For x = 1 To 2
        i(x) = "hello" & x
        e(x - 2) = CDbl(x) / 3.14
    Next
    
    j = Array(CByte(1), CInt(2), CLng(3), CSng(4), CDbl(5), True, CDate("7/7/07"), CCur(8), "9", i, e, a, b, c, d)
    j = Array(j, j)
    
    length = DBLengthOfVariantAsBytes(j)
    ReDim buffer(1 To length) As Byte
    DBVariantAsBytes j, buffer, length + 1, 1
    If DBVerifyStructureOfSerialisedVariant(buffer, length + 1, 1) = False Then Stop
    result = DBBytesAsVariant(buffer, length + 1, 1)
    
    Stop
'    Debug.Print "Byte: " & DBLengthOfVariantAsBytes(CByte(8))
'    Debug.Print "Integer: " & DBLengthOfVariantAsBytes(CInt(8))
'    Debug.Print "Long: " & DBLengthOfVariantAsBytes(CLng(8))
'    Debug.Print "Single: " & DBLengthOfVariantAsBytes(CSng(8))
'    Debug.Print "Double: " & DBLengthOfVariantAsBytes(CDbl(8))
'    Debug.Print "Boolean: " & DBLengthOfVariantAsBytes(True)
'    Debug.Print "Date: " & DBLengthOfVariantAsBytes(CDate(8))
'    Debug.Print "Currency: " & DBLengthOfVariantAsBytes(CCur(8))
'    Debug.Print "String: " & DBLengthOfVariantAsBytes("hello")
'
'    Debug.Print "Byte(): " & DBLengthOfVariantAsBytes(a)
'    Debug.Print "Integer(): " & DBLengthOfVariantAsBytes(b)
'    Debug.Print "Long(): " & DBLengthOfVariantAsBytes(c)
'    Debug.Print "Single(): " & DBLengthOfVariantAsBytes(d)
'    Debug.Print "Double(): " & DBLengthOfVariantAsBytes(e)
'    Debug.Print "Boolean(): " & DBLengthOfVariantAsBytes(f)
'    Debug.Print "Date(): " & DBLengthOfVariantAsBytes(g)
'    Debug.Print "Currency(): " & DBLengthOfVariantAsBytes(h)
'    Debug.Print "String(): " & DBLengthOfVariantAsBytes(i)
'    Debug.Print "Variant(): " & length
'
'    Stop
'    Erase a
'    Debug.Print "Empty Byte(): " & DBLengthOfVariantAsBytes(a)
'    Debug.Print "Variant(Empty Byte()): " & DBLengthOfVariantAsBytes(Array(a))
End Sub

Function getBuffer(aVar As Variant) As Variant
    Dim buffer(1 To 16) As Byte
    apiCopyMemory buffer(1&), aVar, 16
    getBuffer = buffer
End Function


'*************************************************************************************************************************************************************************************************************************************************
' module summary
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function

