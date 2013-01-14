Attribute VB_Name = "mDB_BLIP"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011 David Briant - see https://github.com/DangerMouseB
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Lesser General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'************************************************************************************************************************************************
 
Option Explicit
Option Private Module

' error reporting
Private Const MODULE_NAME As String = "mDB_BLIP"
Private Const MODULE_VERSION As String = "0.0.0.2"

Private Const TYPE_LENGTH As Long = 1
Private Const SIZE_LENGTH As Long = 4
Private Const UINT8_LENGTH As Long = 1
Private Const INT16_LENGTH As Long = 2
Private Const INT32_LENGTH As Long = 4
Private Const INT64_LENGTH As Long = 8
Private Const FLOAT32_LENGTH As Long = 4
Private Const FLOAT64_LENGTH As Long = 8
Private Const BOOLEAN_LENGTH As Long = 1
Private Const XLDATE64_LENGTH As Long = 8
Private Const UNICODE_NULL_TERMINATOR_LENGTH As Long = 2
Private Const UNICODE_WIDTH As Long = 2
Private Const NDIMENSIONS_LENGTH As Long = 1
Private Const DIM_LENGTH As Long = 4


' BLIP TYPES
Private Const BT_MISSING As Long = &H0
'Private Const BT_NULL As Long = &H1
Private Const BT_UINT8 As Long = &H2
'Private Const BT_INT8 As Long = &H3
'Private Const BT_UINT16 As Long = &H4
Private Const BT_INT16 As Long = &H5
'Private Const BT_UINT32 As Long = &H6
Private Const BT_INT32 As Long = &H7
'Private Const BT_UINT64 As Long = &H8
Private Const BT_INT64 As Long = &H9
Private Const BT_FLOAT32 As Long = &HA
Private Const BT_FLOAT64 As Long = &HB
Private Const BT_BOOLEAN As Long = &HC
Private Const BT_XLDATE64 As Long = &HD
Private Const BT_UNICODE16 As Long = &H11
Private Const BT_DICTIONARY As Long = &H7F
Private Const BT_ARRAY As Long = &H80
Private Const BT_ARRAY_UINT8 As Long = &H82
Private Const BT_ARRAY_INT16 As Long = &H85
Private Const BT_ARRAY_INT32 As Long = &H87
Private Const BT_ARRAY_INT64 As Long = &H89
Private Const BT_ARRAY_FLOAT32 As Long = &H8A
Private Const BT_ARRAY_FLOAT64 As Long = &H8B
Private Const BT_ARRAY_BOOLEAN As Long = &H8C
Private Const BT_ARRAY_XLDATE64 As Long = &H8D
Private Const BT_ARRAY_UNICODE16 As Long = &H91


' fixed length types are stored thus:       <typeID>, [data]                                                                        (size is implied in the type so not needed)
' variable length types are stored thus:   <typeID>, <size>, [data]
' arrays are stored thus:                        <typeID>, <size>, [nDims, nDims x (cElements, lLBound), data] - if size = 0 then blank array (do we allow this?)

' <type> is a UINT8
' <size> is a UINT32

' could possibly handle vbUserDefinedType - can get the public type name using TypeName(var) and might need to restrict to well known types e.g. LONGLONG etc
' because I can't access a VB variable pointer, I have to construct arrays using ReDim, so I'm going to limit arrays up to three dimensions as an implementation rather than a protocol restriction
' could add cycle detection due to using VB dict and numpy ndarray of object_ - just an rror message rather than format change


Function DBLengthOfVariantAsBytes(var As Variant) As Long
    Dim SA As SAFEARRAY, i As Long, j As Long, k As Long, dataLength As Long
    
    Const METHOD_NAME As String = "DBLengthOfVariantAsBytes"
    
    Select Case VarType(var)
    
        ' fixed length data types
        Case (vbEmpty)
            DBLengthOfVariantAsBytes = TYPE_LENGTH
        Case (vbByte)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + UINT8_LENGTH
        Case (vbInteger)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + INT16_LENGTH
        Case (vbLong)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + INT32_LENGTH
        Case (vbCurrency)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + INT64_LENGTH
        Case (vbSingle)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + FLOAT32_LENGTH
        Case (vbDouble)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + FLOAT64_LENGTH
        Case (vbBoolean)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + BOOLEAN_LENGTH
        Case (vbDate)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + XLDATE64_LENGTH
            
        ' variable length data types
        Case (vbString)
            DBLengthOfVariantAsBytes = TYPE_LENGTH + SIZE_LENGTH + UNICODE_WIDTH * Len(var) + UNICODE_NULL_TERMINATOR_LENGTH
            
        ' array of fixed length data types
        Case (vbByte Or vbArray), (vbInteger Or vbArray), (vbLong Or vbArray), (vbSingle Or vbArray), (vbDouble Or vbArray), (vbBoolean Or vbArray), (vbDate Or vbArray), (vbCurrency Or vbArray)
            DBGetSafeArrayDetails var, SA
            If SA.cDims = 0 Then
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH
            Else
                dataLength = SA.cbElements
                If SA.cDims > 3 Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for String()"
                For i = 1& To SA.cDims
                    dataLength = dataLength * SA.rgSABound(i).cElements
                Next
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH + DIM_LENGTH * SA.cDims + dataLength
            End If
            
        ' Dictionary
        Case (vbObject)
            If TypeName(var) = "Dictionary" Then
                ' size is length of each string key as ascii + length of value
                ' raise error if not string keys
                DBErrors_raiseNotYetImplemented ModuleSummary(), METHOD_NAME, "Dictionary"
            Else
                DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle vbObject containing type " & TypeName(var)
            End If
                
        ' array of strings
        Case (vbString Or vbArray)
            DBGetSafeArrayDetails var, SA
            If SA.cDims = 0 Then
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH
            Else
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            dataLength = dataLength + SIZE_LENGTH + UNICODE_WIDTH * Len(var(i)) + UNICODE_NULL_TERMINATOR_LENGTH
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                dataLength = dataLength + SIZE_LENGTH + UNICODE_WIDTH * Len(var(i, j)) + UNICODE_NULL_TERMINATOR_LENGTH
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    dataLength = dataLength + SIZE_LENGTH + UNICODE_WIDTH * Len(var(i, j, k)) + UNICODE_NULL_TERMINATOR_LENGTH
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle more than 3 dimensions"
                End Select
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH + DIM_LENGTH * SA.cDims + dataLength
            End If
        
        ' array of variants
        Case (vbVariant Or vbArray)
            DBGetSafeArrayDetails var, SA
            If SA.cDims = 0 Then
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH
            Else
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            dataLength = dataLength + DBLengthOfVariantAsBytes(var(i))
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                dataLength = dataLength + DBLengthOfVariantAsBytes(var(i, j))
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    dataLength = dataLength + DBLengthOfVariantAsBytes(var(i, j, k))
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle more than 3 dimensions"
                End Select
                DBLengthOfVariantAsBytes = TYPE_LENGTH + NDIMENSIONS_LENGTH + DIM_LENGTH * SA.cDims + dataLength
            End If
            
        Case Else
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle data type " & TypeName(var)
            
    End Select
    
End Function


Sub DBVariantAsBytes(var As Variant, buffer() As Byte, indexOfBufferEndPlusOne As Long, indexOfNext As Long)
    Dim SA As SAFEARRAY, i As Long, j As Long, k As Long, dataLength As Long
    
    Const METHOD_NAME As String = "DBVariantAsBytes"
       
    Select Case VarType(var)
    
        ' fixed length data types
        Case (vbEmpty)
            If indexOfNext + TYPE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_MISSING, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
        Case (vbByte)
            If indexOfNext + TYPE_LENGTH + UINT8_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_UINT8, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, UINT8_LENGTH: indexOfNext = indexOfNext + UINT8_LENGTH
        Case (vbInteger)
            If indexOfNext + TYPE_LENGTH + INT16_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_INT16, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, INT16_LENGTH: indexOfNext = indexOfNext + INT16_LENGTH
        Case (vbLong)
            If indexOfNext + TYPE_LENGTH + INT32_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_INT32, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, INT32_LENGTH: indexOfNext = indexOfNext + INT32_LENGTH
        Case (vbCurrency)
            If indexOfNext + TYPE_LENGTH + INT64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_INT64, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, INT64_LENGTH: indexOfNext = indexOfNext + INT64_LENGTH
        Case (vbSingle)
            If indexOfNext + TYPE_LENGTH + FLOAT32_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_FLOAT32, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, FLOAT32_LENGTH: indexOfNext = indexOfNext + FLOAT32_LENGTH
        Case (vbDouble)
            If indexOfNext + TYPE_LENGTH + FLOAT64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_FLOAT64, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, FLOAT64_LENGTH: indexOfNext = indexOfNext + FLOAT64_LENGTH
        Case (vbBoolean)
            If indexOfNext + TYPE_LENGTH + BOOLEAN_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"        ' TRUE -> 255, FALSE -> 0
            apiCopyMemory buffer(indexOfNext), BT_BOOLEAN, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, BOOLEAN_LENGTH: indexOfNext = indexOfNext + BOOLEAN_LENGTH
        Case (vbDate)
            If indexOfNext + TYPE_LENGTH + XLDATE64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_XLDATE64, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal VarPtr(var) + 8&, XLDATE64_LENGTH: indexOfNext = indexOfNext + XLDATE64_LENGTH
            
        ' variable length data types
        Case (vbString)
            dataLength = SIZE_LENGTH + UNICODE_WIDTH * Len(var) + UNICODE_NULL_TERMINATOR_LENGTH
            If indexOfNext + TYPE_LENGTH + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory buffer(indexOfNext), BT_UNICODE16, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
            apiCopyMemory buffer(indexOfNext), ByVal StrPtr(var) - 4, dataLength: indexOfNext = indexOfNext + dataLength
            
        ' array of fixed length data types
        Case (vbByte Or vbArray), (vbInteger Or vbArray), (vbLong Or vbArray), (vbSingle Or vbArray), (vbDouble Or vbArray), (vbBoolean Or vbArray), (vbDate Or vbArray), (vbCurrency Or vbArray)
            DBGetSafeArrayDetails var, SA
            If SA.cDims = 0 Then
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), BTFromVBType(VarType(var)), TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Else
                dataLength = SA.cbElements
                If SA.cDims > 3 Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for String()"
                For i = 1& To SA.cDims
                    dataLength = dataLength * SA.rgSABound(i).cElements
                Next
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH + DIM_LENGTH * SA.cDims + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), BTFromVBType(VarType(var)), TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
                For i = 1& To SA.cDims
                    apiCopyMemory buffer(indexOfNext), SA.rgSABound(i).cElements, DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                Next
                apiCopyMemory buffer(indexOfNext), ByVal SA.pvData, dataLength: indexOfNext = indexOfNext + dataLength
            End If

        ' vbObject containing a Dictionary
        Case (vbObject)
            If TypeName(var) = "Dictionary" Then
                ' size is length of each string key as ascii + length of value
                ' raise error if not string keys
                DBErrors_raiseNotYetImplemented ModuleSummary(), METHOD_NAME, "Dictionary"
            Else
                DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't handle vbObject containing type " & TypeName(var)
            End If
                
        ' array of strings
        Case (vbString Or vbArray)
            DBGetSafeArrayDetails var, SA
            If SA.cDims = 0 Then
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), BT_ARRAY_UNICODE16, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Else
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH + DIM_LENGTH * SA.cDims > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), BT_ARRAY_UNICODE16, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
                For i = 1& To SA.cDims
                    apiCopyMemory buffer(indexOfNext), SA.rgSABound(i).cElements, DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                Next
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            dataLength = SIZE_LENGTH + UNICODE_WIDTH * Len(var(i)) + UNICODE_NULL_TERMINATOR_LENGTH
                            If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            apiCopyMemory buffer(indexOfNext), ByVal StrPtr(var(i)) - 4, dataLength: indexOfNext = indexOfNext + dataLength
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                dataLength = SIZE_LENGTH + UNICODE_WIDTH * Len(var(i, j)) + UNICODE_NULL_TERMINATOR_LENGTH
                                If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                apiCopyMemory buffer(indexOfNext), ByVal StrPtr(var(i, j)) - 4, dataLength: indexOfNext = indexOfNext + dataLength
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    dataLength = SIZE_LENGTH + UNICODE_WIDTH * Len(var(i, j, k)) + UNICODE_NULL_TERMINATOR_LENGTH
                                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                    apiCopyMemory buffer(indexOfNext), ByVal StrPtr(var(i, j, k)) - 4, dataLength: indexOfNext = indexOfNext + dataLength
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for String()"
                End Select
            End If
            
        ' array of variants
        Case (vbVariant Or vbArray)
            DBGetSafeArrayDetails var, SA
            If SA.cDims = 0 Then
                If indexOfNext + TYPE_LENGTH + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), BT_ARRAY, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Else
                If indexOfNext + TYPE_LENGTH + NDIMENSIONS_LENGTH + DIM_LENGTH * SA.cDims > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                apiCopyMemory buffer(indexOfNext), BT_ARRAY, TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
                apiCopyMemory buffer(indexOfNext), SA.cDims, NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
                For i = 1& To SA.cDims
                    apiCopyMemory buffer(indexOfNext), SA.rgSABound(i).cElements, DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                Next
                Select Case SA.cDims
                    Case 1
                        For i = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                            DBVariantAsBytes var(i), buffer, indexOfBufferEndPlusOne, indexOfNext
                        Next
                    Case 2
                        For i = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                            For j = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                DBVariantAsBytes var(i, j), buffer, indexOfBufferEndPlusOne, indexOfNext
                            Next
                        Next
                    Case 3
                        For i = SA.rgSABound(3&).lLbound To SA.rgSABound(3&).lLbound + SA.rgSABound(3&).cElements - 1&
                            For j = SA.rgSABound(2&).lLbound To SA.rgSABound(2&).lLbound + SA.rgSABound(2&).cElements - 1&
                                For k = SA.rgSABound(1&).lLbound To SA.rgSABound(1&).lLbound + SA.rgSABound(1&).cElements - 1&
                                    DBVariantAsBytes var(i, j, k), buffer, indexOfBufferEndPlusOne, indexOfNext
                                Next
                            Next
                        Next
                    Case Else
                        DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise more than 3 dimensions for Variant()"
                End Select
            End If
            
        Case Else
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't serialise data type " & TypeName(var)
            
    End Select
    
End Sub


Private Function BTFromVBType(VT As Integer) As Byte
    Select Case VT
        Case (vbByte Or vbArray)
            BTFromVBType = BT_ARRAY_UINT8
        Case (vbInteger Or vbArray)
            BTFromVBType = BT_ARRAY_INT16
        Case (vbLong Or vbArray)
            BTFromVBType = BT_ARRAY_INT32
        Case (vbCurrency Or vbArray)
            BTFromVBType = BT_ARRAY_INT64
        Case (vbSingle Or vbArray)
            BTFromVBType = BT_ARRAY_FLOAT32
        Case (vbDouble Or vbArray)
            BTFromVBType = BT_ARRAY_FLOAT64
        Case (vbBoolean Or vbArray)
            BTFromVBType = BT_ARRAY_BOOLEAN
        Case (vbDate Or vbArray)
            BTFromVBType = BT_ARRAY_XLDATE64
        Case (vbString Or vbArray)
            BTFromVBType = BT_ARRAY_UNICODE16
        Case Else
            DBErrors_raiseNotYetImplemented ModuleSummary(), "BTFromVBType", CStr(VT)
    End Select
End Function


Private Function VBArrayTypeFromBT(BT As Byte) As Integer
    Select Case BT
        Case BT_ARRAY_UINT8
            VBArrayTypeFromBT = vbByte
        Case BT_ARRAY_INT16
            VBArrayTypeFromBT = vbInteger
        Case BT_ARRAY_INT32
            VBArrayTypeFromBT = vbLong
        Case BT_ARRAY_INT64
            VBArrayTypeFromBT = vbCurrency
        Case BT_ARRAY_FLOAT32
            VBArrayTypeFromBT = vbSingle
        Case BT_ARRAY_FLOAT64
            VBArrayTypeFromBT = vbDouble
        Case BT_ARRAY_BOOLEAN
            VBArrayTypeFromBT = vbBoolean
        Case BT_ARRAY_XLDATE64
            VBArrayTypeFromBT = vbDate
        Case BT_ARRAY_UNICODE16
            VBArrayTypeFromBT = vbString
        Case Else
            DBErrors_raiseNotYetImplemented ModuleSummary(), "VBArrayTypeFromBT", CStr(BT)
    End Select
End Function



Sub DBBytesAsVariant(buffer() As Byte, indexOfBufferEndPlusOne As Long, indexOfNext As Long, var As Variant, base As Long)
    Dim BT As Byte, i As Long, j As Long, k As Long, dataLength As Long, SA As SAFEARRAY
    Dim aByte As Byte, anInteger As Integer, aLong As Long, aSingle As Single, aDouble As Double, aBoolean As Boolean, aDate As Date, aCurrency As Currency, aString As String, anArray As Variant
    Dim arrayType As Integer, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long
    
    Const METHOD_NAME As String = "DBBytesAsVariant"
    
    If indexOfNext + TYPE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
    apiCopyMemory BT, buffer(indexOfNext), TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
    
    Select Case BT
    
        ' fixed length data types
        Case BT_MISSING
            var = Empty
        Case BT_UINT8
            If indexOfNext + UINT8_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aByte), buffer(indexOfNext), UINT8_LENGTH: indexOfNext = indexOfNext + UINT8_LENGTH
            var = aByte
        Case BT_INT16
            If indexOfNext + INT16_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(anInteger), buffer(indexOfNext), INT16_LENGTH: indexOfNext = indexOfNext + INT16_LENGTH
            var = anInteger
        Case BT_INT32
            If indexOfNext + INT32_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aLong), buffer(indexOfNext), INT32_LENGTH: indexOfNext = indexOfNext + INT32_LENGTH
            var = aLong
        Case BT_INT64
            If indexOfNext + INT64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aCurrency), buffer(indexOfNext), INT64_LENGTH: indexOfNext = indexOfNext + INT64_LENGTH
            var = aCurrency
        Case BT_FLOAT32
            If indexOfNext + FLOAT32_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aSingle), buffer(indexOfNext), FLOAT32_LENGTH: indexOfNext = indexOfNext + FLOAT32_LENGTH
            var = aSingle
        Case BT_FLOAT64
            If indexOfNext + FLOAT64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aDouble), buffer(indexOfNext), FLOAT64_LENGTH: indexOfNext = indexOfNext + FLOAT64_LENGTH
            var = aDouble
        Case BT_BOOLEAN
            If indexOfNext + BOOLEAN_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aByte), buffer(indexOfNext), UINT8_LENGTH: indexOfNext = indexOfNext + BOOLEAN_LENGTH
            If aByte = 0 Then var = False Else var = True
        Case BT_XLDATE64
            If indexOfNext + XLDATE64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory ByVal VarPtr(aDate), buffer(indexOfNext), XLDATE64_LENGTH: indexOfNext = indexOfNext + XLDATE64_LENGTH
            var = aDate
            
        ' variable length data types
        Case BT_UNICODE16
            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
            If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            aString = String(dataLength / 2, " ")
            apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH
            var = aString
            
        ' array of fixed length data types
        Case BT_ARRAY_UINT8, BT_ARRAY_INT16, BT_ARRAY_INT32, BT_ARRAY_INT64, BT_ARRAY_FLOAT32, BT_ARRAY_FLOAT64, BT_ARRAY_BOOLEAN, BT_ARRAY_XLDATE64
            arrayType = VBArrayTypeFromBT(BT)
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                    DBCreateNewArrayOfType anArray, arrayType, 1, 1
                    Erase anArray
                Case 1
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewArrayOfType anArray, arrayType, i1, i2
                    DBGetSafeArrayDetails anArray, SA
                    dataLength = (i2 - i1 + 1) * SA.cbElements
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory ByVal SA.pvData, buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength
                Case 2
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = base
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewArrayOfType anArray, arrayType, i1, i2, j1, j2
                    DBGetSafeArrayDetails anArray, SA
                    dataLength = (i2 - i1 + 1) * (j2 - j1 + 1) * SA.cbElements
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory ByVal SA.pvData, buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength
                Case 3
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    k1 = base
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = base
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewArrayOfType anArray, arrayType, i1, i2, j1, j2, k1, k2
                    DBGetSafeArrayDetails anArray, SA
                    dataLength = (i2 - i1 + 1) * (j2 - j1 + 1) * (k2 - k1 + 1) * SA.cbElements
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory ByVal SA.pvData, buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for " & nameFromBT(BT)
            End Select
            var = anArray

        ' convert into a vbObject containing a Dictionary
        Case BT_DICTIONARY
            ' size is length of each string key as ascii + length of value
            ' raise error if not string keys
            DBErrors_raiseNotYetImplemented ModuleSummary(), METHOD_NAME, "Dictionary"
                
        ' array of strings
        Case BT_ARRAY_UNICODE16
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                    DBCreateNewArrayOfType anArray, vbString, 1, 1
                    Erase anArray
                Case 1
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewArrayOfType anArray, vbString, i1, i2
                    For i = i1 To i2
                        If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                        If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        aString = String(dataLength / 2, " ")
                        apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH
                        anArray(i) = aString
                    Next
                Case 2
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = base
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewArrayOfType anArray, vbString, i1, i2, j1, j2
                    For i = i1 To i2
                        For j = j1 To j2
                            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                            If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            aString = String(dataLength / 2, " ")
                            apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH
                            anArray(i, j) = aString
                        Next
                    Next
                Case 3
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    k1 = base
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = base
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewArrayOfType anArray, vbString, i1, i2, j1, j2, k1, k2
                    For i = i1 To i2
                        For j = j1 To j2
                            For k = k1 To k2
                                If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                                If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                aString = String(dataLength / 2, " ")
                                apiCopyMemory ByVal StrPtr(aString), buffer(indexOfNext), dataLength: indexOfNext = indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH
                                anArray(i, j, k) = aString
                            Next
                        Next
                    Next
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for String()"
            End Select
            var = anArray

        ' array of variants
        Case BT_ARRAY
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                    DBCreateNewVariantArray anArray, 1, 1
                    Erase anArray
                Case 1
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewVariantArray anArray, i1, i2
                    For i = i1 To i2
                         DBBytesAsVariant buffer, indexOfBufferEndPlusOne, indexOfNext, anArray(i), base
                    Next
                Case 2
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = base
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewVariantArray anArray, i1, i2, j1, j2
                    For i = i1 To i2
                        For j = j1 To j2
                            DBBytesAsVariant buffer, indexOfBufferEndPlusOne, indexOfNext, anArray(i, j), base
                        Next
                    Next
                Case 3
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    k1 = base
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = base
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = base
                    i2 = i1 + i2 - 1
                    DBCreateNewVariantArray anArray, i1, i2, j1, j2, k1, k2
                    For i = i1 To i2
                        For j = j1 To j2
                            For k = k1 To k2
                                 DBBytesAsVariant buffer, indexOfBufferEndPlusOne, indexOfNext, anArray(i, j, k), base
                            Next
                        Next
                    Next
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for Variant()"
            End Select
            var = anArray

        Case Else
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise BT " & BT
            
    End Select
    
End Sub


Function DBVerifyStructureOfSerialisedVariant(buffer() As Byte, indexOfBufferEndPlusOne As Long, indexOfNext As Long) As Boolean
    Dim BT As Byte, i As Long, j As Long, k As Long, dataLength As Long, nullCheck As Long
    Dim arrayType As Integer, nDimensions As Long, i1 As Long, i2 As Long, j1 As Long, j2 As Long, k1 As Long, k2 As Long
    
    Const METHOD_NAME As String = "DBVerifyStructureOfSerialisedVariant"
    
    If indexOfNext + TYPE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
    apiCopyMemory BT, buffer(indexOfNext), TYPE_LENGTH: indexOfNext = indexOfNext + TYPE_LENGTH
    
    Select Case BT
    
        ' fixed length data types
        Case BT_MISSING
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_UINT8
            If indexOfNext + UINT8_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + UINT8_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_INT16
            If indexOfNext + INT16_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + INT16_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_INT32
            If indexOfNext + INT32_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + INT32_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_INT64
            If indexOfNext + INT64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + INT64_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_FLOAT32
            If indexOfNext + FLOAT32_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + FLOAT32_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_FLOAT64
            If indexOfNext + FLOAT64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + FLOAT64_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_BOOLEAN
            If indexOfNext + BOOLEAN_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"         ' TRUE -> 255, FALSE -> 0
            indexOfNext = indexOfNext + BOOLEAN_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
        Case BT_XLDATE64
            If indexOfNext + XLDATE64_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + XLDATE64_LENGTH
            DBVerifyStructureOfSerialisedVariant = True
            
        ' variable length data types
        Case BT_UNICODE16
            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
            If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            indexOfNext = indexOfNext + dataLength
            apiCopyMemory nullCheck, buffer(indexOfNext), UNICODE_NULL_TERMINATOR_LENGTH: indexOfNext = indexOfNext + UNICODE_NULL_TERMINATOR_LENGTH
            DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
            
        ' array of fixed length data types
        Case BT_ARRAY_UINT8, BT_ARRAY_INT16, BT_ARRAY_INT32, BT_ARRAY_INT64, BT_ARRAY_FLOAT32, BT_ARRAY_FLOAT64, BT_ARRAY_BOOLEAN, BT_ARRAY_XLDATE64
            arrayType = VBArrayTypeFromBT(BT)
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                Case 1
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    dataLength = (i2 - i1 + 1) * lengthOfVbType(arrayType)
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    indexOfNext = indexOfNext + dataLength
                Case 2
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = 0
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    dataLength = (i2 - i1 + 1) * (j2 - j1 + 1) * lengthOfVbType(arrayType)
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    indexOfNext = indexOfNext + dataLength
                Case 3
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    k1 = 0
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = 0
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    dataLength = (i2 - i1 + 1) * (j2 - j1 + 1) * (k2 - k1 + 1) * lengthOfVbType(arrayType)
                    If indexOfNext + dataLength > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    indexOfNext = indexOfNext + dataLength
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for " & nameFromBT(BT)
            End Select
            DBVerifyStructureOfSerialisedVariant = True

        Case BT_DICTIONARY
            ' size is length of each string key as ascii + length of value
            ' raise error if not string keys
            DBErrors_raiseNotYetImplemented ModuleSummary(), METHOD_NAME, "Dictionary"
                
        ' array of strings
        Case BT_ARRAY_UNICODE16
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                Case 1
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    For i = i1 To i2
                        If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                        If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                        indexOfNext = indexOfNext + dataLength
                        apiCopyMemory nullCheck, buffer(indexOfNext), UNICODE_NULL_TERMINATOR_LENGTH: indexOfNext = indexOfNext + UNICODE_NULL_TERMINATOR_LENGTH
                        DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
                    Next
                Case 2
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = 0
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    For i = i1 To i2
                        For j = j1 To j2
                            If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                            If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                            indexOfNext = indexOfNext + dataLength
                            apiCopyMemory nullCheck, buffer(indexOfNext), UNICODE_NULL_TERMINATOR_LENGTH: indexOfNext = indexOfNext + UNICODE_NULL_TERMINATOR_LENGTH
                            DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
                        Next
                    Next
                Case 3
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    k1 = 0
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = 0
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    For i = i1 To i2
                        For j = j1 To j2
                            For k = k1 To k2
                                If indexOfNext + SIZE_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                apiCopyMemory dataLength, buffer(indexOfNext), SIZE_LENGTH: indexOfNext = indexOfNext + SIZE_LENGTH
                                If indexOfNext + dataLength + UNICODE_NULL_TERMINATOR_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                                indexOfNext = indexOfNext + dataLength
                                apiCopyMemory nullCheck, buffer(indexOfNext), UNICODE_NULL_TERMINATOR_LENGTH: indexOfNext = indexOfNext + UNICODE_NULL_TERMINATOR_LENGTH
                                DBVerifyStructureOfSerialisedVariant = (nullCheck = 0)
                            Next
                        Next
                    Next
                Case Else
                    DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise more than 3 dimensions for String()"
            End Select
            DBVerifyStructureOfSerialisedVariant = True

        ' array of variants
        Case BT_ARRAY
            If indexOfNext + NDIMENSIONS_LENGTH > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
            apiCopyMemory nDimensions, buffer(indexOfNext), NDIMENSIONS_LENGTH: indexOfNext = indexOfNext + NDIMENSIONS_LENGTH
            Select Case nDimensions
                Case 0
                Case 1
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    For i = i1 To i2
                        DBVerifyStructureOfSerialisedVariant = DBVerifyStructureOfSerialisedVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                        If DBVerifyStructureOfSerialisedVariant = False Then Exit Function
                    Next
                Case 2
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = 0
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    For i = i1 To i2
                        For j = j1 To j2
                            DBVerifyStructureOfSerialisedVariant = DBVerifyStructureOfSerialisedVariant(buffer, indexOfBufferEndPlusOne, indexOfNext)
                             If DBVerifyStructureOfSerialisedVariant = False Then Exit Function
                        Next
                    Next
                Case 3
                    If indexOfNext + DIM_LENGTH * nDimensions > indexOfBufferEndPlusOne Then DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Out of buffer"
                    apiCopyMemory k2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    k1 = 0
                    k2 = k1 + k2 - 1
                    apiCopyMemory j2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    j1 = 0
                    j2 = j1 + j2 - 1
                    apiCopyMemory i2, buffer(indexOfNext), DIM_LENGTH: indexOfNext = indexOfNext + DIM_LENGTH
                    i1 = 0
                    i2 = i1 + i2 - 1
                    For i = i1 To i2
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
            DBErrors_raiseGeneralError ModuleSummary(), METHOD_NAME, "Can't deserialise BT " & BT
            
    End Select
    
End Function

Private Function nameFromBT(BT As Byte) As String
    If BT = BT_ARRAY Then nameFromBT = "BT_ARRAY": Exit Function
    Select Case Not (vbArray) And BT
        Case BT_MISSING
            nameFromBT = "BT_MISSING"
        Case BT_UINT8
            nameFromBT = "BT_UINT8"
        Case BT_INT16
            nameFromBT = "BT_INT16"
        Case BT_INT32
            nameFromBT = "BT_INT32"
        Case BT_INT64
            nameFromBT = "BT_INT64"
        Case BT_FLOAT32
            nameFromBT = "BT_FLOAT32"
        Case BT_FLOAT64
            nameFromBT = "BT_FLOAT64"
        Case BT_BOOLEAN
            nameFromBT = "BT_BOOLEAN"
        Case BT_XLDATE64
            nameFromBT = "BT_XLDATE64"
        Case BT_UNICODE16
            nameFromBT = "BT_UNICODE16"
        Case BT_DICTIONARY
            nameFromBT = "BT_DICTIONARY"
        Case Else
    End Select
    If vbArray And BT Then nameFromBT = nameFromBT & "()"
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

