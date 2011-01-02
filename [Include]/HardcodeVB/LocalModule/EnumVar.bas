Attribute VB_Name = "MEnumVariant"
Option Explicit

''' Flag must be in standard module so that there is only one copy of it
Public fNotFirstTime As Boolean

'' These functions will be placed in the v-table and executed as the
'' real methods of the IEnumVARIANT object. They must be in a standard
'' module because there must be only one copy of them, and AddressOf
'' only works on standard module procedures.

' Replace IEnumVARIANT_Next
Public Function BasNext(ByVal this As IVBEnumVARIANT, ByVal cv As Long, _
                        av As Variant, ByVal pcvFetched As Long) As Long
    ' this - Object pointer
    ' cv - Count of variants requested for return
    ' av - Array to hold the requested variants
    ' pcvFetched - Pointer to number of variants actually returned
                        
    Dim vTmp As Variant, vEmpty As Variant
    Dim pv As Long, cvFetched As Long, fFetched As Boolean
    Dim i As Integer, vars As CEnumVariant
    ' First hidden argument of an object method is the object pointer--known as
    ' the this pointer in C++. Set this to be an object of our internal
    ' enumeration class.
    Set vars = this
    On Error Resume Next
    ' Get the address of the first variant in array
    pv = VarPtr(av)
    ' Iterate through each requested variant
    For i = 1 To cv
        ' Call the class method that raises a Next event--it returns
        ' true if the next value is fetched
        fFetched = vars.ClsNext(vTmp)
        ' If failure or nothing fetched, we're done
        If (Err) Or fFetched = False Then Exit For
        ' Copy variant to current array position
        CopyMemory ByVal pv, vTmp, 16
        ' Empty work variant without destroying its object or string
        CopyMemory vTmp, vEmpty, 16
        ' Count the variant and point to the next one
        cvFetched = cvFetched + 1
        pv = pv + 16
    Next
    ' If error caused termination, undo what we did
    If Err.Number Then
        ' Iterate back, emptying the invalid fetched variants
        For i = i To 1 Step -1
            ' Copy variant to current array position
            CopyMemory vTmp, ByVal pv, 16
            ' Empty work variant, destroying any object or string
            vTmp = Empty
            ' Empty array variant without destroying any object or string
            CopyMemory ByVal pv, vEmpty, 16
            ' Point to previous array element
            pv = pv - 16
        Next
        ' Convert error to COM format
        BasNext = MapErr(Err)
        ' Return 0 as the number fetched after error
        If pcvFetched Then CopyMemory ByVal pcvFetched, ByVal 0&, 4
    Else
        ' If nothing fetched, break out of enumeration
        If cvFetched = 0 Then BasNext = 1
        ' Copy the actual number fetched to the pointer to fetched count
        If pcvFetched Then CopyMemory ByVal pcvFetched, cvFetched, 4
    End If
End Function

' Replace IEnumVARIANT_Skip
Public Function BasSkip(ByVal this As IVBEnumVARIANT, _
                        ByVal cv As Long) As Long
    Dim vars As CEnumVariant, i As Long
    Set vars = this
    On Error Resume Next
    ' Call the class method that raises a Skip event
    vars.ClsSkip cv
    BasSkip = MapErr(Err)
End Function

' Put the function address (callback) directly into the object v-table
Public Function ReplaceVtableEntry(ByVal pObj As Long, _
                                   ByVal iEntry As Integer, _
                                   ByVal pFunc As Long) As Long
    ' pObj - Pointer to object whose v-table will be modified
    ' iEntry - Index of v-table entry to be modified
    ' pFunc - Function pointer of new v-table method
                            
    Dim pFuncOld As Long, pVTableHead As Long
    Dim pFuncTmp As Long, lOldProtect As Long
    
    ' Object pointer contains a pointer to v-table--copy it to temporary
    CopyMemory pVTableHead, ByVal pObj, 4       ' pVTableHead = *pObj;
    ' Calculate pointer to specified entry
    pFuncTmp = pVTableHead + (iEntry - 1) * 4
    ' Save address of previous method for return
    CopyMemory pFuncOld, ByVal pFuncTmp, 4      ' pFuncOld = *pFuncTmp;
    ' Ignore if they're already the same
    If pFuncOld <> pFunc Then
        ' Need to change page protection to write to code
        VirtualProtect pFuncTmp, 4, PAGE_EXECUTE_READWRITE, lOldProtect
        ' Write the new function address into the v-table
        CopyMemory ByVal pFuncTmp, pFunc, 4     ' *pFuncTmp = pfunc;
        ' Restore the previous page protection
        VirtualProtect pFuncTmp, 4, lOldProtect, lOldProtect 'Optional
    End If
    ReplaceVtableEntry = pFuncOld
End Function

Public Function MapErr(ByVal ErrNumber As Long) As Long
    If ErrNumber Then
        If (ErrNumber And &H80000000) Or (ErrNumber = 1) Then
            'Error HRESULT already set
            MapErr = ErrNumber
        Else
            'Map back to a basic error number
            MapErr = &H800A0000 Or ErrNumber
        End If
    End If
End Function



