Attribute VB_Name = "mProcedure"
Public Sub pGetValues(pRowCount_s, pVStartCol, pVEndCol, pDelta)

For j = nStartRow To nRowCount_d
    cSearched = Trim(Array_ID(j))
    For k = nStartRow To pRowCount_s
        If Cells(k, nIdPos).Value <> "" Then
            cFound = Trim(Cells(k, nIdPos).Value)
            If cFound = cSearched Then
                For l = pVStartCol To pVEndCol
                    Array_Values(j, l) = Cells(k, l + pDelta).Value
                Next l
                Exit For
            End If
        End If
    Next k
Next j

End Sub
Public Sub pWriteValues(pVStartCol, pVEndCol)

For i = nStartRow To nRowCount_d
    For j = pVStartCol To pVEndCol
        Cells(i, j).Value = Array_Values(i, j)
    Next j
Next i

End Sub
Public Sub commandExecute()
    mdataTransfer.DataTransfer Date - 1, Date
End Sub
