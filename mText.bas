Attribute VB_Name = "mText"
'Запись в текстовый файл из переменной:
Function SaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
On Error Resume Next
If Err.Number = 0 Then
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    SaveTXTfile = Err = 0
    Set ts = Nothing: Set FSO = Nothing
Else
    msgR = MsgBox("Запись невозможна! Отсутствуют полномочия!", vbOKOnly, "Настройка констант")
End If
End Function
Function LoadArrayFromTextFile(ByVal filename$, Optional ByVal FirstRow& = 1, _
                               Optional ByVal ColumnsSeparator$ = ";", Optional ByVal RowsSeparator$ = vbNewLine) As Variant
   ' Функция открывает текстовый файл filename$,
   ' и загружает макссив данных, начиная со строки FirstRow&
   ' В качестве параметров можно задать  разделители строк и столбцов для разбиваемой строки
   ' Возвращает двумерный массив - результат преобразования текстового файла в двумерный массив

    On Error Resume Next
    Set FSO = CreateObject("scripting.filesystemobject")        ' читаем текст из выбранного файла
    Set ts = FSO.OpenTextFile(filename$, 1, True): txt$ = ts.ReadAll: ts.Close
    Set ts = Nothing: Set FSO = Nothing

    txt = Trim(txt): Err.Clear        ' разделяем текст на строки и столбцы
    If txt Like "*" & RowsSeparator$ Then txt = Left(txt, Len(txt) - Len(RowsSeparator$))

    If FirstRow& > 1 Then        ' обрезаем ненужные строки
       txt = Split(txt, RowsSeparator$, FirstRow&)(FirstRow& - 1)
    End If

    Err.Clear: tmpArr1 = Split(txt, RowsSeparator$): RowsCount = UBound(tmpArr1) + 1
    ColumnsCount = UBound(Split(tmpArr1(0), ColumnsSeparator$)) + 1
    
    If Err.Number > 0 Then MsgBox "Текст файла " & Dir(filename$, vbNormal) & _
     " не может быть считан в двумерный массив", vbCritical: Exit Function
    ReDim arr(1 To RowsCount, 1 To ColumnsCount)

    For i = LBound(tmpArr1) To UBound(tmpArr1)
        tmpArr2 = Split(Trim(tmpArr1(i)), ColumnsSeparator$)
        For j = 1 To UBound(tmpArr2) + 1
            arr(i + 1, j) = tmpArr2(j - 1)
        Next j
    Next i
  
    LoadArrayFromTextFile = arr

End Function
