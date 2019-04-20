Attribute VB_Name = "mFormActions"
Public Sub save_()
'Save
Dim tArr(1 To 2) As Variant
Dim cTxt As String

cTxt = ""
fWrite = False

If fConst.TextBox1.Value <> "" And fConst.TextBox2.Value <> "" Then
    cPath = fConst.TextBox1.Value
    cName = fConst.TextBox2.Value
    tConst(1, 2) = cPath
    tConst(2, 2) = cName
    For i = LBound(tConst) To UBound(tConst)
        For j = LBound(tArr) To UBound(tArr)
            tArr(j) = tConst(i, j)
        Next j
        cTxt = cTxt + Join(tArr, ";") + "$"
    Next i
    
    fWrite = mText.SaveTXTfile(ThisWorkbook.Path + "\tConst.txt", cTxt)
Else
    msgCrash = MsgBox("Запись не произведена.Введите данные!", vbOKOnly, "Настройка констант")
End If

If fWrite = True Then
    msgHit = MsgBox("Произведена запись в файл настройки!", vbOKOnly, "Настройка констант")
End If
exit_

End Sub
Public Sub exit_()
'Exit
 Unload fConst
End Sub
Public Sub input_()
'Input
fConst.TextBox1.Value = "\\altayoic\D\Zadachi\REJODU\"
fConst.TextBox2.Value = "Режим_ЦДП"
save_

End Sub
Public Sub form_activate()

tConst = mText.LoadArrayFromTextFile(ThisWorkbook.Path + "\tConst.txt", 1, ";", "$")

For h = LBound(tConst) To UBound(tConst)
    If Trim(tConst(h, 1)) = 1 Then
        cPath = tConst(h, 2)
    End If
    If Trim(tConst(h, 1)) = 2 Then
        cName = tConst(h, 2)
    End If
    If Trim(tConst(h, 1)) = 3 Then
        cPass_s = tConst(h, 2)
    End If
    If Trim(tConst(h, 1)) = 4 Then
        cPass = tConst(h, 2)
    End If
Next h

If cPath = "" Then
    cPath = "\\altayoic\D\Zadachi\REJODU\"
End If
If cName = "" Then
    cName = "Режим_ЦДП"
End If
If cPass_s = "" Then
   cPass_s = вампир
End If
If cPass = "" Then
    cPass = Smax
End If

fConst.TextBox1.Value = cPath
fConst.TextBox2.Value = cName

End Sub

