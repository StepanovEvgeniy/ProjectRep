Attribute VB_Name = "mdataTransfer"
Option Explicit
'����������� ��������, ����������� �� �����
Public cPath As String
Public cName As String
Public cPass As String
Public cPass_s As String
'
Public nStartRow As Integer
Public nRowCount_d As Integer
Public nIdPos As Integer
Public Array_ID() As Single
Public Array_Values() As Double
Public tConst() As Variant
Sub auto_open()
    DataTransfer Date - 1, Date
End Sub
Public Sub DataTransfer(pStartDate, pEndDate)

Dim cStartTrans As String
Dim cEndTrans As String
Dim cStartFile As String
Dim cEndFile As String

Dim oWb As Workbook
Dim oWb_d As Workbook
Dim oSh As Worksheet
Dim oSh_d As Worksheet
Dim rConst As Variant
Dim msgR As Variant

Dim nIdPos_d As Integer
Dim nRowCount_s As Integer
Dim nVStartCol As Integer
Dim nVEndCol As Integer
Dim nVStartCol_s1 As Integer
Dim nVStartCol_s2 As Integer
Dim nVEndCol_s1 As Integer
Dim nVEndCol_s2 As Integer
Dim nVStartCol_d1 As Integer
Dim nVStartCol_d2 As Integer
Dim nVEndCol_d1 As Integer
Dim nVEndCol_d2 As Integer

Dim delta_1 As Double
Dim delta_2 As Double
Dim i As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False

rConst = mText.LoadArrayFromTextFile(ThisWorkbook.Path + "\tConst.txt", 1, ";", "$")

For i = LBound(rConst) To UBound(rConst)
    If Trim(rConst(i, 1)) = 1 Then
        cPath = rConst(i, 2)
    End If
    If Trim(rConst(i, 1)) = 2 Then
        cName = rConst(i, 2)
    End If
    If Trim(rConst(i, 1)) = 3 Then
        cPass_s = rConst(i, 2)
    End If
    If Trim(rConst(i, 1)) = 4 Then
        cPass = rConst(i, 2)
    End If
Next i

'����������� �������� �������, ����-�������� �� ���������� ����
nVStartCol_s1 = 23
nVEndCol_s1 = 27
'����������� �������� �������, ����-�������� �� ������� ����
nVStartCol_s2 = 4
nVEndCol_s2 = 22

'����������� �������� �������, ����-�������, ���� ������ �� ���������� ����, ������� � ����������=+20
nVStartCol_d1 = 3
nVEndCol_d1 = 7
'����������� �������� �������, ����-�������, ���� ������ �� ������� ����, ������� � ����������=-4
nVStartCol_d2 = 8
nVEndCol_d2 = 26
'��������� �������
nStartRow = 3
'������� ����������� ������ � ����� ���������
nIdPos = 29
'������� ����������� ������ � ����� ��������
nIdPos_d = 1

'������� ������� ����� ��������� � ���������
delta_1 = nVStartCol_s1 - nVStartCol_d1
delta_2 = nVStartCol_s2 - nVStartCol_d2

Set oWb_d = ActiveWorkbook
Set oSh_d = oWb_d.Sheets("���� �� ������ <�����>")
oSh_d.Activate
ActiveSheet.Protect Password:=cPass, UserInterfaceOnly:=True

nRowCount_d = ActiveCell.SpecialCells(xlLastCell).Row

ReDim Array_ID(nRowCount_d)
ReDim Array_Values(nRowCount_d, 27)

For i = nStartRow To nRowCount_d
    Array_ID(i) = oSh_d.Cells(i, nIdPos_d).Value
Next i

'���������� ����� ������-����������
cStartTrans = Trim(Str(Day(pStartDate))) ' ����� �������� ��������� ������ �� 5  �����  � 0:00 �� 4:00
cEndTrans = Trim(Str(Day(pEndDate)))     ' ����� �������� ��������� ������ �� 19 �����  � 5:00 �� 23:00

cStartFile = "�����_" + cStartTrans + ".xls"
cEndFile = "�����_" + cEndTrans + ".xls"

On Error Resume Next
Set oWb = Workbooks.Open(cPath + cStartFile, False, True, , , , True)
If Err.Number <> 0 Then
    msgR = MsgBox("����������� �������� ������ � ������ '�����'!" + Chr(13) + _
    "��������� ������� ������ �������� ��������� �� ������� � �������� <<������������ ���������>> � ��������� ��������." + Chr(13) + _
    Chr(13) + "�������� ������ �� ���������. ��������� ����� ���������!", vbOKOnly, "������ �������� ������")
    Exit Sub
End If
On Error GoTo 0

oWb.Application.Visible = False

Set oSh = oWb.Worksheets(cName)
oSh.Activate
oSh.Unprotect (cPass_s)
nRowCount_s = ActiveCell.SpecialCells(xlLastCell).Row

'��������� ��������� ��������� ������ ��������� � �������� ������, � ������ ������ � ������
mProcedure.pGetValues nRowCount_s, nVStartCol_d1, nVEndCol_d1, delta_1

oSh.Protect (cPass_s)
oWb.Close

On Error Resume Next
Set oWb = Workbooks.Open(cPath + cEndFile, False, True, , , , True)
If Err.Number <> 0 Then
    msgR = MsgBox("����������� �������� ������ � ������ '�����'!" + Chr(13) + _
    "��������� ������� ������ �������� ��������� �� ������� � �������� <<������������ ���������>> � ��������� ��������." + Chr(13) + _
    Chr(13) + "�������� ������ �� ���������. ��������� ����� ���������!", vbOKOnly, "������ �������� ������")
    Exit Sub
End If
On Error GoTo 0

oWb.Application.Visible = False

Set oSh = oWb.Worksheets(cName)
oSh.Activate
oSh.Unprotect (cPass_s)
nRowCount_s = ActiveCell.SpecialCells(xlLastCell).Row

'��������� ��������� ��������� ������ ��������� � �������� ������, � ������ ������ � ������
mProcedure.pGetValues nRowCount_s, nVStartCol_d2, nVEndCol_d2, delta_2

oSh.Protect (cPass_s)
oWb.Close
oWb_d.Application.Visible = True

oSh_d.Activate

'��������� ��������� ������ ������ � ���� �������
mProcedure.pWriteValues nVStartCol_d1, nVEndCol_d2
Cells(nStartRow, nVStartCol_d1).Select

Application.ScreenUpdating = True
End Sub

