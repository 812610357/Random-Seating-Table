Attribute VB_Name = "ģ��"
Sub ��()
Attribute ��.VB_ProcData.VB_Invoke_Func = " \n14"


'�������
    Dim i, j, Girl, Boy, Last As Long
    Girl = Sheets("��������").Range("F2")
    Boy = Sheets("��������").Range("F3")
    
'���Ů������
    Sheets("Ů��").UsedRange.ClearContents
'��������
    Sheets("��������").Select
    Columns("B:B").Select
    Selection.Copy
    Sheets("Ů��").Select
    Range("A1").Select
    ActiveSheet.Paste
'�������벹λ
    If Sheets("��������").Range("F6") = 1 Then
        Cells(Girl + 2, 1) = "��λ"
        Girl = Girl + 1
    End If
'��������
    For i = 1 To Girl
        Sheets("Ů��").Cells(i + 1, 2) = "=Rand()"
    Next
'�������������������˳��
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Ů��").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Ů��").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Ů��").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
'�����������
    Sheets("����").UsedRange.ClearContents
'��������
    Sheets("��������").Select
    Columns("C:C").Select
    Selection.Copy
    Sheets("����").Select
    Range("A1").Select
    ActiveSheet.Paste
'�������벹λ
    If Sheets("��������").Range("F7") = 1 Then
        Cells(Boy + 2, 1) = "��λ"
        Boy = Boy + 1
    End If
'��������
    For i = 1 To Boy
        Sheets("����").Cells(i + 1, 2) = "=Rand()"
    Next
'�������������������˳��
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("����").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("����").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

        
'�����仺��
    Sheets("���").Rows("1:1").ClearContents
'ת�����Ů��
    Sheets("Ů��").Select
    Range(Cells(2, 1), Cells(Girl + 1, 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-50
    Sheets("���").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'ת���������
    Sheets("����").Select
    Range(Cells(2, 1), Cells(Boy + 1, 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-50
    Sheets("���").Select
    Cells(1, Girl + 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
'���������
    Sheets("����").Columns("B:K").ClearContents
    Sheets("����").Select
    Range("A3").Select
    Selection.AutoFill Destination:=Range("A3:A10"), Type:=xlFillDefault
    Range("A3:A10").Select
'��ȡ��λ��Ϣ
    Sheets("���").Select
    Range("B4:K10").Select
    Selection.Copy
    Sheets("����").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'�����λռλ��
    If Range("K4") = 0 Then
        Range("B4:K10").Select
        Selection.SpecialCells(xlCellTypeConstants, 1).Select
        Selection.ClearContents
    End If
'Ѱ����λ���е����һ��
    For i = 4 To 10
        If Cells(i, 1) <> 0 Then
            Last = i
            Exit For
        End If
    Next
'�������һ�е���λ
    If Cells(Last, 1) = 2 Then
        Range(Cells(Last, 2), Cells(Last, 3)).Select
        Selection.Cut
        Cells(Last, 6).Select
        ActiveSheet.Paste
    Else
        If Cells(Last, 1) = 4 Then
            Range(Cells(Last, 2), Cells(Last, 5)).Select
            Selection.Cut
            Cells(Last, 6).Select
            ActiveSheet.Paste
        Else
            If Cells(Last, 1) = 6 Then
                Range(Cells(Last, 2), Cells(Last, 3)).Select
                Selection.Cut
                Cells(Last, 8).Select
                ActiveSheet.Paste
            End If
        End If
    End If
'����ÿһ�����λ
    For i = 1 To 5
        For j = (i Mod 2 + 1) To Cells(11 + i, 1) / 2 Step 2
            Range(Cells(11 - j, 2 * i), Cells(11 - j, 2 * i + 1)).Select
            Selection.Cut
            Range(Cells(1, 2 * i), Cells(1, 2 * i + 1)).Select
            ActiveSheet.Paste
            Range(Cells(10 - Cells(11 + i, 1) + j, 2 * i), Cells(10 - Cells(11 + i, 1) + j, 2 * i + 1)).Select
            Selection.Cut
            Range(Cells(11 - j, 2 * i), Cells(11 - j, 2 * i + 1)).Select
            ActiveSheet.Paste
            Range(Cells(1, 2 * i), Cells(1, 2 * i + 1)).Select
            Selection.Cut
            Range(Cells(10 - Cells(11 + i, 1) + j, 2 * i), Cells(10 - Cells(11 + i, 1) + j, 2 * i + 1)).Select
            ActiveSheet.Paste
        Next
    Next
'�����λ����
    Sheets("��λ��").Columns("B:O").ClearContents
'�������
    Range("D:D,F:F,H:H,J:J").Select
    Range("J1").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'�����λ��
    Range("B4:O10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("��λ��").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'��λ����
    Sheets("����").Select
    Range("D:D,G:G,J:J,M:M").Select
    Range("M1").Activate
    Selection.Delete Shift:=xlToLeft
'�����λռλ��
    Sheets("��λ��").Select
    Cells.Replace What:="��λ", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
End Sub
