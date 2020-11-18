Attribute VB_Name = "模块"
Sub 宏()
Attribute 宏.VB_ProcData.VB_Invoke_Func = " \n14"


'定义变量
    Dim i, j, Girl, Boy, Last As Long
    Girl = Sheets("导入名单").Range("F2")
    Boy = Sheets("导入名单").Range("F3")
    
'清空女生缓存
    Sheets("女生").UsedRange.ClearContents
'复制名单
    Sheets("导入名单").Select
    Columns("B:B").Select
    Selection.Copy
    Sheets("女生").Select
    Range("A1").Select
    ActiveSheet.Paste
'单桌加入补位
    If Sheets("导入名单").Range("F6") = 1 Then
        Cells(Girl + 2, 1) = "补位"
        Girl = Girl + 1
    End If
'填充随机数
    For i = 1 To Girl
        Sheets("女生").Cells(i + 1, 2) = "=Rand()"
    Next
'排序随机数（打乱排序顺序）
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("女生").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("女生").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("女生").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
'清空男生缓存
    Sheets("男生").UsedRange.ClearContents
'复制名单
    Sheets("导入名单").Select
    Columns("C:C").Select
    Selection.Copy
    Sheets("男生").Select
    Range("A1").Select
    ActiveSheet.Paste
'单桌加入补位
    If Sheets("导入名单").Range("F7") = 1 Then
        Cells(Boy + 2, 1) = "补位"
        Boy = Boy + 1
    End If
'填充随机数
    For i = 1 To Boy
        Sheets("男生").Cells(i + 1, 2) = "=Rand()"
    Next
'排序随机数（打乱排序顺序）
    Range("B1").Activate
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("男生").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("男生").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("男生").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

        
'清空填充缓存
    Sheets("填充").Rows("1:1").ClearContents
'转置填充女生
    Sheets("女生").Select
    Range(Cells(2, 1), Cells(Girl + 1, 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-50
    Sheets("填充").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
'转置填充男生
    Sheets("男生").Select
    Range(Cells(2, 1), Cells(Boy + 1, 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-50
    Sheets("填充").Select
    Cells(1, Girl + 1).Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
'清空整理缓存
    Sheets("整理").Columns("B:K").ClearContents
    Sheets("整理").Select
    Range("A3").Select
    Selection.AutoFill Destination:=Range("A3:A10"), Type:=xlFillDefault
    Range("A3:A10").Select
'获取座位信息
    Sheets("填充").Select
    Range("B4:K10").Select
    Selection.Copy
    Sheets("整理").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'清除空位占位符
    If Range("K4") = 0 Then
        Range("B4:K10").Select
        Selection.SpecialCells(xlCellTypeConstants, 1).Select
        Selection.ClearContents
    End If
'寻找座位表中的最后一行
    For i = 4 To 10
        If Cells(i, 1) <> 0 Then
            Last = i
            Exit For
        End If
    Next
'调整最后一行的座位
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
'调整每一组的座位
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
'清空座位表缓存
    Sheets("座位表").Columns("B:O").ClearContents
'拉开组距
    Range("D:D,F:F,H:H,J:J").Select
    Range("J1").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'输出座位表
    Range("B4:O10").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("座位表").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'复位整理
    Sheets("整理").Select
    Range("D:D,G:G,J:J,M:M").Select
    Range("M1").Activate
    Selection.Delete Shift:=xlToLeft
'清除补位占位符
    Sheets("座位表").Select
    Cells.Replace What:="补位", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
End Sub
