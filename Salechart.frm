VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Salechart 
   Caption         =   "�P����R"
   ClientHeight    =   3885
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9086.001
   OleObjectBlob   =   "Salechart.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Salechart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub But_Orderback_Click()
    Salechart.Hide
    Signin.Show
End Sub

Private Sub But_Ordercheck_Click()
'���R�Ϫ�
    '�ˬd�O�_���ť�
    If Com_Chartyearstart.Value = Empty Then
        MsgBox "�п�J�d�M�϶�!"
        Exit Sub
    End If
    If Com_Chartmonthstart.Value = Empty Then
        MsgBox "�п�J�d�M�϶�!"
        Exit Sub
    End If
    If Com_Chartyearend.Value = Empty Then
        MsgBox "�п�J�d�M�϶�!"
        Exit Sub
    End If
    If Com_Chartmonthend.Value = Empty Then
        MsgBox "�п�J�d�M�϶�!"
        Exit Sub
    End If
     If Com_Chartdayend.Value = Empty Then
        MsgBox "�п�J�d�M�϶�!"
        Exit Sub
    End If
    If Com_Charttype.Value = Empty Then
        MsgBox "�п�J��������!"
        Exit Sub
    End If
     If Com_Chartitem.Value = Empty Then
        MsgBox "�п�J����!"
        Exit Sub
    End If
    
    '���͹Ϫ�
    Dim chartitem As String
    chartitem = Com_Chartitem.Value
    
    Dim sheetindex As Integer
    For sheetindex = 1 To Sheets.Count
        If Sheets(sheetindex).Name = "�Ϫ�" Then
            Application.DisplayAlerts = False
            Sheets("�Ϫ�").Delete
        End If
        Application.DisplayAlerts = True
    Next
    For sheetindex = 1 To Sheets.Count
        If Sheets(sheetindex).Name = "�z��᪺��" Then
            Application.DisplayAlerts = False
            Sheets("�z��᪺��").Delete
        End If
        Application.DisplayAlerts = True
    Next
    
    Dim datestart, dateend As Date
    Dim datediff
    datestart = DateSerial(Com_Chartyearstart.Value, Com_Chartmonthstart.Value, Com_Chartdaystart.Value)
    dateend = DateSerial(Com_Chartyearend.Value, Com_Chartmonthend.Value, Com_Chartdayend.Value)
    datediff = dateend - datestart
    
    '�����z�ﾹ
    Sheets("�P�����").Select
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    
    Select Case Com_Charttype.Text
    
    Case "��B�P����R"
        
        Sheets("�P�����").Select
        Columns("A:F").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.Copy
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�z��᪺��"
        Sheets("�z��᪺��").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:C").Select
        Selection.Delete Shift:=xlToLeft
        Application.CutCopyMode = False
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�Ϫ�"
        Sheets("�Ϫ�").Range("A1").Value = "���"
        Sheets("�Ϫ�").Range("B1").Value = "�\�I�W��"
        Sheets("�Ϫ�").Range("C1").Value = "�P�⦬�q"
        Sheets("�Ϫ�").Range("D1").Value = "�P�⦨��"
        
        Dim rcnt_select, cnt_chart, a, b As Integer
        Sheets("�z��᪺��").Select
        rcnt_select = ActiveSheet.UsedRange.Rows.Count
        Sheets("�Ϫ�").Select
        cnt_chart = ActiveSheet.UsedRange.Rows.Count
        
        '��ɶ��ŦX���K��Ϫ�u�@��̭�
        For a = 2 To rcnt_select
            Sheets("�Ϫ�").Select
            cnt_chart = ActiveSheet.UsedRange.Rows.Count
            If (dateend - Sheets("�z��᪺��").Cells(a, "A").Value < datediff) Or (dateend - Sheets("�z��᪺��").Cells(a, "A").Value = datediff) Then
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "A").Value = Sheets("�z��᪺��").Cells(a, "A").Value
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "B").Value = Sheets("�z��᪺��").Cells(a, "B").Value
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "C").Value = Sheets("�z��᪺��").Cells(a, "C").Value
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "D").Value = Sheets("�z��᪺��").Cells(a, "D").Value
            End If
        Next
        
        '�s�@���q�Ϫ�
        Sheets("�Ϫ�").Select
        Range("B:B,C:C").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("�Ϫ�!$B:$B,�Ϫ�!$C:$C")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        
        '�s�@�����Ϫ�
        Sheets("�Ϫ�").Select
        Range("B:B,D:D").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("�Ϫ�!$B:$B,�Ϫ�!$D:$D")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        
        Application.DisplayAlerts = False
        Sheets("�z��᪺��").Delete
        Application.DisplayAlerts = True
        
        
    Case "��~�P����R"
        
        Sheets("�P�����").Select
        Columns("A:F").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$F$6").AutoFilter Field:=3, Criteria1:=chartitem
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.Copy
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�z��᪺��"
        Sheets("�z��᪺��").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:C").Select
        Selection.Delete Shift:=xlToLeft
        Application.CutCopyMode = False
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�Ϫ�"
        Sheets("�Ϫ�").Range("A1").Value = "���"
        Sheets("�Ϫ�").Range("B1").Value = "�\�I�W��"
        Sheets("�Ϫ�").Range("C1").Value = "�P�⦬�q"
        Sheets("�Ϫ�").Range("D1").Value = "�P�⦨��"
        
        
        Dim rCnt_R, rIdx_R As Integer
        Dim rcnt_chart, rIdx_Chart As Integer
        Dim dtrange_R As Range
        Set dtrange_R = Sheets("�z��᪺��").UsedRange
        rCnt_R = dtrange_R.Rows.Count

        Dim dtrange_Chart As Range
        Set dtrange_Chart = Sheets("�Ϫ�").UsedRange
        rcnt_chart = dtrange_Chart.Rows.Count
        
        '��ɶ��ŦX���K��Ϫ�u�@��̭�
        Sheets("�Ϫ�").Select
        For rIdx_R = 2 To rCnt_R
            Set dtrange_Chart = Sheets("�Ϫ�").UsedRange
            rcnt_chart = dtrange_Chart.Rows.Count
            For rIdx_Chart = 2 To rcnt_chart + 1
                If (dateend - Sheets("�z��᪺��").Cells(rIdx_R, "A").Value < datediff) Or (dateend - Sheets("�z��᪺��").Cells(rIdx_R, "A").Value = datediff) Then
                    Sheets("�Ϫ�").Cells(rIdx_Chart, "A").Value = Sheets("�z��᪺��").Cells(rIdx_R, "A").Value
                    Sheets("�Ϫ�").Cells(rIdx_Chart, "B").Value = Sheets("�z��᪺��").Cells(rIdx_R, "B").Value
                    Sheets("�Ϫ�").Cells(rIdx_Chart, "C").Value = Sheets("�z��᪺��").Cells(rIdx_R, "C").Value
                    Sheets("�Ϫ�").Cells(rIdx_Chart, "D").Value = Sheets("�z��᪺��").Cells(rIdx_R, "D").Value
                End If
            Next
        Next
        
        Application.DisplayAlerts = False
        Sheets("�z��᪺��").Delete
        Application.DisplayAlerts = True
        
        '��P���B�򦨥��ۥ[
        Sheets("�Ϫ�").Activate
        Dim rcnt As Integer
        Dim i As Integer
        rcnt = ActiveSheet.UsedRange.Rows.Count
        
        Sheets("�Ϫ�").Cells(1, "F").Value = "��~�`�P�⦬�q"
        Sheets("�Ϫ�").Cells(1, "G").Value = "��~�`�P�⦨��"
        
        For i = 2 To rcnt
        '�`�P���B
            Sheets("�Ϫ�").Cells(2, "F").Value = Sheets("�Ϫ�").Cells(2, "F").Value + Sheets("�Ϫ�").Cells(rcnt, "C").Value
            Sheets("�Ϫ�").Cells(2, "G").Value = Sheets("�Ϫ�").Cells(2, "G").Value + Sheets("�Ϫ�").Cells(rcnt, "D").Value
        Next
        
        '�s�@��~�P����R
        Sheets("�Ϫ�").Activate
        Range("F:F,G:G").Select
        
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("�Ϫ�!$F:$F,�Ϫ�!$G:$G")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = chartitem
        
    Case "���O�P����R"
    
        Sheets("�P�����").Select
        Columns("A:G").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$G$6").AutoFilter Field:=7, Criteria1:=chartitem
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.Copy
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�z��᪺��"
        Sheets("�z��᪺��").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:C").Select
        Selection.Delete Shift:=xlToLeft
        Application.CutCopyMode = False
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "�Ϫ�"
        Sheets("�Ϫ�").Range("A1").Value = "���"
        Sheets("�Ϫ�").Range("B1").Value = "�\�I�W��"
        Sheets("�Ϫ�").Range("C1").Value = "�P�⦬�q"
        Sheets("�Ϫ�").Range("D1").Value = "�P�⦨��"
        
        
       
        
        Sheets("�z��᪺��").Select
        rcnt_select = ActiveSheet.UsedRange.Rows.Count
        Sheets("�Ϫ�").Select
        cnt_chart = ActiveSheet.UsedRange.Rows.Count
        
        '��ɶ��ŦX���K��Ϫ�u�@��̭�
        For a = 2 To rcnt_select
            Sheets("�Ϫ�").Select
            cnt_chart = ActiveSheet.UsedRange.Rows.Count
            If (dateend - Sheets("�z��᪺��").Cells(a, "A").Value < datediff) Or (dateend - Sheets("�z��᪺��").Cells(a, "A").Value = datediff) Then
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "A").Value = Sheets("�z��᪺��").Cells(a, "A").Value
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "B").Value = Sheets("�z��᪺��").Cells(a, "B").Value
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "C").Value = Sheets("�z��᪺��").Cells(a, "C").Value
                Sheets("�Ϫ�").Cells(cnt_chart + 1, "D").Value = Sheets("�z��᪺��").Cells(a, "D").Value
            End If
        Next
    
    
    '�s�@���q�Ϫ�
        Sheets("�Ϫ�").Select
        Range("B:B,C:C").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("�Ϫ�!$B:$B,�Ϫ�!$C:$C")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = chartitem
        
        '�s�@�����Ϫ�
        Sheets("�Ϫ�").Select
        Range("B:B,D:D").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("�Ϫ�!$B:$B,�Ϫ�!$D:$D")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = chartitem
        
        Application.DisplayAlerts = False
        Sheets("�z��᪺��").Delete
        Application.DisplayAlerts = True
    
    
    End Select
End Sub

Private Sub But_Orderclear_Click()
    Com_Charttype.Value = Null
    Com_Chartyearend.Value = Null
    Com_Chartmonthend.Value = Null
    Com_Chartdayend.Value = Null
    Com_Chartyearstart.Value = Null
    Com_Chartmonthstart.Value = Null
    Com_Chartdaystart.Value = Null
    Com_Chartitem.Value = Null
End Sub

Private Sub UserForm_Initialize()
    '�d�ߵ������]����
    Com_Chartyearend.Value = year(Date)
    Com_Chartmonthend.Value = Month(Date)
    Com_Chartdayend.Value = Day(Date)
    
   '�d�߰϶��}�l���
    Dim yearstart As Integer
    Dim chart_year As Integer
    Dim chart_month As Integer
    Dim chart_day As Integer
    
    yearstart = year(Date)
    For chart_year = 2019 To yearstart
        Com_Chartyearstart.AddItem chart_year
        Com_Chartyearend.AddItem chart_year
    Next
    
    For chart_month = 1 To 12
        Com_Chartmonthstart.AddItem chart_month
        Com_Chartmonthend.AddItem chart_month
    Next
    
    For chart_day = 1 To 31
        Com_Chartdaystart.AddItem chart_day
        Com_Chartdayend.AddItem chart_day
    Next

    
    '���������U�Ԧ����s�W
    Com_Charttype.AddItem "��B�P����R"
    Com_Charttype.AddItem "��~�P����R"
    Com_Charttype.AddItem "���O�P����R"
    
End Sub

Private Sub Com_Charttype_Change()

    Select Case Com_Charttype.Text
        Case Is = "��~�P����R"
            Com_Chartitem.Clear
            Sheets("���޲z").Select
            Dim menurow As Integer
            Dim cnt As Integer
            menurow = ActiveSheet.UsedRange.Rows.Count
            For cnt = 2 To menurow
                Com_Chartitem.AddItem Cells(cnt, "C").Value
            Next
            
        Case Is = "���O�P����R"
            Com_Chartitem.Clear
            Com_Chartitem.AddItem "�ѭ�"
            Com_Chartitem.AddItem "�I��"
            Com_Chartitem.AddItem "����"
            
        Case Is = "��B�P����R"
            Com_Chartitem.Clear
            Com_Chartitem.AddItem "�`��"
    End Select
    

End Sub

