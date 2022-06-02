VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Salechart 
   Caption         =   "銷售分析"
   ClientHeight    =   3885
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   9086.001
   OleObjectBlob   =   "Salechart.frx":0000
   StartUpPosition =   1  '所屬視窗中央
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
'分析圖表
    '檢查是否有空白
    If Com_Chartyearstart.Value = Empty Then
        MsgBox "請輸入查尋區間!"
        Exit Sub
    End If
    If Com_Chartmonthstart.Value = Empty Then
        MsgBox "請輸入查尋區間!"
        Exit Sub
    End If
    If Com_Chartyearend.Value = Empty Then
        MsgBox "請輸入查尋區間!"
        Exit Sub
    End If
    If Com_Chartmonthend.Value = Empty Then
        MsgBox "請輸入查尋區間!"
        Exit Sub
    End If
     If Com_Chartdayend.Value = Empty Then
        MsgBox "請輸入查尋區間!"
        Exit Sub
    End If
    If Com_Charttype.Value = Empty Then
        MsgBox "請輸入報表類型!"
        Exit Sub
    End If
     If Com_Chartitem.Value = Empty Then
        MsgBox "請輸入項目!"
        Exit Sub
    End If
    
    '產生圖表
    Dim chartitem As String
    chartitem = Com_Chartitem.Value
    
    Dim sheetindex As Integer
    For sheetindex = 1 To Sheets.Count
        If Sheets(sheetindex).Name = "圖表" Then
            Application.DisplayAlerts = False
            Sheets("圖表").Delete
        End If
        Application.DisplayAlerts = True
    Next
    For sheetindex = 1 To Sheets.Count
        If Sheets(sheetindex).Name = "篩選後的值" Then
            Application.DisplayAlerts = False
            Sheets("篩選後的值").Delete
        End If
        Application.DisplayAlerts = True
    Next
    
    Dim datestart, dateend As Date
    Dim datediff
    datestart = DateSerial(Com_Chartyearstart.Value, Com_Chartmonthstart.Value, Com_Chartdaystart.Value)
    dateend = DateSerial(Com_Chartyearend.Value, Com_Chartmonthend.Value, Com_Chartdayend.Value)
    datediff = dateend - datestart
    
    '關掉篩選器
    Sheets("銷售紀錄").Select
    If ActiveSheet.AutoFilterMode = True Then
       ActiveSheet.AutoFilterMode = False
    End If

    
    Select Case Com_Charttype.Text
    
    Case "營運銷售分析"
        
        Sheets("銷售紀錄").Select
        Columns("A:F").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.Copy
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "篩選後的值"
        Sheets("篩選後的值").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:C").Select
        Selection.Delete Shift:=xlToLeft
        Application.CutCopyMode = False
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "圖表"
        Sheets("圖表").Range("A1").Value = "日期"
        Sheets("圖表").Range("B1").Value = "餐點名稱"
        Sheets("圖表").Range("C1").Value = "銷售收益"
        Sheets("圖表").Range("D1").Value = "銷售成本"
        
        Dim rcnt_select, cnt_chart, a, b As Integer
        Sheets("篩選後的值").Select
        rcnt_select = ActiveSheet.UsedRange.Rows.Count
        Sheets("圖表").Select
        cnt_chart = ActiveSheet.UsedRange.Rows.Count
        
        '把時間符合的貼到圖表工作表裡面
        For a = 2 To rcnt_select
            Sheets("圖表").Select
            cnt_chart = ActiveSheet.UsedRange.Rows.Count
            If (dateend - Sheets("篩選後的值").Cells(a, "A").Value < datediff) Or (dateend - Sheets("篩選後的值").Cells(a, "A").Value = datediff) Then
                Sheets("圖表").Cells(cnt_chart + 1, "A").Value = Sheets("篩選後的值").Cells(a, "A").Value
                Sheets("圖表").Cells(cnt_chart + 1, "B").Value = Sheets("篩選後的值").Cells(a, "B").Value
                Sheets("圖表").Cells(cnt_chart + 1, "C").Value = Sheets("篩選後的值").Cells(a, "C").Value
                Sheets("圖表").Cells(cnt_chart + 1, "D").Value = Sheets("篩選後的值").Cells(a, "D").Value
            End If
        Next
        
        '製作收益圖表
        Sheets("圖表").Select
        Range("B:B,C:C").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("圖表!$B:$B,圖表!$C:$C")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        
        '製作成本圖表
        Sheets("圖表").Select
        Range("B:B,D:D").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("圖表!$B:$B,圖表!$D:$D")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        
        Application.DisplayAlerts = False
        Sheets("篩選後的值").Delete
        Application.DisplayAlerts = True
        
        
    Case "單品銷售分析"
        
        Sheets("銷售紀錄").Select
        Columns("A:F").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$F$6").AutoFilter Field:=3, Criteria1:=chartitem
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.Copy
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "篩選後的值"
        Sheets("篩選後的值").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:C").Select
        Selection.Delete Shift:=xlToLeft
        Application.CutCopyMode = False
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "圖表"
        Sheets("圖表").Range("A1").Value = "日期"
        Sheets("圖表").Range("B1").Value = "餐點名稱"
        Sheets("圖表").Range("C1").Value = "銷售收益"
        Sheets("圖表").Range("D1").Value = "銷售成本"
        
        
        Dim rCnt_R, rIdx_R As Integer
        Dim rcnt_chart, rIdx_Chart As Integer
        Dim dtrange_R As Range
        Set dtrange_R = Sheets("篩選後的值").UsedRange
        rCnt_R = dtrange_R.Rows.Count

        Dim dtrange_Chart As Range
        Set dtrange_Chart = Sheets("圖表").UsedRange
        rcnt_chart = dtrange_Chart.Rows.Count
        
        '把時間符合的貼到圖表工作表裡面
        Sheets("圖表").Select
        For rIdx_R = 2 To rCnt_R
            Set dtrange_Chart = Sheets("圖表").UsedRange
            rcnt_chart = dtrange_Chart.Rows.Count
            For rIdx_Chart = 2 To rcnt_chart + 1
                If (dateend - Sheets("篩選後的值").Cells(rIdx_R, "A").Value < datediff) Or (dateend - Sheets("篩選後的值").Cells(rIdx_R, "A").Value = datediff) Then
                    Sheets("圖表").Cells(rIdx_Chart, "A").Value = Sheets("篩選後的值").Cells(rIdx_R, "A").Value
                    Sheets("圖表").Cells(rIdx_Chart, "B").Value = Sheets("篩選後的值").Cells(rIdx_R, "B").Value
                    Sheets("圖表").Cells(rIdx_Chart, "C").Value = Sheets("篩選後的值").Cells(rIdx_R, "C").Value
                    Sheets("圖表").Cells(rIdx_Chart, "D").Value = Sheets("篩選後的值").Cells(rIdx_R, "D").Value
                End If
            Next
        Next
        
        Application.DisplayAlerts = False
        Sheets("篩選後的值").Delete
        Application.DisplayAlerts = True
        
        '把銷售額跟成本相加
        Sheets("圖表").Activate
        Dim rcnt As Integer
        Dim i As Integer
        rcnt = ActiveSheet.UsedRange.Rows.Count
        
        Sheets("圖表").Cells(1, "F").Value = "單品總銷售收益"
        Sheets("圖表").Cells(1, "G").Value = "單品總銷售成本"
        
        For i = 2 To rcnt
        '總銷售額
            Sheets("圖表").Cells(2, "F").Value = Sheets("圖表").Cells(2, "F").Value + Sheets("圖表").Cells(rcnt, "C").Value
            Sheets("圖表").Cells(2, "G").Value = Sheets("圖表").Cells(2, "G").Value + Sheets("圖表").Cells(rcnt, "D").Value
        Next
        
        '製作單品銷售分析
        Sheets("圖表").Activate
        Range("F:F,G:G").Select
        
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("圖表!$F:$F,圖表!$G:$G")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = chartitem
        
    Case "類別銷售分析"
    
        Sheets("銷售紀錄").Select
        Columns("A:G").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$G$6").AutoFilter Field:=7, Criteria1:=chartitem
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.Copy
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "篩選後的值"
        Sheets("篩選後的值").Select
        ActiveSheet.Paste
        Columns("A:A").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:C").Select
        Selection.Delete Shift:=xlToLeft
        Application.CutCopyMode = False
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        Worksheets(ActiveWorkbook.Worksheets.Count).Name = "圖表"
        Sheets("圖表").Range("A1").Value = "日期"
        Sheets("圖表").Range("B1").Value = "餐點名稱"
        Sheets("圖表").Range("C1").Value = "銷售收益"
        Sheets("圖表").Range("D1").Value = "銷售成本"
        
        
       
        
        Sheets("篩選後的值").Select
        rcnt_select = ActiveSheet.UsedRange.Rows.Count
        Sheets("圖表").Select
        cnt_chart = ActiveSheet.UsedRange.Rows.Count
        
        '把時間符合的貼到圖表工作表裡面
        For a = 2 To rcnt_select
            Sheets("圖表").Select
            cnt_chart = ActiveSheet.UsedRange.Rows.Count
            If (dateend - Sheets("篩選後的值").Cells(a, "A").Value < datediff) Or (dateend - Sheets("篩選後的值").Cells(a, "A").Value = datediff) Then
                Sheets("圖表").Cells(cnt_chart + 1, "A").Value = Sheets("篩選後的值").Cells(a, "A").Value
                Sheets("圖表").Cells(cnt_chart + 1, "B").Value = Sheets("篩選後的值").Cells(a, "B").Value
                Sheets("圖表").Cells(cnt_chart + 1, "C").Value = Sheets("篩選後的值").Cells(a, "C").Value
                Sheets("圖表").Cells(cnt_chart + 1, "D").Value = Sheets("篩選後的值").Cells(a, "D").Value
            End If
        Next
    
    
    '製作收益圖表
        Sheets("圖表").Select
        Range("B:B,C:C").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("圖表!$B:$B,圖表!$C:$C")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = chartitem
        
        '製作成本圖表
        Sheets("圖表").Select
        Range("B:B,D:D").Select
        Range("D1").Activate
        ActiveSheet.Shapes.AddChart2(251, xlPie).Select
        ActiveChart.SetSourceData Source:=Range("圖表!$B:$B,圖表!$D:$D")
        ActiveChart.SetElement (msoElementDataLabelBestFit)
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = chartitem
        
        Application.DisplayAlerts = False
        Sheets("篩選後的值").Delete
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
    '查詢結束日初設今天
    Com_Chartyearend.Value = year(Date)
    Com_Chartmonthend.Value = Month(Date)
    Com_Chartdayend.Value = Day(Date)
    
   '查詢區間開始日期
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

    
    '報表類型下拉式選單新增
    Com_Charttype.AddItem "營運銷售分析"
    Com_Charttype.AddItem "單品銷售分析"
    Com_Charttype.AddItem "類別銷售分析"
    
End Sub

Private Sub Com_Charttype_Change()

    Select Case Com_Charttype.Text
        Case Is = "單品銷售分析"
            Com_Chartitem.Clear
            Sheets("菜單管理").Select
            Dim menurow As Integer
            Dim cnt As Integer
            menurow = ActiveSheet.UsedRange.Rows.Count
            For cnt = 2 To menurow
                Com_Chartitem.AddItem Cells(cnt, "C").Value
            Next
            
        Case Is = "類別銷售分析"
            Com_Chartitem.Clear
            Com_Chartitem.AddItem "麵食"
            Com_Chartitem.AddItem "點心"
            Com_Chartitem.AddItem "飲料"
            
        Case Is = "營運銷售分析"
            Com_Chartitem.Clear
            Com_Chartitem.AddItem "總體"
    End Select
    

End Sub

