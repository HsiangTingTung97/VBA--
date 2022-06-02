VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Order 
   Caption         =   "�I�\�t��"
   ClientHeight    =   6209
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8547.001
   OleObjectBlob   =   "Order.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub But_Orderback_Click()
    Order.Hide
    Signin.Show
End Sub

Private Sub But_Ordercheck_Click()
    '�e�X�q��
    Dim recordcnt As Integer
    Dim sheetrow As Integer
    Dim cnt As Integer
    Dim textlen As Integer
    Dim ordercnt As Integer
    Dim ordername As String
    Dim list As String
    Dim row As Integer
    Dim i, x As Integer
    Dim menuprice, orderamount As Integer

    recordcnt = List_Orderrecord.ListCount
    
    For cnt = 1 To recordcnt
        Sheets("�P�����").Select
        sheetrow = ActiveSheet.UsedRange.Rows.Count
        list = List_Orderrecord.list(cnt - 1)
        textlen = Len(list)
        ordername = Left(list, textlen - 2)
        ordercnt = CByte(Left(Right(list, 2), 1))
        Cells(sheetrow + 1, "A").Value = Signin.Txt_Username.Text
        Cells(sheetrow + 1, "B").Value = Date
        Cells(sheetrow + 1, "C").Value = ordername
        Cells(sheetrow + 1, "D").Value = ordercnt
        Sheets("�P�����").Select
        sheetrow = ActiveSheet.UsedRange.Rows.Count
        Sheets("���޲z").Select
        row = ActiveSheet.UsedRange.Rows.Count
        For i = 1 To sheetrow
            For x = 1 To row
                If Worksheets("�P�����").Cells(i, "C").Value = Worksheets("���޲z").Cells(x, "C").Value Then
                    menuprice = Worksheets("���޲z").Cells(x, "E").Value
                    orderamount = Worksheets("�P�����").Cells(i, "D").Value
                    Worksheets("�P�����").Cells(i, "E").Value = menuprice * orderamount
                    Worksheets("�P�����").Cells(i, "F").Value = Worksheets("���޲z").Cells(x, "F").Value * orderamount
                End If
            Next
        Next
    Next
    MsgBox ("�w�e�X�q��")
    
    '�N�\�I�k�Ǧ����O
    Dim menu_cnt, sale_cnt As Integer
    Sheets("���޲z").Select
    menu_cnt = ActiveSheet.UsedRange.Rows.Count
    Sheets("�P�����").Select
    sale_cnt = ActiveSheet.UsedRange.Rows.Count
    
    Dim menu_typecnt, sale_typecnt As Integer
    
    For menu_typecnt = 2 To menu_cnt
    
        For sale_typecnt = 2 To sale_cnt
            If (Sheets("�P�����").Cells(sale_typecnt, "C").Value = Sheets("���޲z").Cells(sale_typecnt, "C").Value) Then
                Sheets("�P�����").Cells(sale_typecnt, "G").Value = Sheets("���޲z").Cells(sale_typecnt, "D").Value
            End If
        Next
    Next
    
End Sub

Private Sub But_Orderclear_Click()
    List_Orderrecord.Clear
    Com_Orderamount.Value = Null
End Sub

Private Sub Com_Ordernew_Click()
    '����u��Ʀr��J
    
    If Not IsNumeric(Com_Orderamount) Or Com_Orderamount.Value = Empty Then
        MsgBox ("�п�J���T�ƶq!")
        Exit Sub
    End If
    
    '�N��~�s�W���I�椺�e
    Dim record As String
    record = List_Ordermenu.Value & Com_Orderamount.Value & "��"
    
    List_Orderrecord.AddItem (record)
    Com_Orderamount.Value = Null

End Sub

Private Sub UserForm_Initialize()
    '�I�\����l�Ƭɭ�
    
    '����H�ۤu�@���s

    Sheets("���޲z").Select
    Dim cnt As Integer
    Dim noodlecnt As Integer
    Dim snakecnt As Integer
    Dim drinkcnt As Integer
    Dim ordermenurow As Integer
    Dim noodlearray() As String
    Dim snakearray() As String
    Dim drinkarray() As String
    Dim allarray() As String
    ordermenurow = ActiveSheet.UsedRange.Rows.Count
    
    For cnt = 1 To ordermenurow
        Select Case Cells(cnt, "D").Value
        Case "�ѭ�"
            noodlecnt = noodlecnt + 1
        Case "�I��"
            snakecnt = snakecnt + 1
        Case "����"
            drinkcnt = drinkcnt + 1
        End Select
    Next
    
    noodlecnt = noodlecnt - 1
    snakecnt = snakecnt - 1
    drinkcnt = drinkcnt - 1
    cnt = ordermenurow - 2
    ReDim Preserve noodlearray(noodlecnt)
    ReDim Preserve snakearray(snakecnt)
    ReDim Preserve drinkarray(drinkcnt)
    ReDim Preserve allarray(cnt)
    
    noodlecnt = -1
    snakecnt = -1
    drinkcnt = -1
    
    For cnt = 2 To ordermenurow
        Select Case Cells(cnt, "D").Value
        Case "�ѭ�"
            noodlearray(noodlecnt + 1) = Cells(cnt, "C").Value
            noodlecnt = noodlecnt + 1
        Case "�I��"
            snakearray(snakecnt + 1) = Cells(cnt, "C").Value
            snakecnt = snakecnt + 1
        Case "����"
            drinkarray(drinkcnt + 1) = Cells(cnt, "C").Value
            drinkcnt = drinkcnt + 1
        End Select
    Next
    
    
    
    '�N�T�Ӱ}�C�X�֦��@�ӡA�̾��ѭ��I�߶��ƶ��ǱƧ�
    Dim allcnt As Integer
    Dim ncnt As Integer
    Dim scnt As Integer
    Dim acnt As Integer
    ncnt = noodlecnt
    scnt = snakecnt
    acnt = noodlecnt + snakecnt
    '�ѭ�
    If noodlecnt > 0 Then
        For allcnt = 0 To noodlecnt
            allarray(allcnt) = noodlearray(allcnt)
        Next
    ElseIf noodlecnt = 0 Then
        allarray(0) = noodlearray(0)
    End If
    
    '�I��
    If snakecnt > 0 Then
        If noodlecnt > 0 Or noodlecnt = 0 Then
            For allcnt = 0 To snakecnt
                allarray(ncnt + 1) = snakearray(allcnt)
                ncnt = ncnt + 1
            Next
        ElseIf noodlecnt < 0 Then
            For allcnt = 0 To snakecnt
                allarray(allcnt) = snakearray(allcnt)
            Next
        End If
    ElseIf snakecnt = 0 Then
        If noodlecnt > 0 Or noodlecnt = 0 Then
            allarray(ncnt + 1) = snakearray(0)
        End If
    End If
    
    '����
    ncnt = noodlecnt
    
    If drinkcnt > 0 Then
        If snakecnt > 0 Then
            If noodlecnt > 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(acnt + 2) = drinkarray(allcnt)
                    acnt = acnt + 1
                Next
            ElseIf noodlecnt < 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(scnt) = drinkarray(allcnt) '�n���n�[1
                    scnt = scnt + 1
                Next
            End If
        ElseIf snakecnt = 0 Then
            If noodlecnt > 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(ncnt + 1) = drinkarray(allcnt) '�n���n�[1
                    ncnt = ncnt + 1
                Next
            ElseIf noodlecnt = 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(allcnt) = drinkarray(allcnt)
                Next
            End If
        End If
    ElseIf drinkcnt = 0 Then '��ܦ��@�����ƫ~��
        If snakecnt > 0 Or snakecnt = 0 Then
            If noodlecnt > 0 Or noodlecnt = 0 Then
                allarray(acnt + 1) = drinkarray(0)
            End If
        End If
    End If

    '�̷ӫ~����ܦbListBox
    List_Ordermenu.list() = allarray
    
    '�s�W�ƶq�ܤU�Ԧ����
    Dim amount As Integer
    For amount = 1 To 5
        Com_Orderamount.AddItem amount
    Next
End Sub
