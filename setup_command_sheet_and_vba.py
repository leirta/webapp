from pathlib import Path

import pythoncom
import win32com.client as win32


FILE_PATH = Path(r"C:\Users\user\Documents\New project\報表_自動月報.xlsm")
FALLBACK_XLSX = Path(r"C:\Users\user\Documents\New project\報表.xlsx")
XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52

VBA_CODE = r'''
Option Explicit

Public Sub 更新各區月報()
    Dim wsSrc As Worksheet, wsCmd As Worksheet, wsRegion As Worksheet
    Dim targetMonth As String, region As Variant
    Dim regions As Variant
    Dim lastRow As Long, r As Long
    Dim rowOut As Long, detailRow As Long, startRow As Long
    Dim invoiceType As Variant, customer As Variant, srcRow As Variant
    Dim dictInvoice As Object, dictCustomer As Object
    Dim invKey As String, custKey As String
    Dim itemRows As Collection, custRows As Collection
    Dim hasData As Boolean

    On Error GoTo ErrHandler

    Set wsSrc = ThisWorkbook.Worksheets("總表")
    Set wsCmd = ThisWorkbook.Worksheets("指令")
    targetMonth = Trim(CStr(wsCmd.Range("B1").Value))

    If targetMonth = "" Then
        MsgBox "請先在指令分頁 B1 輸入請款月份（例如 11503）。", vbExclamation
        Exit Sub
    End If

    regions = Array("台北", "桃園", "新竹", "台中", "台南", "高雄")
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    For Each region In regions
        Set wsRegion = GetOrCreateSheet(CStr(region) & "月報")
        wsRegion.Cells.Clear
        rowOut = 1
        hasData = False

        Set dictInvoice = CreateObject("Scripting.Dictionary")

        For r = 2 To lastRow
            If CStr(wsSrc.Cells(r, 1).Value) = targetMonth And CStr(wsSrc.Cells(r, 15).Value) = CStr(region) Then
                hasData = True
                invKey = Trim(CStr(wsSrc.Cells(r, 5).Value))
                If invKey = "" Then invKey = "(無發票別)"
                If Not dictInvoice.Exists(invKey) Then
                    Set itemRows = New Collection
                    dictInvoice.Add invKey, itemRows
                End If
                dictInvoice(invKey).Add CLng(r)
            End If
        Next r

        If Not hasData Then
            wsRegion.Range("A1").Value = "區域"
            wsRegion.Range("B1").Value = CStr(region)
            wsRegion.Range("A2").Value = "請款月份"
            wsRegion.Range("B2").Value = targetMonth
            wsRegion.Range("A3").Value = "狀態"
            wsRegion.Range("B3").Value = "本月份無資料"
            ApplyRegionColumns wsRegion
            GoTo NextRegion
        End If

        For Each invoiceType In dictInvoice.Keys
            wsRegion.Cells(rowOut, 1).Value = "請款月份"
            wsRegion.Cells(rowOut, 2).Value = targetMonth
            wsRegion.Cells(rowOut + 1, 1).Value = "發票別"
            wsRegion.Cells(rowOut + 1, 2).Value = CStr(invoiceType)
            wsRegion.Cells(rowOut + 2, 1).Value = "區域"
            wsRegion.Cells(rowOut + 2, 2).Value = CStr(region)

            wsRegion.Cells(rowOut + 3, 1).Value = "客戶別"
            wsRegion.Cells(rowOut + 3, 2).Value = "公司抬頭"
            wsRegion.Cells(rowOut + 3, 3).Value = "報表名稱"
            wsRegion.Cells(rowOut + 3, 4).Value = "項目"
            wsRegion.Cells(rowOut + 3, 5).Value = "未稅額 "
            wsRegion.Cells(rowOut + 3, 6).Value = "稅金 "
            wsRegion.Cells(rowOut + 3, 7).Value = "小計 "
            wsRegion.Cells(rowOut + 3, 8).Value = "發票號碼"
            wsRegion.Cells(rowOut + 3, 9).Value = "備註"

            Set dictCustomer = CreateObject("Scripting.Dictionary")

            For Each srcRow In dictInvoice(invoiceType)
                custKey = Trim(CStr(wsSrc.Cells(CLng(srcRow), 6).Value))
                If custKey = "" Then custKey = "(無客戶別)"
                If Not dictCustomer.Exists(custKey) Then
                    Set custRows = New Collection
                    dictCustomer.Add custKey, custRows
                End If
                dictCustomer(custKey).Add CLng(srcRow)
            Next srcRow

            detailRow = rowOut + 4
            For Each customer In dictCustomer.Keys
                startRow = detailRow

                For Each srcRow In dictCustomer(customer)
                    wsRegion.Cells(detailRow, 1).Value = wsSrc.Cells(CLng(srcRow), 6).Value
                    wsRegion.Cells(detailRow, 2).Value = wsSrc.Cells(CLng(srcRow), 18).Value
                    wsRegion.Cells(detailRow, 3).Value = wsSrc.Cells(CLng(srcRow), 7).Value
                    wsRegion.Cells(detailRow, 4).Value = wsSrc.Cells(CLng(srcRow), 8).Value
                    wsRegion.Cells(detailRow, 5).Value = wsSrc.Cells(CLng(srcRow), 12).Value
                    wsRegion.Cells(detailRow, 6).Value = wsSrc.Cells(CLng(srcRow), 13).Value
                    wsRegion.Cells(detailRow, 7).Value = wsSrc.Cells(CLng(srcRow), 14).Value
                    wsRegion.Cells(detailRow, 8).Value = wsSrc.Cells(CLng(srcRow), 3).Value
                    wsRegion.Cells(detailRow, 9).Value = wsSrc.Cells(CLng(srcRow), 16).Value
                    detailRow = detailRow + 1
                Next srcRow

                wsRegion.Cells(detailRow, 1).Value = CStr(customer) & " 合計"
                wsRegion.Cells(detailRow, 5).Formula = "=SUM(E" & startRow & ":E" & (detailRow - 1) & ")"
                wsRegion.Cells(detailRow, 6).Formula = "=SUM(F" & startRow & ":F" & (detailRow - 1) & ")"
                wsRegion.Cells(detailRow, 7).Formula = "=SUM(G" & startRow & ":G" & (detailRow - 1) & ")"
                detailRow = detailRow + 2
            Next customer

            rowOut = detailRow
        Next invoiceType

        ApplyRegionColumns wsRegion
NextRegion:
    Next region

    MsgBox "已完成：" & vbCrLf & _
           "請款月份 " & targetMonth & " 的六區月報已更新。", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "更新失敗：" & Err.Description, vbCritical
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub ApplyRegionColumns(ByVal ws As Worksheet)
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 20
End Sub
'''


def get_or_create_sheet(workbook, name):
    for ws in workbook.Worksheets:
        if ws.Name == name:
            return ws
    ws = workbook.Worksheets.Add(Before=workbook.Worksheets(1))
    ws.Name = name
    return ws


def remove_module_if_exists(vb_project, module_name):
    for component in vb_project.VBComponents:
        if component.Name == module_name:
            vb_project.VBComponents.Remove(component)
            break


def main():
    if not FILE_PATH.exists() and not FALLBACK_XLSX.exists():
        raise FileNotFoundError(f"找不到檔案: {FILE_PATH} 或 {FALLBACK_XLSX}")

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = None
        try:
            if FILE_PATH.exists():
                workbook = excel.Workbooks.Open(str(FILE_PATH))
        except Exception:
            workbook = None

        if workbook is None:
            if not FALLBACK_XLSX.exists():
                raise FileNotFoundError(f"無法開啟 {FILE_PATH}，且找不到 {FALLBACK_XLSX}")
            workbook = excel.Workbooks.Open(str(FALLBACK_XLSX))
            workbook.SaveAs(str(FILE_PATH), FileFormat=XL_OPENXML_WORKBOOK_MACRO_ENABLED)
            workbook.Close(SaveChanges=True)
            workbook = excel.Workbooks.Open(str(FILE_PATH))

        cmd_sheet = get_or_create_sheet(workbook, "指令")
        cmd_sheet.Cells.Clear()
        cmd_sheet.Range("A1").Value = "請款月份"
        cmd_sheet.Range("B1").Value = workbook.Worksheets("總表").Range("A2").Value
        cmd_sheet.Range("A2").Value = "說明"
        cmd_sheet.Range("B2").Value = "請在 B1 輸入月份（例如 11503），再按下按鈕。"
        cmd_sheet.Columns("A").ColumnWidth = 16
        cmd_sheet.Columns("B").ColumnWidth = 42

        for shape in list(cmd_sheet.Shapes):
            if shape.Name == "btnUpdateRegionReports":
                shape.Delete()

        button = cmd_sheet.Shapes.AddFormControl(0, 20, 70, 180, 30)
        button.Name = "btnUpdateRegionReports"
        button.TextFrame.Characters().Text = "更新各區月報"
        button.OnAction = "更新各區月報"

        vb_project = workbook.VBProject
        remove_module_if_exists(vb_project, "modRegionReports")
        module = vb_project.VBComponents.Add(1)
        module.Name = "modRegionReports"
        module.CodeModule.AddFromString(VBA_CODE)

        workbook.Save()
        print("OK")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
