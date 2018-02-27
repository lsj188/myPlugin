'******************************************************************************
'* File:     export_excel.vbs
'* Title:    ��ģ�͵�����excel
'* Author:   lsj qq:273364475
'* Created:  2017-11-09
'* Mod By:   
'* Modified: 
'* Version:  1.0
'* Comment:  
'*  v1.0 
'******************************************************************************

'******************************************************************************
Option Explicit
Dim rowsNum
rowsNum = 0
'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------
' Get the current active model
Dim Model
Set Model = ActiveModel

If (Model Is Nothing) Or (Not Model.IsKindOf(PdPDM.cls_Model)) Then
	MsgBox "The current model is not an PDM model."
Else
	' Get the tables collection
	'����EXCEL APP
	Dim beginrow
	Dim EXCEL
	Dim SHEET
	Dim SHEETLIST
	Set EXCEL                         = CreateObject("Excel.Application")
	EXCEL.workbooks.add( - 4167)'��ӹ�����
	EXCEL.workbooks(1).sheets(1).name = "��ṹ"
	Set SHEET                         = EXCEL.workbooks(1).sheets("��ṹ")

	EXCEL.workbooks(1).sheets.add
	EXCEL.workbooks(1).sheets(1).name = "Ŀ¼"
	Set SHEETLIST                     = EXCEL.workbooks(1).sheets("Ŀ¼")
	ShowTableList Model,SHEETLIST

	ShowProperties Model, SHEET,SHEETLIST

	EXCEL.workbooks(1).Sheets(2).Select
	EXCEL.visible                       = True
	'�����п���Զ�����
	sheet.Columns(1).ColumnWidth        = 20
	sheet.Columns(2).ColumnWidth        = 20
	sheet.Columns(3).ColumnWidth        = 20
	sheet.Columns(4).ColumnWidth        = 40
	sheet.Columns(5).ColumnWidth        = 10
	sheet.Columns(6).ColumnWidth        = 10
	sheet.Columns(1).WrapText           = True
	sheet.Columns(2).WrapText           = True
	sheet.Columns(4).WrapText           = True
	'����ʾ������
	EXCEL.ActiveWindow.DisplayGridlines = False

End If

'-----------------------------------------------------------------------------
' Show properties of tables
'-----------------------------------------------------------------------------
Sub ShowProperties(mdl, sheet,SheetList)
	' Show tables of the current model/package
	rowsNum  = 0
	beginrow = rowsNum + 1
	Dim rowIndex
	rowIndex = 3
	' For each table
	output "begin"
	Dim tab

	For Each tab In mdl.tables
		ShowTable tab,sheet,rowIndex,sheetList
		rowIndex = rowIndex + 1
	Next

	If mdl.tables.Count > 0 Then
		sheet.Range("A" & beginrow + 1 & ":A" & rowsNum).Rows.Group
	End If

	output "end"
End Sub

'-----------------------------------------------------------------------------
' Show table properties
'-----------------------------------------------------------------------------
Sub ShowTable(tab, sheet,rowIndex,sheetList)

	If IsObject(tab) Then
		Dim rangFlag
		rowsNum = rowsNum + 1
		' Show properties
		Output "================================"
		sheet.cells(rowsNum, 1) = tab.name
		sheet.cells(rowsNum, 1).HorizontalAlignment = 3
		sheet.cells(rowsNum, 2) = tab.code
		'sheet.cells(rowsNum, 5).HorizontalAlignment=3
		'sheet.cells(rowsNum, 6) = ""
		'sheet.cells(rowsNum, 7) = "��˵��"
		sheet.cells(rowsNum, 3) = tab.comment
		'sheet.cells(rowsNum, 8).HorizontalAlignment=3
		sheet.Range(sheet.cells(rowsNum, 3),sheet.cells(rowsNum, 7)).Merge
		'���ó����ӣ���Ŀ¼�������ȥ�鿴��ṹ
		'�ֶ�������    �ֶ�Ӣ����    �ֶ�����    ע��    �Ƿ�����    �Ƿ�ǿ�    Ĭ��ֵ
		sheetList.Hyperlinks.Add sheetList.cells(rowIndex,2), "","��ṹ" & "!B" & rowsNum
		rowsNum = rowsNum + 1
		sheet.cells(rowsNum, 1) = "�ֶ�������"
		sheet.cells(rowsNum, 2) = "�ֶ�Ӣ����"
		sheet.cells(rowsNum, 3) = "�ֶ�����"
		sheet.cells(rowsNum, 4) = "ע��"
		sheet.cells(rowsNum, 5) = "�Ƿ�����"
		sheet.cells(rowsNum, 6) = "�Ƿ�ǿ�"
		sheet.cells(rowsNum, 7) = "Ĭ��ֵ"
		'���ñ߿�
		sheet.Range(sheet.cells(rowsNum - 1, 1),sheet.cells(rowsNum, 7)).Borders.LineStyle = "1"
		'sheet.Range(sheet.cells(rowsNum-1, 4),sheet.cells(rowsNum, 9)).Borders.LineStyle = "1"
		'����Ϊ10��
		sheet.Range(sheet.cells(rowsNum - 1, 1),sheet.cells(rowsNum, 7)).Font.Size = 10
		Dim col ' running column
		Dim colsNum
		colsNum  = 0

		For Each col in tab.columns
			rowsNum = rowsNum + 1
			colsNum = colsNum + 1
			sheet.cells(rowsNum, 1) = col.name
			'sheet.cells(rowsNum, 3) = ""
			'sheet.cells(rowsNum, 4) = col.name
			sheet.cells(rowsNum, 2) = col.code
			sheet.cells(rowsNum, 3) = col.datatype
			sheet.cells(rowsNum, 4) = col.comment

			If col.Primary = True Then
				sheet.cells(rowsNum, 5) = "Y"
			Else
				sheet.cells(rowsNum, 5) = " "
			End If

			If col.Mandatory = True Then
				sheet.cells(rowsNum, 6) = "Y"
			Else
				sheet.cells(rowsNum, 6) = " "
			End If

			sheet.cells(rowsNum, 7) = col.defaultvalue
		Next

		sheet.Range(sheet.cells(rowsNum - colsNum + 1,1),sheet.cells(rowsNum,7)).Borders.LineStyle = "3"
		'sheet.Range(sheet.cells(rowsNum-colsNum+1,4),sheet.cells(rowsNum,9)).Borders.LineStyle = "3"
		sheet.Range(sheet.cells(rowsNum - colsNum + 1,1),sheet.cells(rowsNum,7)).Font.Size = 10
		rowsNum = rowsNum + 2

		Output "FullDescription: " + tab.Name
	End If

End Sub

'-----------------------------------------------------------------------------
' Show List Of Table
'-----------------------------------------------------------------------------
Sub ShowTableList(mdl, SheetList)
	' Show tables of the current model/package
	Dim rowsNo
	rowsNo = 1
	' For each table
	output "begin"
	SheetList.cells(rowsNo, 1) = "����"
	SheetList.cells(rowsNo, 2) = "��������"
	SheetList.cells(rowsNo, 3) = "��Ӣ����"
	SheetList.cells(rowsNo, 4) = "��˵��"
	rowsNo = rowsNo + 1
	SheetList.cells(rowsNo, 1) = mdl.name
   
	Dim pak
    For Each pak In mdl.Packages
		rowsNo = rowsNo + 1
		SheetList.cells(rowsNo, 1) = pak.name
	next
	Dim tab

	For Each tab In mdl.tables

		If IsObject(tab) Then
			rowsNo = rowsNo + 1
			SheetList.cells(rowsNo, 1) = ""
			SheetList.cells(rowsNo, 2) = tab.name
			SheetList.cells(rowsNo, 3) = tab.code
			SheetList.cells(rowsNo, 4) = tab.comment
         Dim diag
         For Each diag In tab.Diagrams
            SheetList.cells(rowsNo, 5) = SheetList.cells(rowsNo, 5)+','+diag.Name
         next
		End If

	Next

	SheetList.Columns(1).ColumnWidth = 20
	SheetList.Columns(2).ColumnWidth = 20
	SheetList.Columns(3).ColumnWidth = 30
	SheetList.Columns(4).ColumnWidth = 60
	output "end"
End Sub
