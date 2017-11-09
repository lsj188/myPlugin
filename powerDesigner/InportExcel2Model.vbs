
'******************************************************************************
'* File:     excel_inport.vbs
'* Title:    将excel文档导入到模型
'* Author:   lsj qq:273364475
'* Created:  2017-11-09
'* Mod By:   
'* Modified: 
'* Version:  1.0
'* Comment:  
'*  v1.0 
'******************************************************************************

Option Explicit

Dim mdl ' the current model
Set mdl = ActiveModel

If (mdl Is Nothing) Then
	MsgBox "There is no Active Model"
End If

Dim HaveExcel
Dim RQ
RQ         = vbYes 'MsgBox("Is  Excel Installed on your machine ?", vbYesNo + vbInformation, "Confirmation")

If RQ = vbYes Then
	HaveExcel = True
	' Open & Create  Excel Document
	Dim x1 '
	Set x1 = CreateObject("Excel.Application")
	x1.Workbooks.Open "D:\Program Files (x86)\Sybase\PowerDesigner 16\execl2power\11.xlsx" '指定 excel文档路径
	x1.Workbooks(1).Worksheets("Sheet1").Activate '指定要打开的sheet名称
Else
	HaveExcel = False
End If

a x1, mdl

Sub a(x1, mdl)
	Dim rwIndex
	Dim tableName
	Dim colname
	Dim table
	Dim col
	Dim Count
	Dim flag

	On Error Resume Next

	With x1.Workbooks(1).Worksheets("Sheet1")
		flag = 0

		For rwIndex = 1 To 1000 '指定要遍历的 Excel行标 由于第1行是 表头， 从第2行开始
			' 如果excel没有填写表名退出

			If .Cells(rwIndex, 1).Value = "" And rwIndex = 1 Then
				Exit For
			End If

			' 多个实体以空行分格

			If .Cells(rwIndex, 1).Value = "" And rwIndex > 1 Then
				flag = 0
			Else

				' 创建表

				If flag = 0 Then
					Set table  = mdl.Tables.CreateNew '创建一个 表实体
					table.Name = .Cells(rwIndex, 2).Value '指定 表名，如果在 Excel文档里有，也可以 .Cells(rwIndex, 3).Value 这样指定
					table.Code = .Cells(rwIndex, 1).Value '指定 表名
					Count      = Count + 1
					flag       = 1
				Else
					Set col    = table.Columns.CreateNew '创建一列/字段
					col.Code   = .Cells(rwIndex, 1).Value '指定列名

					' 如果excel没有填写字段名，字段名与code相同

					If .Cells(rwIndex, 2).Value = "" Then
						col.Name = .Cells(rwIndex, 1).Value '指定列名
					Else
						col.Name = .Cells(rwIndex, 2).Value
					End If

					col.DataType = .Cells(rwIndex, 3).Value '指定列数据类型

					If .Cells(rwIndex, 4).Value = "否" Then
						col.Mandatory = True '指定列是否可空 true 为不可空 
					End If

					If .Cells(rwIndex, 5).Value = "是" Then
						col.Primary = True '指定主键
					End If

					' 如果excel没有填写字段名，字段名与code相同

					If .Cells(rwIndex, 6).Value = "" Then
						col.Name = col.Name '指定列名
					Else
						col.Comment = .Cells(rwIndex, 6).Value '指定列说明
					End If

				End If

			End If

		Next

	End With

	MsgBox "生成数据 表结构共计 " + CStr(Count), vbOK + vbInformation, " 表"

	Exit Sub
	End Sub
