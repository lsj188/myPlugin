
'******************************************************************************
'* File:     excel_import.vbs
'* Title:    将excel文档导入到模型（win7版）
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

'打开文件选择框获取导入文件
Dim filePath
filePath=BrowseForFile()

Dim HaveExcel
Dim RQ
RQ         = vbYes 'MsgBox("Is  Excel Installed on your machine ?", vbYesNo + vbInformation, "Confirmation")

If RQ = vbYes Then   
	
	HaveExcel = True
	' Open & Create  Excel Document
	
	Dim x1 '
	Set x1 = CreateObject("Excel.Application")
	x1.Workbooks.Open filePath '指定 excel文档路径
	x1.Workbooks(1).Worksheets("表结构").Activate '指定要打开的sheet名称
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
	Dim null_cnt

	On Error Resume Next

	With x1.Workbooks(1).Worksheets("表结构")
		flag = 0
		null_cnt = 0
		
		For rwIndex = 1 To 655536 '指定要遍历的 Excel行标 由于第1行是 表头， 从第2行开始
			' 如果excel没有填写表名退出
			If .Cells(rwIndex, 1).Value = "" And rwIndex = 1 Then
				Exit For
			End If
			
			' 如果连续超过5个空行都没有取到数据就退出
			If null_cnt >= 5 Then
				Exit For
			End If

			' 多个实体以空行分格
			If .Cells(rwIndex, 1).Value = "" And rwIndex > 1 Then
				flag = 0
				null_cnt = null_cnt+1
			Else

			    null_cnt=0
				' 创建表
				If flag = 0 Then
					Set table  = mdl.Tables.CreateNew '创建一个 表实体
					table.Name = .Cells(rwIndex, 2).Value '指定 表名，如果在 Excel文档里有，也可以 .Cells(rwIndex, 3).Value 这样指定
					table.Code = .Cells(rwIndex, 1).Value '指定 表名
					Count      = Count + 1
					flag       = 1
					rwIndex=rwIndex+1
				Else
					Set col    = table.Columns.CreateNew '创建一列/字段

					' 如果excel没有填写字段名，字段名与code相同
					If .Cells(rwIndex, 1).Value = "" Then
						col.Name = .Cells(rwIndex, 2).Value '指定列名
					Else
						col.Name = .Cells(rwIndex, 1).Value
					End If
					
					col.Code   = .Cells(rwIndex, 2).Value '指定列名

					col.DataType = .Cells(rwIndex, 3).Value '指定列数据类型
					
					' 如果excel没有填写字段名，字段名与code相同
					If .Cells(rwIndex, 4).Value = "" Then
						col.Comment = col.Name '指定列名
					Else
						col.Comment = .Cells(rwIndex, 4).Value '指定列说明
					End If

					If .Cells(rwIndex, 5).Value = "Y" Then
						col.Primary = True '指定主键
					End If
					
					If .Cells(rwIndex, 6).Value = "N" Then
						col.Mandatory = True '指定列是否可空 true 为不可空 
					End If		

					col.defaultvalue = .Cells(rwIndex, 7).Value '默认值 

				End If

			End If

		Next

	End With

	MsgBox "生成数据 表结构共计 " + CStr(Count), vbOK + vbInformation, " 表"
    
	x1.Quit '关闭EXCEL
	Set x1 = Nothing '释放EXCEL对象
	Kill ("EXCEL.EXE")
	Exit Sub
	End Sub
	
'-------------------------------------
'文件选择框win7版
Function BrowseForFile()
    Dim shell : Set shell = CreateObject("WScript.Shell")
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
    Dim tempName : tempName = fso.GetTempName()
    Dim tempFile : Set tempFile = tempFolder.CreateTextFile(tempName & ".hta")
    tempFile.Write _
    "<html>" & _
    "<head>" & _
    "<title>Browse</title>" & _
    "</head>" & _
    "<body>" & _
    "<input type='file' id='f' />" & _
    "<script type='text/javascript'>" & _
    "var f = document.getElementById('f');" & _
    "f.click();" & _
    "var shell = new ActiveXObject('WScript.Shell');" & _
    "shell.RegWrite('HKEY_CURRENT_USER\\Volatile Environment\\MsgResp', f.value);" & _
    "window.close();" & _
    "</script>" & _
    "</body>" & _
    "</html>"
    tempFile.Close
    shell.Run tempFolder & "\" & tempName & ".hta", 0, True
    BrowseForFile = shell.RegRead("HKEY_CURRENT_USER\Volatile Environment\MsgResp")
    shell.RegDelete "HKEY_CURRENT_USER\Volatile Environment\MsgResp"
End Function

Private Sub Kill(str)
 Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
 Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & str & "'")
 For Each objProcess In colProcessList
   objProcess.Terminate '结束进程
 Next
 Set objProcess = Nothing
 Set colProcessList = Nothing
 Set objWMIService = Nothing
End Sub