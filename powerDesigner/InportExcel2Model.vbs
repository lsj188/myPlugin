
'******************************************************************************
'* File:     excel_inport.vbs
'* Title:    ��excel�ĵ����뵽ģ�ͣ�win7�棩
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

'���ļ�ѡ����ȡ�����ļ�
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
	x1.Workbooks.Open filePath 'ָ�� excel�ĵ�·��
	x1.Workbooks(1).Worksheets("Sheet1").Activate 'ָ��Ҫ�򿪵�sheet����
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

		For rwIndex = 1 To 1000 'ָ��Ҫ������ Excel�б� ���ڵ�1���� ��ͷ�� �ӵ�2�п�ʼ
			' ���excelû����д�����˳�

			If .Cells(rwIndex, 1).Value = "" And rwIndex = 1 Then
				Exit For
			End If

			' ���ʵ���Կ��зָ�

			If .Cells(rwIndex, 1).Value = "" And rwIndex > 1 Then
				flag = 0
			Else

				' ������

				If flag = 0 Then
					Set table  = mdl.Tables.CreateNew '����һ�� ��ʵ��
					table.Name = .Cells(rwIndex, 2).Value 'ָ�� ����������� Excel�ĵ����У�Ҳ���� .Cells(rwIndex, 3).Value ����ָ��
					table.Code = .Cells(rwIndex, 1).Value 'ָ�� ����
					Count      = Count + 1
					flag       = 1
				Else
					Set col    = table.Columns.CreateNew '����һ��/�ֶ�
					col.Code   = .Cells(rwIndex, 1).Value 'ָ������

					' ���excelû����д�ֶ������ֶ�����code��ͬ

					If .Cells(rwIndex, 2).Value = "" Then
						col.Name = .Cells(rwIndex, 1).Value 'ָ������
					Else
						col.Name = .Cells(rwIndex, 2).Value
					End If

					col.DataType = .Cells(rwIndex, 3).Value 'ָ������������

					If .Cells(rwIndex, 4).Value = "��" Then
						col.Mandatory = True 'ָ�����Ƿ�ɿ� true Ϊ���ɿ� 
					End If

					If .Cells(rwIndex, 5).Value = "��" Then
						col.Primary = True 'ָ������
					End If

					' ���excelû����д�ֶ������ֶ�����code��ͬ

					If .Cells(rwIndex, 6).Value = "" Then
						col.Name = col.Name 'ָ������
					Else
						col.Comment = .Cells(rwIndex, 6).Value 'ָ����˵��
					End If

				End If

			End If

		Next

	End With

	MsgBox "�������� ��ṹ���� " + CStr(Count), vbOK + vbInformation, " ��"
    x1.Workbooks.close False
	Exit Sub
	End Sub
	
'-------------------------------------
'�ļ�ѡ���win7��
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
