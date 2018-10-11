
'******************************************************************************
'* File:     excel_import.vbs
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

'----------------------------------Ŀ¼ҳ����-----------------------------------
CONST COL_TABLE_SCHEMA          = "B"           '��ģʽ�У�Owner��
CONST COL_TABLE_CODE            = "C"           '��Ӣ������
CONST COL_TABLE_NAME            = "D"           '����������
CONST COL_DEAL_FLAG             = "E"           '�����־��
CONST COL_TABLE_COMMENT         = "F"           '��˵���У�Comment��
'----------------------------------����ҳ����-----------------------------------
CONST COL_COL_NAME              = "B"           '�ֶ�������
CONST COL_COL_CODE              = "C"           '�ֶ�Ӣ����
CONST COL_COL_DATATYPE          = "D"           '�ֶ�����
CONST COL_COL_PRIMARY           = "F"           '����
CONST COL_COL_MANDATORY         = "G"           '�ǿ�
CONST COL_COL_DISTRIBUTIONKEYS  = "H"           '�ֲ���
CONST COL_COL_COMMENT           = "I"           '˵����Comment��
'-------------------------------------------------------------------------------
CONST BEG_ROW = 6                               '��������-��ʼ��
CONST MAX_TABLES       = 1000                   '����������
CONST MAX_COLUMNS      = 1000                   '�����ֶ���������
CONST PHYSICAL_OPTIONS = "WITH(APPENDONLY=TRUE,COMPRESSLEVEL=6,ORIENTATION=COLUMN,COMPRESSTYPE=ZLIB)"

Dim mCR,mLF
mCR = Chr(10)       '����
mLF = Chr(13)       '�س�
'-------------------------------------------------------------------------------

'����PDM
Dim mdl
Dim input_file, log_file
Dim errCount, errString, errMsg
errMsg=""
errCount=0
Set mdl = ActiveModel
If ( mdl Is Nothing ) Then
    MsgBox "There is no Active Model"
Else
    '���ļ�ѡ����ȡ�����ļ�
    input_file=BrowseForFile()
    log_file=left(input_file,instrrev(input_file,".",-1,1))+".log"
    
    Output "input_file "&input_file
    output "log_file   "&log_file
    output "========================================================"
    
    Dim HaveExcel
    Dim RQ
    RQ         = vbYes 'MsgBox("Is  Excel Installed on your machine ?", vbYesNo + vbInformation, "Confirmation")
    
    If RQ = vbYes Then   
        
        HaveExcel = True
        ' Open & Create  Excel Document
        Dim ExcelApp, x1
        Set ExcelApp = CreateObject("Excel.Application")
        ExcelApp.visible=FALSE
        'ExcelApp.Workbooks.Open(INPUT_FILE)
        'Set x1 = CreateObject("Excel.Application")
        Set x1=ExcelApp.Workbooks.Open(input_file) 'ָ�� excel�ĵ�·��
    Else
        HaveExcel = False
    End If
    
    importTables x1,mdl
    If errCount > 0 Then
        output "������Ϣ: " + errMsg
    End If
    MsgBox "�������,����"+Cstr(errCount)+"������!"
    If Trim(errMsg)<>"" Then
        MsgBox errMsg
    End If
End If


    

'�����ṹ
Sub importTables(x1,mdl)

    Dim ExcelBook, ExcelSheet

    Set ExcelBook = x1

    '-------------------��ȡĿ¼-------------------
    Dim tblSchema, tblName, tblCode, tblComment
    Dim tblIdx, tblCnt
    tblIdx = 0
    tblCnt = 0
    For tblIdx = 2 To MAX_TABLES+2
        If ExcelBook.Worksheets("Ŀ¼").Cells(tblIdx, "A").Value = "" Then
            Exit For
        End If

        If UCase(ExcelBook.Worksheets("Ŀ¼").Cells(tblIdx,COL_DEAL_FLAG).value) = "Y" Then
            tblCnt = tblCnt + 1

            tblSchema    = Trim(ExcelBook.Worksheets("Ŀ¼").Cells(tblIdx, COL_TABLE_SCHEMA).Value)
            tblCode      = Trim(ExcelBook.Worksheets("Ŀ¼").Cells(tblIdx, COL_TABLE_CODE).Value)
            tblName      = Trim(ExcelBook.Worksheets("Ŀ¼").Cells(tblIdx, COL_TABLE_NAME).Value)
            tblComment   = Trim(ExcelBook.Worksheets("Ŀ¼").Cells(tblIdx, COL_TABLE_COMMENT).Value)
            If Len(tblComment) = 0 Then
                tblComment = tblName
            End If

            '-------------------��ȡSheetҳ-------------------
            On Error Resume Next
            Dim shtIdx
            shtIdx = ExcelBook.Worksheets(tblCode).Index
            If Err.Number <> 0 Then
                output "Table[" + tblCode + "][" + tblName + "] �Ҳ�����Sheetҳ��"
                errCount  = errCount + 1
                errString = cstr(now) + " <" + cstr(errCount) + "> [�ļ�����]---[" + tblCode + "][" + tblName + "] �Ҳ�����Sheetҳ��" + mLF
                        errMsg = errMsg + errString
            Else
                output "[" + tblCode + "][" + tblName + "]"
                ExcelBook.Worksheets(shtIdx).Activate

                '����
                createTable mdl,ExcelBook,shtIdx,tblName,tblCode,tblComment,tblSchema

            End If
            '-------------------��ȡSheetҳ END---------------
        End If
    Next
    '-------------------��ȡĿ¼ END---------------

    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    Kill ("EXCEL.EXE")

    output "�������, ������ " + Cstr(tblCnt) + " �ű�!"
    
    Dim fs, ft
    Set fs = CreateObject("scripting.filesystemobject")
    Set ft = fs.createtextfile(log_file)
    ft.WriteLine (errMsg)
    ft.Close
    Set ft = Nothing: Set fs = Nothing
    
    
    Exit Sub
End Sub

'����
Sub createTable(mdl,ExcelBook,shtIdx,tblName,tblCode,tblComment,tblOwner)

    '�����û�
    'Dim tblOwner
    'Set tblOwner = mdl.FindChildByCode(tblSchema,cls_User)
    'If ( tblOwner Is Nothing ) Then
    '    output "δ�ҵ��û�[" + tblSchema + "]"
    '    errString = errString + mLF + "δ�ҵ��û�[" + tblSchema + "]"
    '    errCount  = errCount + 1
    'End If

    '�����Ƿ��Ѿ�����
    Dim objTable, col
    set objTable = mdl.FindChildByName(tblName,cls_Table)
    If ( objTable Is Nothing ) Then
    Else
        output "��[" + tblName + "]�Ѿ����ڣ�"
        errCount  = errCount + 1
        errString = cstr(now) + " <" + cstr(errCount) + "> [�����]-----[" + tblName + "] �Ѿ����ڣ�" + mLF
        errMsg = errMsg + errString
        Exit Sub
    End If
    set objTable = mdl.FindChildByCode(tblCode,cls_Table)
    If ( objTable Is Nothing ) Then
    Else
        output "��[" + tblCode + "]�Ѿ����ڣ�"
        errCount  = errCount + 1
        errString = cstr(now) + " <" + cstr(errCount) + "> [�����]-----[" + tblCode + "] �Ѿ����ڣ�" + mLF
        errMsg = errMsg + errString
        Exit Sub
    End If

    '����
    Set objTable = mdl.Tables.CreateNew
    objTable.Name    = tblName
    objTable.Code    = tblCode
    objTable.Comment = tblComment
    'objTable.Owner   = tblOwner
    objTable.PhysicalOptions = PHYSICAL_OPTIONS

    '�����ֶ�
    Dim colIdx, colCnt
    colCnt = 0
    For colIdx = BEG_ROW To MAX_COLUMNS+BEG_ROW
        If ExcelBook.Worksheets(shtIdx).Cells(colIdx, "A").Value = "" Then
            Exit For
        End If

        Dim colName, colCode, colDataType, colComment, colMandatory, colPrimary, colDistributionKeys
        colName      = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_NAME).Value))       '�ֶ�������
        colCode      = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_CODE).Value))       '�ֶ�Ӣ����
        colDataType  = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_DATATYPE).Value))   '�ֶ��ֶ�����
        colPrimary   = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_PRIMARY).Value))    '����
        colMandatory = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_MANDATORY).Value))  '�ǿ�
        'colDistributionKeys =                                                                          '�ֲ���
        colComment   = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_COMMENT).Value))    '˵��
        If Len(colComment) = 0 Then
            colComment = colName
        End If

        '���ֶ�
        With ExcelBook
            Set col = objTable.Columns.CreateNew
            
            '�����Ƿ��Ѿ�����
                Dim objColumn
                    set objColumn = objTable.FindChildByName(colName,cls_Column)
                    If ( objColumn Is Nothing ) Then
                    Else
                        output "�ֶ�[" + colName + "]�Ѿ����ڣ�"
                        errCount  = errCount + 1
                        errString = cstr(now) + " <" + cstr(errCount) + "> [�ֶδ���]---[" + objTable.Name + "." + colName + "] �Ѿ����ڣ�" + mLF
                        errMsg = errMsg + errString
                        
                        Exit Sub
                    End If        
                    set objColumn = objTable.FindChildByCode(colCode,cls_Column)
                    If ( objColumn Is Nothing ) Then
                    Else
                        output "�ֶ�[" + colCode + "]�Ѿ����ڣ�"
                        errCount  = errCount + 1
                        errString = cstr(now) + " <" + cstr(errCount) + "> [�ֶδ���]---[" + objTable.Name + "." + colCode + "] �Ѿ����ڣ�" + mLF
                        errMsg = errMsg + errString
                        Exit Sub
                    End If        
            
            col.Name = colName
            col.Code = UCase(colCode)
            col.DataType = UCase(colDataType)
            col.Comment = colComment
            If UCase(colMandatory) = "Y" Or UCase(colMandatory) = "YES" Then
                col.Mandatory = true                'ָ�����Ƿ�ɿ� true Ϊ���ɿ�
            End If
            If UCase(colPrimary) = "Y" Or UCase(colPrimary) = "YES" Then
                col.Primary = true                  'ָ������
            End If
        End With
        colCnt = colCnt + 1

    Next

    '���÷ֲ���
    'If Len(colDistributionKeys) > 0 Then
    '    objTable.PhysicalOptions = objTable.PhysicalOptions + mLF + "distributed by (" + colDistributeKeys + ")"
    'End If

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

Private Sub Kill(str)
 Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
 Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & str & "'")
 For Each objProcess In colProcessList
   objProcess.Terminate '��������
 Next
 Set objProcess = Nothing
 Set colProcessList = Nothing
 Set objWMIService = Nothing
End Sub
