
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

'----------------------------------请按需设置-----------------------------------
CONST INPUT_FILE  = "D:\File_Name.xlsm"        '导入文件路径
CONST LOG_FILE  = "D:\pd_2_excel.log"          '日志文件路径
'----------------------------------目录页设置-----------------------------------
CONST COL_TABLE_SCHEMA          = "B"           '表模式列（Owner）
CONST COL_TABLE_CODE            = "C"           '表英文名列
CONST COL_TABLE_NAME            = "D"           '表中文名列
CONST COL_DEAL_FLAG             = "E"           '处理标志列
CONST COL_TABLE_COMMENT         = "F"           '表说明列（Comment）
'----------------------------------内容页设置-----------------------------------
CONST COL_COL_NAME              = "B"           '字段中文名
CONST COL_COL_CODE              = "C"           '字段英文名
CONST COL_COL_DATATYPE          = "D"           '字段类型
CONST COL_COL_PRIMARY           = "F"           '主键
CONST COL_COL_MANDATORY         = "G"           '非空
CONST COL_COL_DISTRIBUTIONKEYS  = "H"           '分布键
CONST COL_COL_COMMENT           = "I"           '说明（Comment）
'-------------------------------------------------------------------------------
CONST BEG_ROW = 6                               '数据区域-开始行
CONST MAX_TABLES       = 1000                   '表数量上限
CONST MAX_COLUMNS      = 1000                   '单表字段数量上限
CONST PHYSICAL_OPTIONS = "WITH(APPENDONLY=TRUE,COMPRESSLEVEL=6,ORIENTATION=COLUMN,COMPRESSTYPE=ZLIB)"

Dim mCR,mLF
mCR = Chr(10)       '换行
mLF = Chr(13)       '回车
'-------------------------------------------------------------------------------

'定义PDM
Dim mdl
Dim errCount, errString, errMsg
errMsg=""
errCount=0
Set mdl = ActiveModel
If ( mdl Is Nothing ) Then
    MsgBox "There is no Active Model"
Else
    importTables mdl
    If errCount > 0 Then
        output "错误信息: " + errMsg
    End If
    MsgBox "处理完毕,共有"+Cstr(errCount)+"个错误!"
    If Trim(errMsg)<>"" Then
        MsgBox errMsg
    End If
End If

'导入表结构
Sub importTables(mdl)

    Dim ExcelApp, ExcelBook, ExcelSheet

    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.visible=FALSE
    Set ExcelBook = ExcelApp.Workbooks.Open(INPUT_FILE)

    '-------------------读取目录-------------------
    Dim tblSchema, tblName, tblCode, tblComment
    Dim tblIdx, tblCnt
    tblIdx = 0
    tblCnt = 0
    For tblIdx = 2 To MAX_TABLES+2
        If ExcelBook.Worksheets("目录").Cells(tblIdx, "A").Value = "" Then
            Exit For
        End If

        If UCase(ExcelBook.Worksheets("目录").Cells(tblIdx,COL_DEAL_FLAG).value) = "Y" Then
            tblCnt = tblCnt + 1

            tblSchema    = Trim(ExcelBook.Worksheets("目录").Cells(tblIdx, COL_TABLE_SCHEMA).Value)
            tblCode      = Trim(ExcelBook.Worksheets("目录").Cells(tblIdx, COL_TABLE_CODE).Value)
            tblName      = Trim(ExcelBook.Worksheets("目录").Cells(tblIdx, COL_TABLE_NAME).Value)
            tblComment   = Trim(ExcelBook.Worksheets("目录").Cells(tblIdx, COL_TABLE_COMMENT).Value)
            If Len(tblComment) = 0 Then
                tblComment = tblName
            End If

            '-------------------读取Sheet页-------------------
            On Error Resume Next
            Dim shtIdx
            shtIdx = ExcelBook.Worksheets(tblCode).Index
            If Err.Number <> 0 Then
                output "Table[" + tblCode + "][" + tblName + "] 找不到该Sheet页！"
                errCount  = errCount + 1
                errString = cstr(now) + " <" + cstr(errCount) + "> [文件错误]---[" + tblCode + "][" + tblName + "] 找不到该Sheet页！" + mLF
                        errMsg = errMsg + errString
            Else
                output "[" + tblCode + "][" + tblName + "]"
                ExcelBook.Worksheets(shtIdx).Activate

                '建表
                createTable mdl,ExcelBook,shtIdx,tblName,tblCode,tblComment,tblSchema

            End If
            '-------------------读取Sheet页 END---------------
        End If
    Next
    '-------------------读取目录 END---------------

    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing

    output "导入完毕, 共导入 " + Cstr(tblCnt) + " 张表!"
    
    Dim fs, ft
    Set fs = CreateObject("scripting.filesystemobject")
    Set ft = fs.createtextfile(LOG_FILE)
    ft.WriteLine (errMsg)
    ft.Close
    Set ft = Nothing: Set fs = Nothing
    
    
    Exit Sub
End Sub

'建表
Sub createTable(mdl,ExcelBook,shtIdx,tblName,tblCode,tblComment,tblOwner)

    '查找用户
    'Dim tblOwner
    'Set tblOwner = mdl.FindChildByCode(tblSchema,cls_User)
    'If ( tblOwner Is Nothing ) Then
    '    output "未找到用户[" + tblSchema + "]"
    '    errString = errString + mLF + "未找到用户[" + tblSchema + "]"
    '    errCount  = errCount + 1
    'End If

    '检查表是否已经存在
    Dim objTable, col
    set objTable = mdl.FindChildByName(tblName,cls_Table)
    If ( objTable Is Nothing ) Then
    Else
        output "表[" + tblName + "]已经存在！"
        errCount  = errCount + 1
        errString = cstr(now) + " <" + cstr(errCount) + "> [表错误]-----[" + tblName + "] 已经存在！" + mLF
        errMsg = errMsg + errString
        Exit Sub
    End If
    set objTable = mdl.FindChildByCode(tblCode,cls_Table)
    If ( objTable Is Nothing ) Then
    Else
        output "表[" + tblCode + "]已经存在！"
        errCount  = errCount + 1
        errString = cstr(now) + " <" + cstr(errCount) + "> [表错误]-----[" + tblCode + "] 已经存在！" + mLF
        errMsg = errMsg + errString
        Exit Sub
    End If

    '建表
    Set objTable = mdl.Tables.CreateNew
    objTable.Name    = tblName
    objTable.Code    = tblCode
    objTable.Comment = tblComment
    'objTable.Owner   = tblOwner
    objTable.PhysicalOptions = PHYSICAL_OPTIONS

    '解析字段
    Dim colIdx, colCnt
    colCnt = 0
    For colIdx = BEG_ROW To MAX_COLUMNS+BEG_ROW
        If ExcelBook.Worksheets(shtIdx).Cells(colIdx, "A").Value = "" Then
            Exit For
        End If

        Dim colName, colCode, colDataType, colComment, colMandatory, colPrimary, colDistributionKeys
        colName      = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_NAME).Value))       '字段中文名
        colCode      = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_CODE).Value))       '字段英文名
        colDataType  = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_DATATYPE).Value))   '字段字段类型
        colPrimary   = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_PRIMARY).Value))    '主键
        colMandatory = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_MANDATORY).Value))  '非空
        'colDistributionKeys =                                                                          '分布键
        colComment   = Trim(CStr(ExcelBook.Worksheets(shtIdx).Cells(colIdx, COL_COL_COMMENT).Value))    '说明
        If Len(colComment) = 0 Then
            colComment = colName
        End If

        '建字段
        With ExcelBook
            Set col = objTable.Columns.CreateNew
            
            '检查表是否已经存在
                Dim objColumn
                    set objColumn = objTable.FindChildByName(colName,cls_Column)
                    If ( objColumn Is Nothing ) Then
                    Else
                        output "字段[" + colName + "]已经存在！"
                        errCount  = errCount + 1
                        errString = cstr(now) + " <" + cstr(errCount) + "> [字段错误]---[" + objTable.Name + "." + colName + "] 已经存在！" + mLF
                        errMsg = errMsg + errString
                        
                        Exit Sub
                    End If        
                    set objColumn = objTable.FindChildByCode(colCode,cls_Column)
                    If ( objColumn Is Nothing ) Then
                    Else
                        output "字段[" + colCode + "]已经存在！"
                        errCount  = errCount + 1
                        errString = cstr(now) + " <" + cstr(errCount) + "> [字段错误]---[" + objTable.Name + "." + colCode + "] 已经存在！" + mLF
                        errMsg = errMsg + errString
                        Exit Sub
                    End If        
            
            col.Name = colName
            col.Code = UCase(colCode)
            col.DataType = UCase(colDataType)
            col.Comment = colComment
            If UCase(colMandatory) = "Y" Or UCase(colMandatory) = "YES" Then
                col.Mandatory = true                '指定列是否可空 true 为不可空
            End If
            If UCase(colPrimary) = "Y" Or UCase(colPrimary) = "YES" Then
                col.Primary = true                  '指定主键
            End If
        End With
        colCnt = colCnt + 1

    Next

    '设置分布键
    'If Len(colDistributionKeys) > 0 Then
    '    objTable.PhysicalOptions = objTable.PhysicalOptions + mLF + "distributed by (" + colDistributeKeys + ")"
    'End If

End Sub