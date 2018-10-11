'******************************************************************************
'* File:     export_excel.vbs
'* Title:    将模型导出到excel
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
'----------------------------------请按需设置-----------------------------------
CONST GEN_MENU    = "Y"                         '是否生成目录文件 [ Y-是 N-否 ]
CONST GEN_TABLE   = "Y"                         '是否生成模型结构 [ Y-是 N-否 ]
CONST SHOW_DISTRIBUTION_KEYS  = "Y"             '是否显示分布键   [ Y-是 N-否 ]
'----------------------------------目录页设置-----------------------------------
CONST COL_TABLE_CODE = "C"                      '表英文名列
CONST COL_TABLE_NAME = "D"                      '表中文名列
CONST COL_DEAL_FLAG  = "E"                      '处理标志列
CONST COL_TABLE_COMMENT  = "F"                  '处理描述列
'-------------------------------------------------------------------------------
CONST BEG_ROW = 6                               '数据区域-开始行
CONST END_COL = "J"                             '数据区域-结束列
CONST MAX_TABLES = 1000                         '表数量上限

CONST DATA_TYPE_DATE_LEN      = 10              'DATE类型数据长度
CONST DATA_TYPE_TIMESTAMP_LEN = 19              'TIMESTAMP类型数据长度
CONST DATA_TYPE_INTEGER_LEN   = 12              'INTEGER类型数据长度

CONST D_COLOR_BLUE     = 16764057               '天蓝色
CONST D_COLOR_GREEN    = 13434828               '浅绿色
CONST D_COLOR_ORAGNE   = 49407                  '橙色

Dim mCR,mLF
mCR = Chr(10)       '换行
mLF = Chr(13)       '回车
'-------------------------------------------------------------------------------

'定义PDM
Dim mdl
Dim file_path
Dim errCount, errString
errCount=0
Set mdl = ActiveModel
If ( mdl Is Nothing ) Then
    MsgBox "There is no Active Model"
Else
    file_path=InputBox("请输入导出文件路径")
    If UCase(GEN_MENU) = "Y" Then
        createMenuSheet mdl         '生成目录
    End If

    If UCase(GEN_TABLE) = "Y" Then
        createTableSheet mdl        '根据目录生成表结构
    End If

    If errCount > 0 Then
        output "错误信息: " + errString
    End If
    MsgBox "处理完毕,共有"+Cstr(errCount)+"个错误!"
End If

'-------------------------------------------------------------------------------
'生成目录
'   序号|模式名|表名|处理标志(Y/N)|中文表名|备注
'   处理标志默认全部为Y
'-------------------------------------------------------------------------------
sub createMenuSheet(mdl)

    Dim ExcelApp, ExcelBook, ExcelSheet

    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.visible=FALSE
    Set ExcelBook = ExcelApp.Workbooks.Add
    Set ExcelSheet = ExcelBook.Sheets.Add
    ExcelSheet.Name = "目录"

    '目录标题栏
    With ExcelSheet
        '内容
        .Cells(1,"A").Value = "序号"
        .Cells(1,"B").Value = "模式名"
        .Cells(1,"C").Value = "表英文名"
        .Cells(1,"D").Value = "中文表名"
        .Cells(1,"E").Value = "处理标志(Y/N)"
        .Cells(1,"F").Value = "备注"

        '样式-居中
        .Rows(1).HorizontalAlignment = 3      '左右居中   5-填充，左对齐，不会覆盖右边的单元格
        .Rows(1).VerticalAlignment = 2        '上下居中
        '样式-宽高
        .Rows(1).RowHeight = 1/0.035          '高1厘米
        .Columns(1).ColumnWidth = 5           '宽，单位：字符
        .Columns(2).ColumnWidth = 6
        .Columns(3).ColumnWidth = 31
        .Columns(4).ColumnWidth = 41
        .Columns(5).ColumnWidth = 9
        .Columns(6).ColumnWidth = 21
        '样式-四周边框
        .Range("A1","F1").Borders(1).LineStyle = 1
        .Range("A1","F1").Borders(2).LineStyle = 1
        .Range("A1","F1").Borders(3).LineStyle = 1
        .Range("A1","F1").Borders(4).LineStyle = 1
        '样式-其他
        .Rows(1).WrapText = True              '自动换行
        .Range("A1","F1").Interior.Color = D_COLOR_BLUE   '背景色-天蓝色
        .Range("A1","F1").Font.Size = 10                '字体
        .Rows(1).Font.Bold = True             '粗体
    End With


    Dim rowCnt
    rowCnt = 2

    '生成表清单
    output "开始生成表清单..."
    ListObjects mdl,ExcelSheet,rowCnt       '遍历模型

    '样式-设置部分列为左右居中
    With ExcelSheet
        .Columns(1).HorizontalAlignment = 3      '左右居中
        .Columns(2).HorizontalAlignment = 3      '左右居中
        .Columns(5).HorizontalAlignment = 3      '左右居中
    End With

    '调整整个数据区域样式
    Dim rowEnd
    rowEnd = rowCnt-1                '最后一行行号
    With ExcelSheet.Range("A2","F"+Cstr(rowEnd))
        .Borders(1).LineStyle = 1                       '四周边框
        .Borders(2).LineStyle = 1
        .Borders(3).LineStyle = 1
        .Borders(4).LineStyle = 1
    End With
    ExcelSheet.Range("A1","F"+Cstr(rowEnd)).Font.Size = 10       '字体

    '按层名、表名排序
    ExcelApp.AddCustomList Array("ODM", "FDM", "ADM", "MDM", "PUBLIC")
    ExcelSheet.Sort.SortFields.Clear
    ExcelSheet.Sort.SortFields.Add ExcelSheet.Range("B2","B"+Cstr(rowEnd)), 0, 1, "ODM,FDM,ADM,DMD,PUBLIC", 0
    ExcelSheet.Sort.SortFields.Add ExcelSheet.Range("C2","C"+Cstr(rowEnd)), 0, 1, "", 0
    With ExcelSheet.Sort
        .SetRange ExcelSheet.Range("B1","F"+Cstr(rowEnd))
        .Header = 1
        .MatchCase = False
        .Apply
    End With

    '筛选
    ExcelApp.Selection.AutoFilter

    '冻结首行
    ExcelApp.ActiveWindow.SplitRow = 1          '行
    ExcelApp.ActiveWindow.SplitColumn = 0       '列
    ExcelApp.ActiveWindow.FreezePanes = True

    ExcelBook.SaveAs file_path
    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing

    output "表清单生成完毕, 共 " + Cstr(rowCnt-2) + " 张表!"
    Exit Sub
End Sub


'遍历模型
Private Sub ListObjects(fldr,ExcelSheet,rowCnt)
    Dim obj
    For Each obj In fldr.children
        getTables fldr,obj,ExcelSheet,rowCnt
    Next

    Dim f
    For Each f In fldr.Packages
        ListObjects f,ExcelSheet,rowCnt
    Next
End Sub

'获取表清单
Private Sub getTables(CurrentFldr,CurrentObject,ExcelSheet,rowCnt)
    Dim col
    Dim colType
    If CurrentObject.IsKindOf(cls_Table) then
        ExcelSheet.Cells(rowCnt,"A").Value = rowCnt - 1
        If ( CurrentObject.Owner Is Nothing ) Then
            ExcelSheet.Cells(rowCnt,"B").Value = "PUBLIC"
        Else
            ExcelSheet.Cells(rowCnt,"B").Value = CurrentObject.Owner.Code
        End If
        ExcelSheet.Cells(rowCnt,"C").Value = CurrentObject.Code
        ExcelSheet.Cells(rowCnt,"D").Value = CurrentObject.Name
        ExcelSheet.Cells(rowCnt,"E").Value = "Y"
        ExcelSheet.Cells(rowCnt,"F").Value = CurrentObject.comment
        rowCnt = rowCnt + 1
    else
        exit sub
    end if
End Sub



'-------------------------------------------------------------------------------
'根据目录生成表结构，每个表一个Sheet。
'-------------------------------------------------------------------------------
sub createTableSheet(mdl)

    Dim ExcelApp, ExcelBook, ExcelSheet, ExcelMenu
    Dim rowIdx, menuIdx
    Dim tableCnt, colCnt
    Dim tableNum
    Dim tableCode, tableName, tableOwner, tableComment, tableFlag
    tableCnt = 0
    tableNum = 0

    '当用户指定目录文件时，重定义输出文件，以免生成过程中出错，或对输出结果不满意时，需要重新恢复目录文件。
    Dim InputFile, OutputFile
    InputFile = file_path
    If UCase(GEN_MENU) = "N" Then
        OutputFile = Mid(InputFile, 1, InstrRev(InputFile,".")-1) + "_out" + Mid(InputFile, InstrRev(InputFile,"."))
    Else
        OutputFile = InputFile
    End If

    '读取目录文件
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.visible=FALSE
    Set ExcelBook = ExcelApp.Workbooks.Open(file_path)
    Set ExcelMenu = ExcelBook.Sheets("目录")
    menuIdx = ExcelMenu.Index

    For rowIdx = 2 To MAX_TABLES+2
        If ExcelMenu.Cells(rowIdx, "A").Value = "" Then
            Exit For
        Else
            tableNum = tableNum + 1
        End If

        '获取表信息
        tableOwner = ExcelMenu.Cells(rowIdx, "B").Value
        tableCode = ExcelMenu.Cells(rowIdx, COL_TABLE_CODE).Value
        tableName = ExcelMenu.Cells(rowIdx, COL_TABLE_NAME).Value
        tableComment = ExcelMenu.Cells(rowIdx, COL_TABLE_COMMENT).Value
        tableFlag = ExcelMenu.Cells(rowIdx, COL_DEAL_FLAG).Value

        If UCase(tableFlag) = "Y" AND ( Len(tableCode)>0 OR Len(tableName)>0 ) Then     '处理标志非Y则跳过

            '检查表是否存在
            Dim iFlag
            iFlag = 0
            checkTable mdl,ExcelSheet,tableCode,tableName,iFlag

            '表存在则继续处理
            If iFlag = 1 Then

                tableCnt = tableCnt + 1

                '创建Sheet页
                Set ExcelSheet = ExcelBook.Sheets.Add(,ExcelBook.Sheets(menuIdx))       '在目录后面插入，第一个参数为空
                
                'excel sheet名不能超过31个字符，excel会报错
                ExcelSheet.Name = left(tableCode,31)

                output "["+Cstr(tableCnt)+"] "+tableCode

                '添加自定义名称  范围-工作簿
                ExcelBook.Names.Add tableOwner+"."+tableCode,"="+ExcelMenu.Name+"!R"+Cstr(rowIdx)+"C3"       'R=row C=col R2C3=$2$3=C2

                '生成表头
                With ExcelSheet
                    '第一行
                    .Cells(1,"A").Value = "<<返回目录"
                    '超链接，指向自定义名称
                    .Hyperlinks.Add ExcelSheet.Range("A1"),"",tableOwner+"."+tableCode,"",ExcelSheet.Cells(1,"A").Value
                    '超链接，直接定位到单元格，但这样的话，如果目标单元格发生变化，就跳错了。
                    '.Hyperlinks.Add ExcelSheet.Range("A1"),"",ExcelMenu.Name+"!C"+Cstr(rowIdx),"",ExcelSheet.Cells(1,"A").Value

                    '第二行
                    .Cells(2,"A").Value = "英文名"
                    .Range("B2","C2").Merge
                    .Cells(2,"B").Value = tableCode

                    .Cells(2,"D").Value = "模式名"
                    .Cells(2,"E").Value = tableOwner

                    '第三行
                    .Cells(3,"A").Value = "中文名"
                    .Range("B3","E3").Merge
                    .Cells(3,"B").Value = tableName

                    '第四行
                    .Cells(4,"A").Value = "描述"
                    .Range("B4","E4").Merge
                    .Cells(4,"B").Value = tableComment

                    '设置样式-表头
                    .Range("A2","A4").Interior.Color = D_COLOR_GREEN  '背景色-浅绿色
                    .Range("A2","A4").Font.Bold = True              '粗体
                    .Range("A2","A4").HorizontalAlignment = 3       '左右居中

                    .Cells(2,"D").Interior.Color = D_COLOR_GREEN      '背景色-浅绿色
                    .Cells(2,"D").Font.Bold = True                  '粗体
                    .Cells(2,"D").HorizontalAlignment = 3           '左右居中

                    .Range("A1","E4").Font.Size = 10                '字体
                    .Range("A2","E4").Borders(1).LineStyle = 1      '四周边框
                    .Range("A2","E4").Borders(2).LineStyle = 1
                    .Range("A2","E4").Borders(3).LineStyle = 1
                    .Range("A2","E4").Borders(4).LineStyle = 1

                    '第五行-标题栏
                    .Cells(5,"A").Value = "序号"
                    .Cells(5,"B").Value = "字段中文名"
                    .Cells(5,"C").Value = "字段英文名"
                    .Cells(5,"D").Value = "字段类型"
                    .Cells(5,"E").Value = "数据长度"
                    .Cells(5,"F").Value = "主键"
                    .Cells(5,"G").Value = "非空"
                    .Cells(5,"H").Value = "分布键"
                    .Cells(5,"I").Value = "说明"
                    .Cells(5,"J").Value = "备注"

                    '设置样式-第五行-标题栏
                    With .Range("A5","J5")
                        .Interior.Color = D_COLOR_BLUE  '背景色-天蓝色
                        .Font.Bold = True               '粗体
                        .HorizontalAlignment = 3        '左右居中
                        .Font.Size = 10                 '字体
                        .Borders(1).LineStyle = 1       '四周边框
                        .Borders(2).LineStyle = 1
                        .Borders(3).LineStyle = 1
                        .Borders(4).LineStyle = 1
                    End With

                End With

                '生成字段内容
                colCnt=0
                getColumns mdl,ExcelSheet,tableCode,colCnt

                '调整整个数据区域样式
                Dim rowEnd
                rowEnd = colCnt+BEG_ROW-1       '最后一行行号
                With ExcelSheet.Range("A"+Cstr(BEG_ROW),END_COL+Cstr(rowEnd))
                    .Borders(1).LineStyle = 1    '四周边框
                    .Borders(2).LineStyle = 1
                    .Borders(3).LineStyle = 1
                    .Borders(4).LineStyle = 1
                End With
                ExcelSheet.Range("A"+Cstr(BEG_ROW),END_COL+Cstr(rowEnd)).Font.Size = 10              '字体-整个数据区域

                ExcelSheet.Range("A"+Cstr(BEG_ROW),"A"+Cstr(rowEnd)).HorizontalAlignment = 3     '左右居中-序号
                ExcelSheet.Range("F"+Cstr(BEG_ROW),"H"+Cstr(rowEnd)).HorizontalAlignment = 3     '左右居中-主键、非空、分布键

                '创建目录中的超链接
                ExcelMenu.Hyperlinks.Add ExcelMenu.Range(COL_TABLE_CODE+Cstr(rowIdx)),"",ExcelSheet.Name+"!A1","",ExcelSheet.Name
                ExcelMenu.Range(COL_TABLE_CODE+Cstr(rowIdx)).Font.Size = 10
                '更新目录中的表中文名
                ExcelMenu.Range(COL_TABLE_NAME+Cstr(rowIdx)).Value = tableName

                '设置宽度
                With ExcelSheet
                    .Columns("A:H").EntireColumn.AutoFit    '前8列-自适应
                    .Columns(9).ColumnWidth = 30            '说明   宽，单位：字符
                    .Columns(10).ColumnWidth = 10           '备注
                End With

                '拆分冻结单元格
                ExcelApp.ActiveWindow.SplitRow = BEG_ROW-1  '行
                ExcelApp.ActiveWindow.SplitColumn = 5       '列
                ExcelApp.ActiveWindow.FreezePanes = True

                '是否显示分布键
                If UCase(SHOW_DISTRIBUTION_KEYS) <> "Y" Then
                    ExcelSheet.Columns(8).Delete             '删除分布键列
                End If
            End If
        End If
    Next

    '设置目录页为活动页面，效果：打开EXCEL时，首页为目录页面
    ExcelMenu.Activate

    '筛选处理标志为Y的记录
    ExcelMenu.Range("$A$1:$"+COL_DEAL_FLAG+"$"+Cstr(tableNum)).AutoFilter Asc(COL_DEAL_FLAG)-Asc("A")+1,"=Y"

    ExcelBook.SaveAs OutputFile         '另存为输出文件
    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelMenu  = Nothing
    Set ExcelSheet = Nothing
    Set ExcelBook  = Nothing
    Set ExcelApp   = Nothing

    output "输出文件为：[" + OutputFile + "]"
    Exit Sub
End Sub

'检查表是否存在
Sub checkTable(mdl,ExcelSheet,tableCode,tableName,iFlag)
    Dim tb

    If Len(tableCode) > 0 Then
        set tb = mdl.FindChildByCode(tableCode,cls_Table)
        If ( tb Is Nothing ) Then
            output "未找到表[" + tableCode + "]"
            errString = errString + mLF + "未找到表[" + tableCode + "]"
            errCount  = errCount + 1
        Else
            iFlag = 1
            tableName = tb.Name
        End If
    Else
        set tb = mdl.FindChildByName(tableName,cls_Table)
        If ( tb Is Nothing ) Then
            output "未找到表[" + tableName + "]"
            errString = errString + mLF + "未找到表[" + tableName + "]"
            errCount  = errCount + 1
        Else
            iFlag = 1
            tableCode = tb.Code
        End If
    End If

End Sub

'生成字段
Sub getColumns(mdl,ExcelSheet,tableCode,colCnt)

    Dim tb, col, rowIdx
    set tb = mdl.FindChildByCode(tableCode,cls_Table)           '在模型中查找目标表
    If ( tb Is Nothing ) Then
        output "未找到表[" + tableCode + "]"
        errString = errString + mLF + "未找到表[" + tableCode + "]"
        errCount  = errCount + 1
    End If

    Dim colDistributionKeys, dKeys, iKeys, iKeysFlag         '分布键
    Dim tPhysicalOptions, iIdx1, iIdx2, sStr1, sStr2
    iKeysFlag = 0
    If Len(tb.PhysicalOptions) > 0 Then
        tPhysicalOptions = Replace(UCase(tb.PhysicalOptions), mLF, "")      '去换行
        iIdx1 = Instr(tPhysicalOptions, "DISTRIBUTED")                          'DISTRIBUTED在字符串中的位置
        If iIdx1 > 0 Then
            sStr1 = Mid(tPhysicalOptions, iIdx1)                                '从distributed开始的子串
            sStr2 = Mid(sStr1, 1, Instr(sStr1, ")")-1)                          'distributed by (...  没有")"
            colDistributionKeys = Mid(sStr2, Instr(sStr2, "(")+1)               '分布键子串，有多个的话逗号分隔
            dKeys = Split( colDistributionKeys, "," )                           '拆分成数组
            iKeys = ubound(dKeys)                                               '数组最大下标
            iKeysFlag = 1
        End If
    End If

    rowIdx = 5
    For Each col In tb.Columns
        rowIdx = rowIdx + 1
        colCnt = colCnt + 1

        '单元格-中英文表名、数据类型、长度
        ExcelSheet.Cells(rowIdx,"A").Value = colCnt
        ExcelSheet.Cells(rowIdx,"B").Value = col.Name
        ExcelSheet.Cells(rowIdx,"C").Value = col.Code
        ExcelSheet.Cells(rowIdx,"D").Value = col.DataType
        ExcelSheet.Cells(rowIdx,"E").Value = col.Length

        '截取字段类型
        Dim colType, strPair
        If Len(col.DataType) > 0 Then
            strPair = Split( col.DataType, "(" )
            colType = strPair(0)
        Else
            colType = ""
            output "表[" + tableCode + "] 字段["+ col.Name + "] 类型为空！"
            errString = errString + mLF + "表[" + tableCode + "] 字段["+ col.Name + "] 类型为空！"
            errCount  = errCount + 1
        End If

        '根据字段类型，获取数据长度，CHAR类型的PDM自带长度，不需另外处理
        If UCase(colType) = "DATE" Then
            ExcelSheet.Cells(rowIdx,"E").Value = DATA_TYPE_DATE_LEN
        End If
        If UCase(colType) = "TIMESTAMP" Then
            ExcelSheet.Cells(rowIdx,"E").Value = DATA_TYPE_TIMESTAMP_LEN
        End If
        If UCase(colType) = "INTEGER" Then
            ExcelSheet.Cells(rowIdx,"E").Value = DATA_TYPE_INTEGER_LEN
        End If
        If UCase(colType) = "DECIMAL" Or UCase(colType) = "NUMERIC" Then '20150728 新增NUMERIC判断
            Dim str1, str2, colLen
            str1 = Split( strPair(1), ")" )     '截取括号内的值，如15,2或8
            str2 = Split( str1(0), "," )        '截取总长度
            colLen = str2(0)
            ExcelSheet.Cells(rowIdx,"E").Value = Cint(colLen)+2
        End If

        '单元格-主键
        If col.Primary = true Then
            ExcelSheet.Cells(rowIdx,"F").Value = "Y"
        End If

        '单元格-非空
        If col.Mandatory = true Then
            ExcelSheet.Cells(rowIdx,"G").Value = "Y"
        End If

        '单元格-分布键
        If iKeysFlag = 1 Then
            Dim keyIdx
            For keyIdx = 0 To iKeys
                If col.Code = Trim(dKeys(keyIdx)) Then
                    ExcelSheet.Cells(rowIdx,"H").Value = "Y"
                    Exit For
                End If
            Next
        End If

        '单元格-说明
        ExcelSheet.Cells(rowIdx,"I").Value = col.Comment
    Next

    Exit Sub
End Sub