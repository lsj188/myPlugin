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
'----------------------------------�밴������-----------------------------------
CONST GEN_MENU    = "Y"                         '�Ƿ�����Ŀ¼�ļ� [ Y-�� N-�� ]
CONST GEN_TABLE   = "Y"                         '�Ƿ�����ģ�ͽṹ [ Y-�� N-�� ]
CONST SHOW_DISTRIBUTION_KEYS  = "Y"             '�Ƿ���ʾ�ֲ���   [ Y-�� N-�� ]
'----------------------------------Ŀ¼ҳ����-----------------------------------
CONST COL_TABLE_CODE = "C"                      '��Ӣ������
CONST COL_TABLE_NAME = "D"                      '����������
CONST COL_DEAL_FLAG  = "E"                      '�����־��
CONST COL_TABLE_COMMENT  = "F"                  '����������
'-------------------------------------------------------------------------------
CONST BEG_ROW = 6                               '��������-��ʼ��
CONST END_COL = "J"                             '��������-������
CONST MAX_TABLES = 1000                         '����������

CONST DATA_TYPE_DATE_LEN      = 10              'DATE�������ݳ���
CONST DATA_TYPE_TIMESTAMP_LEN = 19              'TIMESTAMP�������ݳ���
CONST DATA_TYPE_INTEGER_LEN   = 12              'INTEGER�������ݳ���

CONST D_COLOR_BLUE     = 16764057               '����ɫ
CONST D_COLOR_GREEN    = 13434828               'ǳ��ɫ
CONST D_COLOR_ORAGNE   = 49407                  '��ɫ

Dim mCR,mLF
mCR = Chr(10)       '����
mLF = Chr(13)       '�س�
'-------------------------------------------------------------------------------

'����PDM
Dim mdl
Dim file_path
Dim errCount, errString
errCount=0
Set mdl = ActiveModel
If ( mdl Is Nothing ) Then
    MsgBox "There is no Active Model"
Else
    file_path=InputBox("�����뵼���ļ�·��")
    If UCase(GEN_MENU) = "Y" Then
        createMenuSheet mdl         '����Ŀ¼
    End If

    If UCase(GEN_TABLE) = "Y" Then
        createTableSheet mdl        '����Ŀ¼���ɱ�ṹ
    End If

    If errCount > 0 Then
        output "������Ϣ: " + errString
    End If
    MsgBox "�������,����"+Cstr(errCount)+"������!"
End If

'-------------------------------------------------------------------------------
'����Ŀ¼
'   ���|ģʽ��|����|�����־(Y/N)|���ı���|��ע
'   �����־Ĭ��ȫ��ΪY
'-------------------------------------------------------------------------------
sub createMenuSheet(mdl)

    Dim ExcelApp, ExcelBook, ExcelSheet

    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.visible=FALSE
    Set ExcelBook = ExcelApp.Workbooks.Add
    Set ExcelSheet = ExcelBook.Sheets.Add
    ExcelSheet.Name = "Ŀ¼"

    'Ŀ¼������
    With ExcelSheet
        '����
        .Cells(1,"A").Value = "���"
        .Cells(1,"B").Value = "ģʽ��"
        .Cells(1,"C").Value = "��Ӣ����"
        .Cells(1,"D").Value = "���ı���"
        .Cells(1,"E").Value = "�����־(Y/N)"
        .Cells(1,"F").Value = "��ע"

        '��ʽ-����
        .Rows(1).HorizontalAlignment = 3      '���Ҿ���   5-��䣬����룬���Ḳ���ұߵĵ�Ԫ��
        .Rows(1).VerticalAlignment = 2        '���¾���
        '��ʽ-���
        .Rows(1).RowHeight = 1/0.035          '��1����
        .Columns(1).ColumnWidth = 5           '����λ���ַ�
        .Columns(2).ColumnWidth = 6
        .Columns(3).ColumnWidth = 31
        .Columns(4).ColumnWidth = 41
        .Columns(5).ColumnWidth = 9
        .Columns(6).ColumnWidth = 21
        '��ʽ-���ܱ߿�
        .Range("A1","F1").Borders(1).LineStyle = 1
        .Range("A1","F1").Borders(2).LineStyle = 1
        .Range("A1","F1").Borders(3).LineStyle = 1
        .Range("A1","F1").Borders(4).LineStyle = 1
        '��ʽ-����
        .Rows(1).WrapText = True              '�Զ�����
        .Range("A1","F1").Interior.Color = D_COLOR_BLUE   '����ɫ-����ɫ
        .Range("A1","F1").Font.Size = 10                '����
        .Rows(1).Font.Bold = True             '����
    End With


    Dim rowCnt
    rowCnt = 2

    '���ɱ��嵥
    output "��ʼ���ɱ��嵥..."
    ListObjects mdl,ExcelSheet,rowCnt       '����ģ��

    '��ʽ-���ò�����Ϊ���Ҿ���
    With ExcelSheet
        .Columns(1).HorizontalAlignment = 3      '���Ҿ���
        .Columns(2).HorizontalAlignment = 3      '���Ҿ���
        .Columns(5).HorizontalAlignment = 3      '���Ҿ���
    End With

    '������������������ʽ
    Dim rowEnd
    rowEnd = rowCnt-1                '���һ���к�
    With ExcelSheet.Range("A2","F"+Cstr(rowEnd))
        .Borders(1).LineStyle = 1                       '���ܱ߿�
        .Borders(2).LineStyle = 1
        .Borders(3).LineStyle = 1
        .Borders(4).LineStyle = 1
    End With
    ExcelSheet.Range("A1","F"+Cstr(rowEnd)).Font.Size = 10       '����

    '����������������
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

    'ɸѡ
    ExcelApp.Selection.AutoFilter

    '��������
    ExcelApp.ActiveWindow.SplitRow = 1          '��
    ExcelApp.ActiveWindow.SplitColumn = 0       '��
    ExcelApp.ActiveWindow.FreezePanes = True

    ExcelBook.SaveAs file_path
    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing

    output "���嵥�������, �� " + Cstr(rowCnt-2) + " �ű�!"
    Exit Sub
End Sub


'����ģ��
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

'��ȡ���嵥
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
'����Ŀ¼���ɱ�ṹ��ÿ����һ��Sheet��
'-------------------------------------------------------------------------------
sub createTableSheet(mdl)

    Dim ExcelApp, ExcelBook, ExcelSheet, ExcelMenu
    Dim rowIdx, menuIdx
    Dim tableCnt, colCnt
    Dim tableNum
    Dim tableCode, tableName, tableOwner, tableComment, tableFlag
    tableCnt = 0
    tableNum = 0

    '���û�ָ��Ŀ¼�ļ�ʱ���ض�������ļ����������ɹ����г������������������ʱ����Ҫ���»ָ�Ŀ¼�ļ���
    Dim InputFile, OutputFile
    InputFile = file_path
    If UCase(GEN_MENU) = "N" Then
        OutputFile = Mid(InputFile, 1, InstrRev(InputFile,".")-1) + "_out" + Mid(InputFile, InstrRev(InputFile,"."))
    Else
        OutputFile = InputFile
    End If

    '��ȡĿ¼�ļ�
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.visible=FALSE
    Set ExcelBook = ExcelApp.Workbooks.Open(file_path)
    Set ExcelMenu = ExcelBook.Sheets("Ŀ¼")
    menuIdx = ExcelMenu.Index

    For rowIdx = 2 To MAX_TABLES+2
        If ExcelMenu.Cells(rowIdx, "A").Value = "" Then
            Exit For
        Else
            tableNum = tableNum + 1
        End If

        '��ȡ����Ϣ
        tableOwner = ExcelMenu.Cells(rowIdx, "B").Value
        tableCode = ExcelMenu.Cells(rowIdx, COL_TABLE_CODE).Value
        tableName = ExcelMenu.Cells(rowIdx, COL_TABLE_NAME).Value
        tableComment = ExcelMenu.Cells(rowIdx, COL_TABLE_COMMENT).Value
        tableFlag = ExcelMenu.Cells(rowIdx, COL_DEAL_FLAG).Value

        If UCase(tableFlag) = "Y" AND ( Len(tableCode)>0 OR Len(tableName)>0 ) Then     '�����־��Y������

            '�����Ƿ����
            Dim iFlag
            iFlag = 0
            checkTable mdl,ExcelSheet,tableCode,tableName,iFlag

            '��������������
            If iFlag = 1 Then

                tableCnt = tableCnt + 1

                '����Sheetҳ
                Set ExcelSheet = ExcelBook.Sheets.Add(,ExcelBook.Sheets(menuIdx))       '��Ŀ¼������룬��һ������Ϊ��
                
                'excel sheet�����ܳ���31���ַ���excel�ᱨ��
                ExcelSheet.Name = left(tableCode,31)

                output "["+Cstr(tableCnt)+"] "+tableCode

                '����Զ�������  ��Χ-������
                ExcelBook.Names.Add tableOwner+"."+tableCode,"="+ExcelMenu.Name+"!R"+Cstr(rowIdx)+"C3"       'R=row C=col R2C3=$2$3=C2

                '���ɱ�ͷ
                With ExcelSheet
                    '��һ��
                    .Cells(1,"A").Value = "<<����Ŀ¼"
                    '�����ӣ�ָ���Զ�������
                    .Hyperlinks.Add ExcelSheet.Range("A1"),"",tableOwner+"."+tableCode,"",ExcelSheet.Cells(1,"A").Value
                    '�����ӣ�ֱ�Ӷ�λ����Ԫ�񣬵������Ļ������Ŀ�굥Ԫ�����仯���������ˡ�
                    '.Hyperlinks.Add ExcelSheet.Range("A1"),"",ExcelMenu.Name+"!C"+Cstr(rowIdx),"",ExcelSheet.Cells(1,"A").Value

                    '�ڶ���
                    .Cells(2,"A").Value = "Ӣ����"
                    .Range("B2","C2").Merge
                    .Cells(2,"B").Value = tableCode

                    .Cells(2,"D").Value = "ģʽ��"
                    .Cells(2,"E").Value = tableOwner

                    '������
                    .Cells(3,"A").Value = "������"
                    .Range("B3","E3").Merge
                    .Cells(3,"B").Value = tableName

                    '������
                    .Cells(4,"A").Value = "����"
                    .Range("B4","E4").Merge
                    .Cells(4,"B").Value = tableComment

                    '������ʽ-��ͷ
                    .Range("A2","A4").Interior.Color = D_COLOR_GREEN  '����ɫ-ǳ��ɫ
                    .Range("A2","A4").Font.Bold = True              '����
                    .Range("A2","A4").HorizontalAlignment = 3       '���Ҿ���

                    .Cells(2,"D").Interior.Color = D_COLOR_GREEN      '����ɫ-ǳ��ɫ
                    .Cells(2,"D").Font.Bold = True                  '����
                    .Cells(2,"D").HorizontalAlignment = 3           '���Ҿ���

                    .Range("A1","E4").Font.Size = 10                '����
                    .Range("A2","E4").Borders(1).LineStyle = 1      '���ܱ߿�
                    .Range("A2","E4").Borders(2).LineStyle = 1
                    .Range("A2","E4").Borders(3).LineStyle = 1
                    .Range("A2","E4").Borders(4).LineStyle = 1

                    '������-������
                    .Cells(5,"A").Value = "���"
                    .Cells(5,"B").Value = "�ֶ�������"
                    .Cells(5,"C").Value = "�ֶ�Ӣ����"
                    .Cells(5,"D").Value = "�ֶ�����"
                    .Cells(5,"E").Value = "���ݳ���"
                    .Cells(5,"F").Value = "����"
                    .Cells(5,"G").Value = "�ǿ�"
                    .Cells(5,"H").Value = "�ֲ���"
                    .Cells(5,"I").Value = "˵��"
                    .Cells(5,"J").Value = "��ע"

                    '������ʽ-������-������
                    With .Range("A5","J5")
                        .Interior.Color = D_COLOR_BLUE  '����ɫ-����ɫ
                        .Font.Bold = True               '����
                        .HorizontalAlignment = 3        '���Ҿ���
                        .Font.Size = 10                 '����
                        .Borders(1).LineStyle = 1       '���ܱ߿�
                        .Borders(2).LineStyle = 1
                        .Borders(3).LineStyle = 1
                        .Borders(4).LineStyle = 1
                    End With

                End With

                '�����ֶ�����
                colCnt=0
                getColumns mdl,ExcelSheet,tableCode,colCnt

                '������������������ʽ
                Dim rowEnd
                rowEnd = colCnt+BEG_ROW-1       '���һ���к�
                With ExcelSheet.Range("A"+Cstr(BEG_ROW),END_COL+Cstr(rowEnd))
                    .Borders(1).LineStyle = 1    '���ܱ߿�
                    .Borders(2).LineStyle = 1
                    .Borders(3).LineStyle = 1
                    .Borders(4).LineStyle = 1
                End With
                ExcelSheet.Range("A"+Cstr(BEG_ROW),END_COL+Cstr(rowEnd)).Font.Size = 10              '����-������������

                ExcelSheet.Range("A"+Cstr(BEG_ROW),"A"+Cstr(rowEnd)).HorizontalAlignment = 3     '���Ҿ���-���
                ExcelSheet.Range("F"+Cstr(BEG_ROW),"H"+Cstr(rowEnd)).HorizontalAlignment = 3     '���Ҿ���-�������ǿա��ֲ���

                '����Ŀ¼�еĳ�����
                ExcelMenu.Hyperlinks.Add ExcelMenu.Range(COL_TABLE_CODE+Cstr(rowIdx)),"",ExcelSheet.Name+"!A1","",ExcelSheet.Name
                ExcelMenu.Range(COL_TABLE_CODE+Cstr(rowIdx)).Font.Size = 10
                '����Ŀ¼�еı�������
                ExcelMenu.Range(COL_TABLE_NAME+Cstr(rowIdx)).Value = tableName

                '���ÿ��
                With ExcelSheet
                    .Columns("A:H").EntireColumn.AutoFit    'ǰ8��-����Ӧ
                    .Columns(9).ColumnWidth = 30            '˵��   ����λ���ַ�
                    .Columns(10).ColumnWidth = 10           '��ע
                End With

                '��ֶ��ᵥԪ��
                ExcelApp.ActiveWindow.SplitRow = BEG_ROW-1  '��
                ExcelApp.ActiveWindow.SplitColumn = 5       '��
                ExcelApp.ActiveWindow.FreezePanes = True

                '�Ƿ���ʾ�ֲ���
                If UCase(SHOW_DISTRIBUTION_KEYS) <> "Y" Then
                    ExcelSheet.Columns(8).Delete             'ɾ���ֲ�����
                End If
            End If
        End If
    Next

    '����Ŀ¼ҳΪ�ҳ�棬Ч������EXCELʱ����ҳΪĿ¼ҳ��
    ExcelMenu.Activate

    'ɸѡ�����־ΪY�ļ�¼
    ExcelMenu.Range("$A$1:$"+COL_DEAL_FLAG+"$"+Cstr(tableNum)).AutoFilter Asc(COL_DEAL_FLAG)-Asc("A")+1,"=Y"

    ExcelBook.SaveAs OutputFile         '���Ϊ����ļ�
    ExcelBook.Close
    ExcelApp.Quit
    Set ExcelMenu  = Nothing
    Set ExcelSheet = Nothing
    Set ExcelBook  = Nothing
    Set ExcelApp   = Nothing

    output "����ļ�Ϊ��[" + OutputFile + "]"
    Exit Sub
End Sub

'�����Ƿ����
Sub checkTable(mdl,ExcelSheet,tableCode,tableName,iFlag)
    Dim tb

    If Len(tableCode) > 0 Then
        set tb = mdl.FindChildByCode(tableCode,cls_Table)
        If ( tb Is Nothing ) Then
            output "δ�ҵ���[" + tableCode + "]"
            errString = errString + mLF + "δ�ҵ���[" + tableCode + "]"
            errCount  = errCount + 1
        Else
            iFlag = 1
            tableName = tb.Name
        End If
    Else
        set tb = mdl.FindChildByName(tableName,cls_Table)
        If ( tb Is Nothing ) Then
            output "δ�ҵ���[" + tableName + "]"
            errString = errString + mLF + "δ�ҵ���[" + tableName + "]"
            errCount  = errCount + 1
        Else
            iFlag = 1
            tableCode = tb.Code
        End If
    End If

End Sub

'�����ֶ�
Sub getColumns(mdl,ExcelSheet,tableCode,colCnt)

    Dim tb, col, rowIdx
    set tb = mdl.FindChildByCode(tableCode,cls_Table)           '��ģ���в���Ŀ���
    If ( tb Is Nothing ) Then
        output "δ�ҵ���[" + tableCode + "]"
        errString = errString + mLF + "δ�ҵ���[" + tableCode + "]"
        errCount  = errCount + 1
    End If

    Dim colDistributionKeys, dKeys, iKeys, iKeysFlag         '�ֲ���
    Dim tPhysicalOptions, iIdx1, iIdx2, sStr1, sStr2
    iKeysFlag = 0
    If Len(tb.PhysicalOptions) > 0 Then
        tPhysicalOptions = Replace(UCase(tb.PhysicalOptions), mLF, "")      'ȥ����
        iIdx1 = Instr(tPhysicalOptions, "DISTRIBUTED")                          'DISTRIBUTED���ַ����е�λ��
        If iIdx1 > 0 Then
            sStr1 = Mid(tPhysicalOptions, iIdx1)                                '��distributed��ʼ���Ӵ�
            sStr2 = Mid(sStr1, 1, Instr(sStr1, ")")-1)                          'distributed by (...  û��")"
            colDistributionKeys = Mid(sStr2, Instr(sStr2, "(")+1)               '�ֲ����Ӵ����ж���Ļ����ŷָ�
            dKeys = Split( colDistributionKeys, "," )                           '��ֳ�����
            iKeys = ubound(dKeys)                                               '��������±�
            iKeysFlag = 1
        End If
    End If

    rowIdx = 5
    For Each col In tb.Columns
        rowIdx = rowIdx + 1
        colCnt = colCnt + 1

        '��Ԫ��-��Ӣ�ı������������͡�����
        ExcelSheet.Cells(rowIdx,"A").Value = colCnt
        ExcelSheet.Cells(rowIdx,"B").Value = col.Name
        ExcelSheet.Cells(rowIdx,"C").Value = col.Code
        ExcelSheet.Cells(rowIdx,"D").Value = col.DataType
        ExcelSheet.Cells(rowIdx,"E").Value = col.Length

        '��ȡ�ֶ�����
        Dim colType, strPair
        If Len(col.DataType) > 0 Then
            strPair = Split( col.DataType, "(" )
            colType = strPair(0)
        Else
            colType = ""
            output "��[" + tableCode + "] �ֶ�["+ col.Name + "] ����Ϊ�գ�"
            errString = errString + mLF + "��[" + tableCode + "] �ֶ�["+ col.Name + "] ����Ϊ�գ�"
            errCount  = errCount + 1
        End If

        '�����ֶ����ͣ���ȡ���ݳ��ȣ�CHAR���͵�PDM�Դ����ȣ��������⴦��
        If UCase(colType) = "DATE" Then
            ExcelSheet.Cells(rowIdx,"E").Value = DATA_TYPE_DATE_LEN
        End If
        If UCase(colType) = "TIMESTAMP" Then
            ExcelSheet.Cells(rowIdx,"E").Value = DATA_TYPE_TIMESTAMP_LEN
        End If
        If UCase(colType) = "INTEGER" Then
            ExcelSheet.Cells(rowIdx,"E").Value = DATA_TYPE_INTEGER_LEN
        End If
        If UCase(colType) = "DECIMAL" Or UCase(colType) = "NUMERIC" Then '20150728 ����NUMERIC�ж�
            Dim str1, str2, colLen
            str1 = Split( strPair(1), ")" )     '��ȡ�����ڵ�ֵ����15,2��8
            str2 = Split( str1(0), "," )        '��ȡ�ܳ���
            colLen = str2(0)
            ExcelSheet.Cells(rowIdx,"E").Value = Cint(colLen)+2
        End If

        '��Ԫ��-����
        If col.Primary = true Then
            ExcelSheet.Cells(rowIdx,"F").Value = "Y"
        End If

        '��Ԫ��-�ǿ�
        If col.Mandatory = true Then
            ExcelSheet.Cells(rowIdx,"G").Value = "Y"
        End If

        '��Ԫ��-�ֲ���
        If iKeysFlag = 1 Then
            Dim keyIdx
            For keyIdx = 0 To iKeys
                If col.Code = Trim(dKeys(keyIdx)) Then
                    ExcelSheet.Cells(rowIdx,"H").Value = "Y"
                    Exit For
                End If
            Next
        End If

        '��Ԫ��-˵��
        ExcelSheet.Cells(rowIdx,"I").Value = col.Comment
    Next

    Exit Sub
End Sub