Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim AppXls As Microsoft.Office.Interop.Excel.Application    '声明Excel对象
        Dim AppWokBook As Microsoft.Office.Interop.Excel.Workbook    '声明工作簿对象
        Dim AppSheet As New Microsoft.Office.Interop.Excel.Worksheet    '声明工作表对象

        AppXls = New Microsoft.Office.Interop.Excel.Application     '实例化Excel对象
        AppXls.Workbooks.Open("C:\学生成绩.xls")                    '打开已经存在的EXCEL文件
        AppXls.Visible = False                                      '使Excel不可见

        'AppWokBook = New Microsoft.Office.Interop.Excel.Workbook    '实例化工作簿对象
        'AppSheet = New Microsoft.Office.Interop.Excel.Worksheet     '实例化工作表对象

        AppWokBook = AppXls.Workbooks(1)            'AppWokBook对象指向工作簿"C:\学生成绩.xls"
        AppSheet = AppWokBook.Sheets("Sheet1")                      'AppSheet对象指向AppWokBook对象中的表“Sheet1”，即："C:\学生成绩.xls"中的表“Sheet1”

        '下面举一些例子：
        '1、如果不声明工作表对象 AppSheet ，那么应用AppWokBook对象中的表“Sheet1”的语句就是：AppWokBook.Sheets("Sheet1")
        '2、如果不声明工作簿对象 AppWokBook ，那么应用"C:\学生成绩.xls"中的表“Sheet1”的语句就是：AppXls.Workbooks("C:\学生成绩.xls").Sheets("Sheet1")

        '要读取数据表"Sheet1"中的单元格“A1”的值，到变量S1里
        Dim S1 As String
        '方法一
        S1 = AppXls.Workbooks(1).Sheets("Sheet1").Range("A1").Value
        MsgBox(S1)

        '方法二
        S1 = AppWokBook.Sheets("Sheet1").Range("A1").Value
        MsgBox(S1)

        '方法三
        S1 = AppSheet.Range("A1").Value
        MsgBox(S1)

        '把数据写入到单元格“H2”，就是第2行第8个单元格
        '方法一
        AppXls.Workbooks(1).Sheets("Sheet1").Cells(2, 8).Value = "您好！"
        S1 = AppXls.Workbooks(1).Sheets("Sheet1").Cells(2, 8).Value        '为了验证，读取并显示它
        MsgBox(S1)

        '方法二
        AppWokBook.Sheets("Sheet1").Cells(2, 8).Value = "你们好！"
        S1 = AppWokBook.Sheets("Sheet1").Cells(2, 8).Value                 '为了验证，读取并显示它
        MsgBox(S1)

        '方法二
        AppSheet.Cells(2, 8).Value = "大家好！"
        S1 = AppSheet.Cells(2, 8).Value                                   '为了验证，读取并显示它
        MsgBox(S1)

        '使用完毕必须关闭EXCEL，并退出
        AppXls.ActiveWorkbook.Close(SaveChanges:=True)
        AppXls.Quit()