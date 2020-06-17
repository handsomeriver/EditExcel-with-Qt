#include "check_system.h"
#include "excel_helper.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
//    ExcelHelper excel;
//    excel.New(true);
//    excel.ActiveSheet("sheet1");
//    excel.SetCellValue(1, 1, "文档编号");
//    excel.SetCellValue(1, 2, "正常");
//    excel.SetCellValue(1, 3, "干啰");
//    excel.SetCellValue(1, 4, "湿啰");
//    excel.SetCellValue(1, 5, "干啰湿啰");
//    excel.SetCellValue(1, 6, "备注");
//    excel.GetRange("A1:E1")->Font("宋体", true, false, 12, QColor(0, 0, 0))->ColumnWidth(9)->RowHeight(29)->Hcenter()->Vcenter();
//    excel.GetRange("F1:F1")->Font("宋体", true, false, 12, QColor(0, 0, 0))->ColumnWidth(16)->RowHeight(29)->Hcenter()->Vcenter();
//    excel.Save("D:\\aaaaa.xlsx");
//    excel.Close();
    CheckSystem w;
    w.show();

    return a.exec();
}
