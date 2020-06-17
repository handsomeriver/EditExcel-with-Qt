#include <QDebug>
#include "excel_helper.h"
ExcelHelper::ExcelHelper(QObject* parent) :QObject(parent)
{
    m_pApplication = new QAxObject();
    m_pApplication->setControl("Excel.Application"); //连接Excel控件
}
//打开excel
void ExcelHelper::Open(const QString &fileName, bool bVisible)
{
    QFile file(fileName);
    if (file.exists())
    {
        m_pApplication->dynamicCall("SetVisible(bool)", bVisible); //是否显示窗体
        m_pApplication->setProperty("DisplayAlerts", false); //不显示任何警告信息。
        m_pWorkBooks = m_pApplication->querySubObject("Workbooks");
        m_pWorkBook = m_pWorkBooks->querySubObject("Open(const QString &)", fileName);
        //默认有一个sheet
        m_pSheets = m_pWorkBook->querySubObject("Sheets");
        m_pActiveSheet = m_pSheets->querySubObject("Item(int)", 1);
    }
    else
    {
        qDebug() << "打开失败！";
    }
}
//新建Execl文件
void ExcelHelper::New(bool bVisible)
{
    m_pApplication->dynamicCall("SetVisible(bool)", bVisible); //是否显示窗体
    m_pApplication->setProperty("DisplayAlerts", false); //不显示任何警告信息。
    m_pWorkBooks = m_pApplication->querySubObject("Workbooks");
    m_pWorkBooks->dynamicCall("Add");
    m_pWorkBook = m_pApplication->querySubObject("ActiveWorkBook");
    //默认有一个sheet
    m_pSheets = m_pWorkBook->querySubObject("Sheets");
    m_pActiveSheet = m_pSheets->querySubObject("Item(int)", 1);
}

//指定操作的sheet
void ExcelHelper::ActiveSheet(int nItem)
{
    m_pActiveSheet = m_pSheets->querySubObject("Item(int)", nItem);
    m_pActiveSheet->dynamicCall("Select");
}
void ExcelHelper::ActiveSheet(const QString &sheetName)
{
    m_pActiveSheet = m_pWorkBook->querySubObject("Sheets(string)", sheetName);
    m_pActiveSheet->dynamicCall("Select");
}

//添加Sheet
void ExcelHelper::AddSheet(const QString &sheetName)
{
    int cnt = 1;
    QAxObject *pLastSheet = m_pSheets->querySubObject("Item(int)", cnt);
    m_pSheets->querySubObject("Add(QVariant)", pLastSheet->asVariant());
    m_pActiveSheet = m_pSheets->querySubObject("Item(int)", cnt);
    pLastSheet->dynamicCall("Move(QVariant)", m_pActiveSheet->asVariant());
    m_pActiveSheet->setProperty("Name", sheetName);
}

//获取所有表名集合
QStringList ExcelHelper::GetSheetNames()
{
    int sheet_count = m_pSheets->property("Count").toInt();  //获取工作表数目
    QStringList sheet_names;
    for(int i=1; i<=sheet_count; i++)
    {
        QAxObject *work_sheet = m_pWorkBook->querySubObject("Sheets(int)", i);  //Sheets(int)也可换用Worksheets(int)
        QString work_sheet_name = work_sheet->property("Name").toString();  //获取工作表名称
        sheet_names.append(work_sheet_name);
    }
    return sheet_names;
}

//修改指定单元格的值
void ExcelHelper::SetCellValue(int row, int column, const QVariant &value)
{
    QAxObject *pRange = m_pActiveSheet->querySubObject("Cells(int,int)", row, column);
    pRange->dynamicCall("Value", value);
}

//获取指定单元格的值
QVariant ExcelHelper::GetCellValue(int row, int column)
{
    QAxObject *pRange = m_pActiveSheet->querySubObject("Cells(int,int)", row, column);
    QVariant var = pRange->dynamicCall("Value");
    return var;
}
Range* ExcelHelper::GetRange(int row, int column)
{
    QAxObject *pRange = m_pActiveSheet->querySubObject("Cells(int,int)", row, column);
    return new Range(pRange);
}
Range* ExcelHelper::GetRange(QString str)
{
    QAxObject *pRange = m_pActiveSheet->querySubObject("Range(string)",str);
    return new Range(pRange,this);
}
//保存Excel
void ExcelHelper::Save(const QString &fileName)
{
    m_pWorkBook->dynamicCall("SaveAs(string)", fileName);
}

//关闭Excel
void ExcelHelper::Close()
{
    if (m_pApplication != NULL) {
        m_pWorkBook->dynamicCall("Save()");
        m_pApplication->dynamicCall("Quit()");
        delete m_pApplication;
        m_pApplication = NULL;
    }
}
