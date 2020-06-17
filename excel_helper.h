#ifndef EXCELHANDLE_H
#define EXCELHANDLE_H

#include <QAxObject>
#include <QDir>
#include <QFile>
#include <QColor>
#include <QMessageBox>

class Range:QObject
{
public:
    Range(QAxObject* p,QObject* parent=nullptr):QObject(parent) { m_pRange = p; }
    ~Range()
    {
    }
public:
    //垂直居中
    Range* Vcenter()
    {
        m_pRange->setProperty("VerticalAlignment", -4108);
        return this;
    }
    //水平居中
    Range* Hcenter()
    {
        m_pRange->setProperty("HorizontalAlignment", -4108);
        return this;
    }
    //行高
    Range* RowHeight(int nHeight)
    {
        m_pRange->setProperty("RowHeight", nHeight);
        return this;
    }
    //列宽
    Range* ColumnWidth(int nWidth)
    {
        m_pRange->setProperty("ColumnWidth", nWidth);
        return this;
    }
    //自动换行
    Range* AutoWrapText()
    {
        m_pRange->setProperty("WrapText", true); //内容过多，自动换行
        return this;
    }
    //背景色
    Range* BackgroundColor(QColor crBg)
    {
        QAxObject* interior = m_pRange->querySubObject("Interior");
        interior->setProperty("Color", crBg);
        return this;
    }
    Range* BorderColor(QColor crBorder)
    {
        QAxObject* border = m_pRange->querySubObject("Borders");
        border->setProperty("Color", crBorder); //设置单元格边框色（蓝色）
        /*
        .LineStyle = xlContinuous   border->setProperty("LineStyle", 4);,下面类同
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin*/
        return this;
    }
    //设置值
    Range* Value(QVariant v)
    {
        m_pRange->dynamicCall("Value", v);
        return this;
    }

    //设置字体
    Range* Font(QString strFaceName, bool bBold, bool bItalic, int nSize, QColor crText)
    {
        QAxObject *font = m_pRange->querySubObject("Font");  //获取单元格字体
        font->setProperty("Name", strFaceName);  //设置单元格字体font->setProperty("Name", QStringLiteral("华文彩云"));  //设置单元格字体
        font->setProperty("Bold", bBold);  //设置单元格字体加粗
        font->setProperty("Size", nSize);  //设置单元格字体大小
        font->setProperty("Italic", bItalic);  //设置单元格字体斜体
        //font->setProperty("Underline", 2);  //设置单元格下划线
        font->setProperty("Color", crText);  //设置单元格字体颜色（红色）
        return this;
    }
private:
    QAxObject* m_pRange;
};

class ExcelHelper:public QObject
{
public:
    ExcelHelper(QObject* parent = nullptr);
    void Open(const QString &fileName, bool bVisible=true);
    void New(bool bVisible=true);
    void ActiveSheet(int nItem);                                                //激活sheet
    void ActiveSheet(const QString &sheetName);                                    //激活sheet
    void AddSheet(const QString &sheetName);
    int GetRows();
    int GetColumns();
    QStringList GetSheetNames();

    Range* GetRange(int row, int column);//获得一个单元格区域
    Range* GetRange(QString str);//获得一个区域,如A3:B18或A1
    void SetCellValue(int row, int column,const QVariant &value);
    QVariant GetCellValue(int row, int column);
    void Save(const QString &fileName);
    void Close();

    QAxObject    *m_pApplication;
    QAxObject    *m_pWorkBooks;
    QAxObject    *m_pWorkBook;
    QAxObject    *m_pSheets;
    QAxObject    *m_pActiveSheet;
};
#endif // EXCELHANDLE_H
