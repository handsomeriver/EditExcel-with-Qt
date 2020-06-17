#include "check_system.h"
#include "ui_check_system.h"
#include <QDebug>

CheckSystem::CheckSystem(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::CheckSystem)
{
    ui->setupUi(this);
    Init();
}

CheckSystem::~CheckSystem()
{
    excel->Close();
    delete ui;
}
//遍历文件夹的文件名
QStringList CheckSystem::getFileNames(const QString &path)
{
    QDir dir(path);
    QStringList nameFilters;
    nameFilters << "*.wav";
    QStringList files = dir.entryList(nameFilters, QDir::Files|QDir::Readable, QDir::Name);
    return files;
}

//写入excel，num为行数
void CheckSystem::Editrow(int num)
{
    name = ui->name_lineEdit->text();
    if(name == "")
        name = "X";
    int i = sheet_names.indexOf(name);
    if(i != -1)
        excel->ActiveSheet(name);
    else
    {
        sheet_names.append(name);
        excel->AddSheet(name);
        excel->ActiveSheet(name);
        excel->SetCellValue(1, 1, "文档编号");
        excel->SetCellValue(1, 2, "正常");
        excel->SetCellValue(1, 3, "干啰");
        excel->SetCellValue(1, 4, "湿啰");
        excel->SetCellValue(1, 5, "干啰湿啰");
        excel->SetCellValue(1, 6, "备注");
        excel->SetCellValue(1, 7, "打标人");
        excel->GetRange("A1:G1")->Font("宋体", true, false, 12, QColor(0, 0, 0))->ColumnWidth(9)->RowHeight(29)->Hcenter()->Vcenter();
        excel->GetRange("F1:F1")->Font("宋体", true, false, 12, QColor(0, 0, 0))->ColumnWidth(16)->RowHeight(29)->Hcenter()->Vcenter();
    }

    int the_num = num + 2;
    QString file_id = file_list[current_num];
    file_id = file_id.left(6);
    QString ps;
    QString name;
    ps = ui->ps_textEdit->toPlainText();
    name = ui->name_lineEdit->text();
    excel->SetCellValue(the_num,1,file_id);
    excel->SetCellValue(the_num,2,0);
    excel->SetCellValue(the_num,3,0);
    excel->SetCellValue(the_num,4,0);
    excel->SetCellValue(the_num,5,0);
    excel->SetCellValue(the_num,6,ps);
    excel->SetCellValue(the_num,7,name);
    excel->SetCellValue(the_num,btn_flag+2,1);
}

//初始化excel和界面
void CheckSystem::Init()
{
    effect = new QSoundEffect(this);
    excel = new ExcelHelper(this);
    timer = new QTimer(this);


    QString excel_file_path = QDir::currentPath()+"/check.xlsx";//excel表相对路径
    excel_file_path = QDir::toNativeSeparators(excel_file_path);//转成绝对路径
    QString num_file_path = QDir::currentPath()+"/num.txt";//记录当前音频的txt相对路径
    num_file_path = QDir::toNativeSeparators(num_file_path);//转成绝对路径
    qDebug() << excel_file_path << num_file_path;
    num_file = num_file_path;
    excel->Open(excel_file_path,false);
    file_list = getFileNames("./wav/");//音频路径

    sheet_names = excel->GetSheetNames();
    excel->SetCellValue(1, 1, "文档编号");
    excel->SetCellValue(1, 2, "正常");
    excel->SetCellValue(1, 3, "干啰");
    excel->SetCellValue(1, 4, "湿啰");
    excel->SetCellValue(1, 5, "干啰湿啰");
    excel->SetCellValue(1, 6, "备注");
    excel->SetCellValue(1, 7, "打标人");
    excel->GetRange("A1:G1")->Font("宋体", true, false, 12, QColor(0, 0, 0))->ColumnWidth(9)->RowHeight(29)->Hcenter()->Vcenter();
    excel->GetRange("F1:F1")->Font("宋体", true, false, 12, QColor(0, 0, 0))->ColumnWidth(16)->RowHeight(29)->Hcenter()->Vcenter();
    QFile file(num_file);



    timer->start(500);
    connect(timer,&QTimer::timeout,this,&CheckSystem::FinishWav);
    current_num = 0;
    current_file = "./wav/" + file_list[current_num];
    play_flag = false;
    btn_flag = 0;
    if(!file.open(QIODevice::ReadWrite | QIODevice::Text)) {
        qDebug()<<"Can't open the file!"<<endl;
    }
    else {
           QTextStream text(&file);
           QString str = text.readLine();
           current_num = str.toInt();
           qDebug()<<"num:"<<current_num<< endl;
           file.close();
           QString file_id = file_list[current_num];
           file_id = file_id.left(6);
           ui->id_label->setText(file_id);
           QVariant var = excel->GetCellValue(current_num+2,6);
           ui->ps_textEdit->setText(var.toString());
    }
}


void CheckSystem::FinishWav()
{
    if(!effect->isPlaying())
    {
        effect->stop();
        play_flag = false;
        ui->play_btn->setText("播放");
    }
}

void CheckSystem::on_next_btn_clicked()
{
    name = ui->name_lineEdit->text();
    if(play_flag)
    {
        effect->stop();
        play_flag = false;
        ui->play_btn->setText("播放");
    }
    btn_flag = ui->buttonGroup->checkedId();
    qDebug() << "id:" << btn_flag;
    Editrow(current_num);
    if(current_num < (file_list.size()-1))
    {
        current_num += 1;
        current_file = "./wav/" + file_list[current_num];
        qDebug() << current_file;
        QString file_id = file_list[current_num];
        file_id = file_id.left(6);
        ui->id_label->setText(file_id);
        QVariant var = excel->GetCellValue(current_num+2,6);
        ui->ps_textEdit->setText(var.toString());
    }
    QFile file(num_file);
    if(!file.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        qDebug()<<"Can't open the file!"<<endl;
    }
    else {
           QTextStream text(&file);
           text.seek(file.size());
           text << current_num << "\n";
           file.close();
    }
}

void CheckSystem::on_play_btn_clicked()
{
    effect->setSource(QUrl::fromLocalFile(current_file));
    if(play_flag)
    {
        ui->play_btn->setText("播放");
        effect->stop();
        play_flag = false;
    }
    else
    {
        qDebug() << current_file;
        effect->play();
        play_flag = true;
        ui->play_btn->setText("停止");
    }
}

void CheckSystem::on_pre_btn_clicked()
{
    name = ui->name_lineEdit->text();
    if(play_flag)
    {
        effect->stop();
        play_flag = false;
        ui->play_btn->setText("播放");
    }
    btn_flag = ui->buttonGroup->checkedId();
    Editrow(current_num);
    qDebug() << "id:" << btn_flag;
    if(current_num > 0)
    {
        current_num -= 1;
        current_file = "./wav/" + file_list[current_num];
        qDebug() << current_file;
        QString file_id = file_list[current_num];
        file_id = file_id.left(6);
        ui->id_label->setText(file_id);
        QVariant var = excel->GetCellValue(current_num+2,6);
        ui->ps_textEdit->setText(var.toString());
    }
    QFile file(num_file);
    if(!file.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        qDebug()<<"Can't open the file!"<<endl;
    }
    else {
           QTextStream text(&file);
           text.seek(file.size());
           text << current_num << "\n";
           file.close();
    }
}

void CheckSystem::on_head_btn_clicked()
{
    name = ui->name_lineEdit->text();
    if(play_flag)
    {
        effect->stop();
        play_flag = false;
        ui->play_btn->setText("播放");
    }
    btn_flag = ui->buttonGroup->checkedId();
    Editrow(current_num);
    qDebug() << "id:" << btn_flag;
    current_num = 0;
    current_file = "./wav/" + file_list[current_num];
    QString file_id = file_list[current_num];
    file_id = file_id.left(6);
    ui->id_label->setText(file_id);
    QVariant var = excel->GetCellValue(current_num+2,6);
    ui->ps_textEdit->setText(var.toString());

    QFile file(num_file);
    if(!file.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        qDebug()<<"Can't open the file!"<<endl;
    }
    else {
           QTextStream text(&file);
           text.seek(file.size());
           text << current_num << "\n";
           file.close();
    }
}
