#ifndef CHECK_SYSTEM_H
#define CHECK_SYSTEM_H

#include <QMainWindow>
#include <QDir>
#include <QFile>
#include <QSoundEffect>
#include <QTimer>
#include "excel_helper.h"


namespace Ui {
class CheckSystem;
}

class CheckSystem : public QMainWindow
{
    Q_OBJECT

public:
    explicit CheckSystem(QWidget *parent = 0);
    ~CheckSystem();

private slots:
    void on_next_btn_clicked();

    void on_play_btn_clicked();

    void on_pre_btn_clicked();

    void on_head_btn_clicked();

private:
    Ui::CheckSystem *ui;
    QStringList getFileNames(const QString &path);
    QSoundEffect *effect;
    QStringList file_list;
    QString current_file;
    QString num_file;
    QString name;
    ExcelHelper *excel;
    QTimer *timer;
    QStringList sheet_names;
    void Editrow(int num);
    void Init();
    void FinishWav();
    int current_num;
    int btn_flag;
    bool play_flag;
};

#endif // CHECK_SYSTEM_H
