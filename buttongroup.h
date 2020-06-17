#ifndef BUTTONGROUP_H
#define BUTTONGROUP_H

#include <QButtonGroup>
#include <QAbstractButton>
#include <QChildEvent>
#include <QKeyEvent>
#include <QGroupBox>

class ButtonGroup : public QGroupBox
{
    Q_OBJECT

    Q_PROPERTY(bool exclusive READ exclusive WRITE setExclusive)
    //添加属性exclusive，读方法exclusive()，写方法setExclusive(bool)，
    //注册属性，使designer属性窗可以配置
public:
    explicit ButtonGroup(QWidget *parent = 0);

    void setExclusive(bool);
    bool exclusive() const;

    //将QButtonGroup的接口再封装一遍
    void addButton(QAbstractButton *, int id = -1);
    void removeButton(QAbstractButton *);

    QList<QAbstractButton*> buttons() const;

    QAbstractButton * checkedButton() const;

    QAbstractButton *button(int id) const;
    void setId(QAbstractButton *button, int id);
    int id(QAbstractButton *button) const;
    int checkedId() const;

signals:
    void buttonClicked(QAbstractButton *);
    void buttonClicked(int);
    void buttonPressed(QAbstractButton *);
    void buttonPressed(int);
    void buttonReleased(QAbstractButton *);
    void buttonReleased(int);
    void buttonToggled(QAbstractButton *, bool);
    void buttonToggled(int, bool);

protected:
    virtual void childEvent(QChildEvent *event);
    virtual void keyPressEvent(QKeyEvent *event);
private:
    QButtonGroup *qButtonGroup;
    int _id;
    int FocusId;
};

#endif
