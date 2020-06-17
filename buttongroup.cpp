#include "buttongroup.h"

ButtonGroup::ButtonGroup(QWidget *parent) :
    QGroupBox(parent),
    qButtonGroup(new QButtonGroup(this)),
    _id(-1)
{
    qButtonGroup->setExclusive(false);
    connect(qButtonGroup,SIGNAL(buttonClicked(int)),SIGNAL(buttonClicked(int)));
    connect(qButtonGroup,SIGNAL(buttonClicked(QAbstractButton*)),SIGNAL(buttonClicked(QAbstractButton*)));
    connect(qButtonGroup,SIGNAL(buttonPressed(int)),SIGNAL(buttonPressed(int)));
    connect(qButtonGroup,SIGNAL(buttonPressed(QAbstractButton*)),SIGNAL(buttonPressed(QAbstractButton*)));
    connect(qButtonGroup,SIGNAL(buttonReleased(int)),SIGNAL(buttonReleased(int)));
    connect(qButtonGroup,SIGNAL(buttonReleased(QAbstractButton*)),SIGNAL(buttonReleased(QAbstractButton*)));
    connect(qButtonGroup,SIGNAL(buttonToggled(int,bool)),SIGNAL(buttonToggled(int,bool)));
    connect(qButtonGroup,SIGNAL(buttonToggled(QAbstractButton*,bool)),SIGNAL(buttonToggled(QAbstractButton*,bool)));
}

void ButtonGroup::setExclusive(bool value)
{
    qButtonGroup->setExclusive(value);
}

bool ButtonGroup::exclusive() const
{
    return qButtonGroup->exclusive();
}

void ButtonGroup::addButton(QAbstractButton *button, int id)
{
    qButtonGroup->addButton(button,id);

}

void ButtonGroup::removeButton(QAbstractButton *button)
{
    qButtonGroup->removeButton(button);
}

QList<QAbstractButton *> ButtonGroup::buttons() const
{
    return qButtonGroup->buttons();
}

QAbstractButton *ButtonGroup::checkedButton() const
{
    return qButtonGroup->checkedButton();
}

QAbstractButton *ButtonGroup::button(int id) const
{
    return qButtonGroup->button(id);
}

void ButtonGroup::setId(QAbstractButton *button, int id)
{
    qButtonGroup->setId(button,id);
}

int ButtonGroup::id(QAbstractButton *button) const
{
    return qButtonGroup->id(button);
}

int ButtonGroup::checkedId() const
{
    return qButtonGroup->checkedId();
}

void ButtonGroup::childEvent(QChildEvent *event)
{
    if(event->type() ==  QEvent::ChildAdded)
    {
        QWidget *flat = dynamic_cast<QWidget*>(event->child());
        if(flat != NULL)
        {
            _id += 1;
            addButton(static_cast<QAbstractButton*>(flat),_id);
        }
    }
}

void ButtonGroup::keyPressEvent(QKeyEvent *event)
{

    QKeyEvent *key_event = static_cast<QKeyEvent* >(event);
    QAbstractButton *FocusButton = dynamic_cast<QAbstractButton*>(QWidget::focusWidget());
    if(FocusButton == NULL)
    {
        QGroupBox::keyPressEvent(event);
        return;
    }
    FocusId = qButtonGroup->id(FocusButton);
    if(key_event->key() == Qt::Key_Down)
    {
        if(FocusId < _id)
        {
            qButtonGroup->button(FocusId+1)->setFocus();
            QAbstractButton *FocusButton = dynamic_cast<QAbstractButton*>(QWidget::focusWidget());
            int new_FocusId = qButtonGroup->id(FocusButton);
            if(new_FocusId == FocusId)
            {
                focusNextChild();
            }

        }
        else
        {
            focusNextChild();
        }
    }
    if(key_event->key() == Qt::Key_Up)
    {
         if(FocusId > 0)
         {
             qButtonGroup->button(FocusId-1)->setFocus();
         }
         else
         {
             focusPreviousChild();
         }
    }
}
