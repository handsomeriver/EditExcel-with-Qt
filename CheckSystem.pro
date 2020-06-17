#-------------------------------------------------
#
# Project created by QtCreator 2020-06-16T14:02:57
#
#-------------------------------------------------

QT       += core gui
QT       += multimedia
QT       += axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = CheckSystem
TEMPLATE = app


SOURCES += main.cpp\
        check_system.cpp \
    buttongroup.cpp \
    excel_helper.cpp

HEADERS  += check_system.h \
    buttongroup.h \
    excel_helper.h

FORMS    += check_system.ui

RESOURCES +=
