#-------------------------------------------------
#
# Project created by QtCreator 2018-02-16T19:09:16
#
#-------------------------------------------------

QT       += core gui
CONFIG += qaxcontainer c++14

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = JessikaHelp
TEMPLATE = app


SOURCES += main.cpp\
        receiveddatadisplay.cpp \
    fileprocessing.cpp \
    projectcontroller.cpp

HEADERS  += receiveddatadisplay.h \
    fileprocessing.h \
    projectcontroller.h

FORMS    += receiveddatadisplay.ui
