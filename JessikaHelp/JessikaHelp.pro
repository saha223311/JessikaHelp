#-------------------------------------------------
#
# Project created by QtCreator 2018-02-16T19:09:16
#
#-------------------------------------------------

QT       += core gui
CONFIG += qaxcontainer
QMAKE_CXXFLAGS += -std=c++11

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = JessikaHelp
TEMPLATE = app


SOURCES += main.cpp\
    tester.cpp \
    file_controller.cpp \
    file_processing.cpp \
    project_controller.cpp \
    received_data_display.cpp \
    report_maker.cpp \
    libtap/cpp_patch_tap.c \
    libtap/tap.c

HEADERS  += \
    tester.h \
    file_controller.h \
    file_processing.h \
    project_controller.h \
    received_data_display.h \
    report_maker.h \
    libtap/cpp_patch_tap.h \
    libtap/cpp_tap.h \
    libtap/tap.h

FORMS    += \
    received_data_display.ui

DISTFILES +=
