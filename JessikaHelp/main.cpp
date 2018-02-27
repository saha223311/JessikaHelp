#include "project_controller.h"
#include "tester.h"

#include <QApplication>
#include <QStringList>
#include <QDebug>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
/*
    Tester tester;
    QList<QStringList> result;
*/
    project_controller project;
    project.StartProgram();
/*
    if (project.testMode(tester.getTest(0).path, tester.getTest(0).objects)
            == tester.getResult(0)) qDebug() << "1";

    project.clearProgram();

    if (project.testMode(tester.getTest(0).path, tester.getTest(0).objects)
            == tester.getResult(0)) qDebug() << "2";
*/
    return a.exec();
}
