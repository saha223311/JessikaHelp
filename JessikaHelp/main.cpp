#include "projectcontroller.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    ProjectController project;
    project.startProgram();

    return a.exec();
}
