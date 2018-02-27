#include "project_controller.h"
#include "tester.h"

#include <QApplication>
#include <QStringList>
#include <QDebug>

#define TAP_COMPILE
#include "libtap\cpp_tap.h"

int main(int argc, char *argv[])
{
  QApplication a(argc, argv);

  tester myTester;
  QList<QStringList> result;

  project_controller project;
 // project.StartProgram();

  plan_tests(2);
  //
  ok(project.TestMode(myTester.GetTest(0).path_, myTester.GetTest(0).objects_) ==
    myTester.GetResult(0), "Result of test 1");
  //
  project.ClearProgram();
  //
  ok(project.TestMode(myTester.GetTest(1).path_, myTester.GetTest(1).objects_) ==
    myTester.GetResult(1), "Result of test 2");
  //
  return exit_status();
  //
  return a.exec();
}
