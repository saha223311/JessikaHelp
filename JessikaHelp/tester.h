#ifndef TESTER_H
#define TESTER_H

#include <QApplication>
#include <QString>
#include <QStringList>
#include <QList>

struct test{
  QString path_;
  QList<int> objects_;
};

class tester{
public:
  tester();
  test GetTest(int aIndex);
  QList<QStringList> GetResult(int aIndex);
private:
  QString testsPath_;
  QList<test> tests_;
  QList<QList<QStringList> > results_;
};

#endif // TESTER_H
