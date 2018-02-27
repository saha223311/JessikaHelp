#include "tester.h"

tester::tester() {
  QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();
  for(int i=0;i< temporaryApplicationDirPath.length();i++){
    if(temporaryApplicationDirPath[i] == '/'){
      testsPath_+="\\";
    } else{
      testsPath_+=temporaryApplicationDirPath[i];
    }
  }
  testsPath_+="\\Tests\\";
  //
  test test1;
  test1.path_ = testsPath_ + "test_passport_1.xls";
  test1.objects_ << 1 << 3 << 5 << 7 << 9;
  //
  QList<QStringList> result;
  result.push_back(QStringList() << QString::fromUtf8("тритикале") << QString::fromUtf8("1") <<
                    QString::fromUtf8("Германия") << QString::fromUtf8("Rimpau") <<
                    QString::fromUtf8("8x") << QString::fromUtf8("п/яр.") << QString::fromUtf8(""));
  result.push_back(QStringList() << QString::fromUtf8("тритикале") << QString::fromUtf8("3") <<
                    QString::fromUtf8("Россия, Московская обл.") << QString::fromUtf8("25 АД 20") <<
                    QString::fromUtf8("8x") << QString::fromUtf8("яр.") << QString::fromUtf8(""));
  result.push_back(QStringList() << QString::fromUtf8("тритикале") << QString::fromUtf8("5") <<
                    QString::fromUtf8("Россия, Московская обл.") << QString::fromUtf8("31 АД 72") <<
                    QString::fromUtf8("8x") << QString::fromUtf8("оз.") << QString::fromUtf8(""));
  result.push_back(QStringList() << QString::fromUtf8("тритикале") << QString::fromUtf8("7") <<
                    QString::fromUtf8("Россия, Московская обл.") << QString::fromUtf8("АД 110") <<
                    QString::fromUtf8("8x") << QString::fromUtf8("оз.") << QString::fromUtf8(""));
  result.push_back(QStringList() << QString::fromUtf8("тритикале") << QString::fromUtf8("9") <<
                    QString::fromUtf8("Россия, Московская обл.") << QString::fromUtf8("АД 114") <<
                    QString::fromUtf8("8x") << QString::fromUtf8("оз.") << QString::fromUtf8(""));
  results_.push_back(result);

  //
  test test2;
  test2.path_ = testsPath_ + "test_passport_1.xls";
  test2.objects_ << 30 << 60 << 70;
  //
  result.clear();
  result.push_back(QStringList() << QString::fromUtf8("тритикале") << QString::fromUtf8("30") <<
                    QString::fromUtf8("Россия, Московская обл.") << QString::fromUtf8("25 АД 20") <<
                    QString::fromUtf8("8x") << QString::fromUtf8("яр.") << QString::fromUtf8(""));
  result.push_back(QStringList() << QString::fromUtf8("тритикале") << QString::fromUtf8("60") <<
                    QString::fromUtf8("Россия, Московская обл.") << QString::fromUtf8("НАД 120") <<
                    QString::fromUtf8("8x") << QString::fromUtf8("оз.") << QString::fromUtf8(""));
  results_.push_back(result);
  //
  tests_.push_back(test1);
  tests_.push_back(test2);
}

test tester::GetTest(int aIndex){
  return tests_.at(aIndex);
}

QList<QStringList> tester::GetResult(int aIndex){
  return results_.at(aIndex);
}

