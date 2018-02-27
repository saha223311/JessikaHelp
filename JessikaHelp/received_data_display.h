#ifndef RECEIVEDDATADISPLAY_H
#define RECEIVEDDATADISPLAY_H

#include <QWidget>
#include <QKeyEvent>
#include <QStringList>
#include "file_processing.h"

namespace Ui {
class ReceivedDataDisplay;
}

class received_data_display : public QWidget{
  Q_OBJECT
public:
  explicit received_data_display(QWidget *parent = 0);
  QList<QStringList> GetAllPassportExcelModels();
  ~received_data_display();
protected:
  void KeyPressEvent(QKeyEvent*);
public slots:
  void SetMaximumValueForProgressBar(int aValue);
  void SetCountOfRowsFromPassportExcelFile(int aValue);
  void SetCountOfColsFromPassportExcelFile(int aValue);
  void DisplayStartDataProcessing();
  void DisplayEndDataProcessing();
  void SetApplicationDirPath(QString aPath);
  void IncreaseProgressBar();
  void StartCreateReports();
  void StartCreateCoolReport();
  void FileProcessing(QString aFileName);
  void EndFileProcessing();
  void CloseTextBrowser();
  void AppendTextToTextBrowser(QString aText);
  //
  void FindButtonProcessing(int object = 0);
  void AddPassportExcelModel(passport_excel_model aData);
  //
  void DeleteButtonProcessing();
  void DeleteAllButtonProcessing();
signals:
  void NeedToGetPassportExcelModel(int);
  void ReportButtonTriggered();
  void LongStorageReportButtonTriggered();
private:
  Ui::ReceivedDataDisplay *ui;
  //
  int countOfRowsFromPassportExcelFile_;
  int countOfColsFromPassportExcelFile_;
  //
  void UpdateCountLabel();
};

#endif // RECEIVEDDATADISPLAY_H
