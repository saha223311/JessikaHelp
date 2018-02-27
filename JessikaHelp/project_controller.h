#ifndef PROJECTCONTROLLER_H
#define PROJECTCONTROLLER_H

#include <QObject>
#include <QStringList>
#include "file_processing.h"
#include "received_data_display.h"
#include "report_maker.h"

class project_controller : public QObject{
  Q_OBJECT
public:
  explicit project_controller(QObject *parent = 0);
  ~project_controller();
  //
  void StartProgram();
  void ClearProgram();
  QList<QStringList> TestMode(QString aFileName, const QList<int>& aObjects);
public slots:
  void CreateReports();
  void CreateLongStorageReport();
private:
  received_data_display* receivedDataDisplay_;
  file_processing* fileProcessing_;
  report_maker* reportMaker_;
};

#endif // PROJECTCONTROLLER_H
