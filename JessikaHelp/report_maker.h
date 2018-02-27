#ifndef REPORTMAKER_H
#define REPORTMAKER_H

#include <QObject>
#include <QDateTime>
#include <QFile>
#include <QApplication>
#include <QAxObject>
#include <QStringList>
#include "file_controller.h"

class report_maker : public QObject{
  Q_OBJECT
public:
  explicit report_maker(QObject *parent = 0);
  ~report_maker();
signals:
  void StartedCreateReports();
  void StartedCreateCoolReport();
  void StartCreatingFile(QString);
  void EndCreatingFile();
  void EndCreateReports();
public slots:
  void MakeAllReports(const QList<QStringList>& aData);
  void MakeReportLongStorage(const QList<QStringList>& aData);
  //
  void MakeReportWordLabel(const QList<QStringList>& aData);
  void MakeReportWordEnvelope(const QList<QStringList>& aData);
  void MakeReportWordEnvelopeA4(const QList<QStringList>& aData);
  void MakeReportExcel(const QList<QStringList>& aData);
private:
  QDateTime dateTime_;
  file_controller* fileController_;
  QString MakeTitle(bool aHeader, QStringList aData);
};

#endif // REPORTMAKER_H
