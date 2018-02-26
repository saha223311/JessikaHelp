#ifndef REPORTMAKER_H
#define REPORTMAKER_H

#include <QObject>
#include <QDateTime>
#include <QFile>
#include <QApplication>
#include <QAxObject>
#include <QStringList>

class ReportMaker : public QObject
{
    Q_OBJECT
public:
    explicit ReportMaker(QObject *parent = 0);


signals:

public slots:
    void makeAllReports(const QList<QStringList>& data);
    void makeReportLongStorage(const QList<QStringList>& data);

    void makeReportWordLabel(const QList<QStringList>& data);
    void makeReportWordEnvelope(const QList<QStringList>& data);
    void makeReportWordEnvelopeA4(const QList<QStringList>& data);
    void makeReportExcel(const QList<QStringList>& data);

private:
    QDateTime dateTime;
};

#endif // REPORTMAKER_H
