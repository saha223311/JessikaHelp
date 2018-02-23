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
    void makeReportWordLabel(QList<QStringList> data);

private:
    QDateTime dateTime;
};

#endif // REPORTMAKER_H
