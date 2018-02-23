#ifndef PROJECTCONTROLLER_H
#define PROJECTCONTROLLER_H

#include <QObject>
#include <QStringList>

#include "fileprocessing.h"
#include "receiveddatadisplay.h"
#include "reportmaker.h"

class ProjectController : public QObject
{
    Q_OBJECT
public:
    explicit ProjectController(QObject *parent = 0);
    ~ProjectController();

    void startProgram();
signals:

public slots:
    void createReports();

private:
    ReceivedDataDisplay* mReceivedDataDisplay;
    FileProcessing* mFileProcessing;
    ReportMaker* mReportMaker;
};

#endif // PROJECTCONTROLLER_H
