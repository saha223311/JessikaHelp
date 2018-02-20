#ifndef PROJECTCONTROLLER_H
#define PROJECTCONTROLLER_H

#include <QObject>

#include "fileprocessing.h"
#include "receiveddatadisplay.h"

class ProjectController : public QObject
{
    Q_OBJECT
public:
    explicit ProjectController(QObject *parent = 0);
    ~ProjectController();

    void startProgram();
signals:

public slots:

private:
    ReceivedDataDisplay* mReceivedDataDisplay;
    FileProcessing* mFileProcessing;
};

#endif // PROJECTCONTROLLER_H
