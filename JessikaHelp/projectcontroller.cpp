#include "projectcontroller.h"

ProjectController::ProjectController(QObject *parent) : QObject(parent) {
    mReceivedDataDisplay = new ReceivedDataDisplay;
    mFileProcessing = new FileProcessing;

    QObject::connect(mFileProcessing, SIGNAL(readSomeData()),
                     mReceivedDataDisplay, SLOT(increaseProgressBar()));

    QObject::connect(mFileProcessing, SIGNAL(calculatedCountOfRowsFromPassportExcelFile(int)),
                     mReceivedDataDisplay,SLOT(setMaximumValueForProgressBar(int)));

    QObject::connect(mFileProcessing, SIGNAL(calculatedCountOfRowsFromPassportExcelFile(int)),
                     mReceivedDataDisplay,SLOT(setCountOfRowsFromPassportExcelFile(int)));

    QObject::connect(mFileProcessing, SIGNAL(calculatedCountOfColsFromPassportExcelFile(int)),
                     mReceivedDataDisplay,SLOT(setCountOfColsFromPassportExcelFile(int)));

    QObject::connect(mFileProcessing, SIGNAL(startDataProcessing()),
                     mReceivedDataDisplay, SLOT(displayStartDataProcessing()));

    QObject::connect(mFileProcessing, SIGNAL(endDataProcessing()),
                     mReceivedDataDisplay, SLOT(displayEndDataProcessing()));

    QObject::connect(mFileProcessing, SIGNAL(calculatedApplicationDirPath(QString)),
                     mReceivedDataDisplay, SLOT(setApplicationDirPath(QString)));

    QObject::connect(mReceivedDataDisplay, SIGNAL(needToGetPassportExcelModel(int)),
                     mFileProcessing, SLOT(findPassportExcelModel(int)));

    QObject::connect(mFileProcessing, SIGNAL(foundPassportExcelModel(PassportExcelModel)),
                     mReceivedDataDisplay, SLOT(addPassportExcelModel(PassportExcelModel)));

    QObject::connect(mReceivedDataDisplay, SIGNAL(reportButtonTriggered()),
                     this, SLOT(createReports()));

}

void ProjectController::createReports(){
    mReportMaker->makeAllReports(mReceivedDataDisplay->getAllPassportExcelModels());
}

void ProjectController::startProgram(){
    mReceivedDataDisplay->show();
    mFileProcessing->parsePassportExcelFile();
    mReceivedDataDisplay->setFocus();
}

ProjectController::~ProjectController(){
    delete mReceivedDataDisplay;
    delete mFileProcessing;
}
