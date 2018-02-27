#include "project_controller.h"

project_controller::project_controller(QObject *parent) : QObject(parent) {
  receivedDataDisplay_ = new received_data_display;
  fileProcessing_ = new file_processing;
  reportMaker_ = new report_maker;

  QObject::connect(fileProcessing_, SIGNAL(ReadSomeData()),
           receivedDataDisplay_, SLOT(IncreaseProgressBar()));
  //
  QObject::connect(fileProcessing_, SIGNAL(CalculatedCountOfRowsFromPassportExcelFile(int)),
           receivedDataDisplay_,SLOT(SetMaximumValueForProgressBar(int)));
  //
  QObject::connect(fileProcessing_, SIGNAL(CalculatedCountOfRowsFromPassportExcelFile(int)),
           receivedDataDisplay_,SLOT(SetCountOfRowsFromPassportExcelFile(int)));
  //
  QObject::connect(fileProcessing_, SIGNAL(CalculatedCountOfColsFromPassportExcelFile(int)),
           receivedDataDisplay_,SLOT(SetCountOfColsFromPassportExcelFile(int)));
  //
  QObject::connect(fileProcessing_, SIGNAL(StartDataProcessing()),
           receivedDataDisplay_, SLOT(DisplayStartDataProcessing()));
  //
  QObject::connect(fileProcessing_, SIGNAL(EndDataProcessing()),
           receivedDataDisplay_, SLOT(DisplayEndDataProcessing()));
  //
  QObject::connect(fileProcessing_, SIGNAL(CalculatedApplicationDirPath(QString)),
           receivedDataDisplay_, SLOT(SetApplicationDirPath(QString)));
  //
  QObject::connect(receivedDataDisplay_, SIGNAL(NeedToGetPassportExcelModel(int)),
           fileProcessing_, SLOT(FindPassportExcelModel(int)));
  //
  QObject::connect(fileProcessing_, SIGNAL(FoundPassportExcelModel(passport_excel_model)),
           receivedDataDisplay_, SLOT(AddPassportExcelModel(passport_excel_model)));
  //
  QObject::connect(receivedDataDisplay_, SIGNAL(ReportButtonTriggered()),
           this, SLOT(CreateReports()));
  //
  QObject::connect(receivedDataDisplay_, SIGNAL(LongStorageReportButtonTriggered()),
           this, SLOT(CreateLongStorageReport()));
  //
  QObject::connect(reportMaker_, SIGNAL(StartedCreateReports()),
           receivedDataDisplay_, SLOT(StartCreateReports()));
  //
  QObject::connect(reportMaker_, SIGNAL(StartedCreateCoolReport()),
           receivedDataDisplay_, SLOT(StartCreateCoolReport()));
  //
  QObject::connect(reportMaker_, SIGNAL(StartCreatingFile(QString)),
           receivedDataDisplay_, SLOT(FileProcessing(QString)));
  //
  QObject::connect(reportMaker_, SIGNAL(EndCreatingFile()),
           receivedDataDisplay_, SLOT(EndFileProcessing()));
  //
  QObject::connect(reportMaker_, SIGNAL(EndCreateReports()),
           receivedDataDisplay_, SLOT(DisplayEndDataProcessing()));
}

void project_controller::CreateReports(){
  reportMaker_->MakeAllReports(receivedDataDisplay_->GetAllPassportExcelModels());
}

void project_controller::CreateLongStorageReport(){
  reportMaker_->MakeReportLongStorage(receivedDataDisplay_->GetAllPassportExcelModels());
}

void project_controller::StartProgram(){
  receivedDataDisplay_->show();
  fileProcessing_->ParsePassportExcelFile();
  receivedDataDisplay_->setFocus();
}

QList<QStringList> project_controller::TestMode(QString aFileName, const QList<int> &aObjects){
  receivedDataDisplay_->show();
  fileProcessing_->ParsePassportExcelFile(aFileName);
  receivedDataDisplay_->setFocus();
  receivedDataDisplay_->CloseTextBrowser();
  for (int object : aObjects){
    receivedDataDisplay_->FindButtonProcessing(object);
  }
  return receivedDataDisplay_->GetAllPassportExcelModels();
}

project_controller::~project_controller(){
  delete receivedDataDisplay_;
  delete fileProcessing_;
  delete reportMaker_;
}

void project_controller::ClearProgram(){
  emit receivedDataDisplay_->DeleteAllButtonProcessing();
}
