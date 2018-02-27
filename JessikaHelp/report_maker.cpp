#include "report_maker.h"
#include <QDebug>
#include <QString>

report_maker::report_maker(QObject *parent) : QObject(parent){
  fileController_ = new file_controller;
}

void report_maker::MakeAllReports(const QList<QStringList> &aData){
  emit StartedCreateReports();
  MakeReportWordLabel(aData);
  MakeReportWordEnvelope(aData);
  MakeReportWordEnvelopeA4(aData);
  MakeReportExcel(aData);
  emit EndCreateReports();
}

QString report_maker::MakeTitle(bool aHeader, QStringList aData){
  QString text;
  if (aHeader) text = QString::fromUtf8("ФГБНУ ВИР им. Н.И.Вавилова");
  text += "\n" + aData.at(0) + "\nK-" + aData.at(1) +
          "\n" + aData.at(2) + "\n" + aData.at(3) +
          "\n" + aData.at(4) + "\n" + aData.at(5);
  if (!aData.at(6).isEmpty()){
    text += "\n" + aData.at(6);
  }
  return text;
}

void report_maker::MakeReportWordLabel(const QList<QStringList> &aData){
  QString newFile = fileController_->CreateLabel();
  emit StartCreatingFile(newFile);
  fileController_->OpenWordFile(newFile);
  //
  const int columnsLabel = 4;
  const int rowsLabel = (aData.size() + columnsLabel - 1) / columnsLabel;
  const int pageWidth = 210 - 2;
  const int marginLeft = 1;
  const int marginRight = 1;
  //
  const int widthCenter = (pageWidth - marginLeft - marginRight) / columnsLabel;
  const double kMM = 110.0 / 38.8;
  //
  fileController_->CreateWordTable(rowsLabel, columnsLabel);
  //
  int counter = 0;
  for (int x = 1; x <=rowsLabel; x++){
    for(int y = 1; y <= columnsLabel; ++y) {
      fileController_->ChooseWordCell(x, y);
      fileController_->SetWordCellWidth(widthCenter * kMM);
      fileController_->SetWordCellBoardStyle(true, true, true, true,
                                             board_style::DOT_STYLE);
      //
      if (counter < aData.size()) {
        fileController_->SetWordCellText(this->MakeTitle(false, aData.at(counter)));
      }
      else{
        if (counter != aData.size()){
          fileController_->SetWordCellBoardStyle(false, true, false, false,
                                                 board_style::NULL_STYLE);
        }
        if (x == 1){
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                                                 board_style::NULL_STYLE);
        }
        fileController_->SetWordCellBoardStyle(false, false, true, true,
                                               board_style::NULL_STYLE);
      }
      counter++;
    }
  }
  fileController_->SaveWordFile();
  fileController_->CloseWordFile();
  emit EndCreatingFile();
}


void report_maker::MakeReportWordEnvelope(const QList<QStringList> &aData){
  QString newFile = fileController_->CreateEnvelope();
  emit StartCreatingFile(newFile);
  fileController_->OpenWordFile(newFile);
  //
  const int pageWidth = 210 - 2, pageHeight = 297 - 2;
  const int marginLeft = 15;
  const int marginTop = 15, marginBottom = 15;
   //
  const int widthOverlap = 25;
  const int widthCenter = (pageWidth - widthOverlap) / 2;
  const int widthLeft = (pageWidth - widthCenter) / 2 - marginLeft;
  const int widthRight = (pageWidth - widthCenter) / 2 - widthOverlap;
  //
  const int heightOverlap = 20;
  const int heightSpace = 55;
  const int heightHalf = pageHeight / 2 ;
  const int heightText = heightHalf - heightSpace - heightOverlap;
  const int h1 = heightSpace - marginTop;
  const int h2txt = heightText;
  const int h3 = heightOverlap;
  const int h4 = heightSpace;
  const int h5txt = heightText;
  const int h6 = heightOverlap - marginBottom;
  const int cellPerPage = 6;
  const double kMM = 110.0 / 38.8;
  //
  QString dateTimeText;
  dateTimeText += dateTime_.currentDateTime().toString("yyyy-MM-dd");
  dateTimeText += " " + dateTime_.currentDateTime().toString("hh:mm");
  //
  fileController_->CreateWordTable(cellPerPage * ((aData.size() + 1) / 2), 3);
  int counter = 0;
  for (int row = 1; row <= (aData.size() + 1) / 2; row++){
    int index = (row - 1)  * cellPerPage;
    for (int x = index + 1; x <= index + 6; ++x){
      for(int y = 1; y <= 3; ++y) {
        fileController_->ChooseWordCell(x, y);
        fileController_->SetWordCellBoardStyle(false, false, false, true,
                                               board_style::DOT_STYLE);
        switch (x % 6) {
        case 1:
          fileController_->SetWordCellHeight(h1 * kMM);
          fileController_->SetWordCellBoardStyle(true, false, true, false,
                                                 board_style::NULL_STYLE);
          break;
        case 2:
          fileController_->SetWordCellHeight(h2txt * kMM);
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                                                 board_style::NULL_STYLE);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                                                 board_style::DOT_STYLE);
          break;
        case 3:
          fileController_->SetWordCellHeight(h3 * kMM);
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                                                 board_style::DOT_STYLE);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                                                 board_style::LINE_STYLE);
          break;
        case 4:
          fileController_->SetWordCellHeight(h4 * kMM);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                           board_style::NULL_STYLE);
          break;
        case 5:
          fileController_->SetWordCellHeight(h5txt * kMM);
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                           board_style::NULL_STYLE);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                           board_style::DOT_STYLE);
          break;
        case 0:
          fileController_->SetWordCellHeight(h6 * kMM);
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                           board_style::DOT_STYLE);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                           board_style::NULL_STYLE);
          break;
        default:
          break;
        }
        //
        switch (y) {
        case 1:
          fileController_->SetWordCellWidth(widthLeft * kMM);
          fileController_->SetWordCellBoardStyle(false, true, false, false,
                           board_style::NULL_STYLE);
          break;
        case 2:
          fileController_->SetWordCellWidth(widthCenter * kMM);
          break;
        case 3:
          fileController_->SetWordCellWidth(widthRight * kMM);
          break;
        default:
          break;
        }
        //
        if ((counter < aData.size()) && (y == 2) &&(((x % 6) == 2) || ((x % 6) == 5))){
          fileController_->SetWordCellText(this->MakeTitle(true, aData.at(counter)));
        }
        if ((counter < aData.size()) && (y == 2) && ((x % 3) == 0)){
          fileController_->SetWordCellText(dateTimeText);
          counter++;
        }
      }
    }
  }
  fileController_->SaveWordFile();
  fileController_->CloseWordFile();
  emit EndCreatingFile();
}

void report_maker::MakeReportWordEnvelopeA4(const QList<QStringList> &aData){
  QString newFile = fileController_->CreateEnvelopeA4();
  emit StartCreatingFile(newFile);
  fileController_->OpenWordFile(newFile);
  QString dateTimeText;
  dateTimeText += dateTime_.currentDateTime().toString("yyyy-MM-dd");
  dateTimeText += " " + dateTime_.currentDateTime().toString("hh:mm");
  //
  const int pageWidth = 297 - 2, pageHeight = 210 - 2;
  const int marginLeft = 15;
  const int marginTop = 15, marginBottom = 15;
  //
  const int widthOverlap = 30;
  const int widthCenter = (pageWidth - widthOverlap) / 2;
  const int widthLeft = (pageWidth - widthCenter) / 2 - marginLeft;
  const int widthRight = (pageWidth - widthCenter) / 2 - widthOverlap;
  //
  const int heightOverlap = 30;
  const int heightSpace = 100;
  const int heightText = pageHeight - heightSpace - heightOverlap;
  const int h1 = heightSpace - marginTop;
  const int h2txt = heightText;
  const int h3 = heightOverlap - marginBottom;
  //
  const int cellPerRecord = 3;
  const double kMM = 110.0 / 38.8;
  //
  fileController_->CreateWordTable(3 * aData.size(), 3);
  int counter = 0;
  for (int row = 1; row <= aData.size(); row++){
    int index = (row - 1)  * cellPerRecord;
    for (int x = index + 1; x <= index + 3; ++x){
      for(int y = 1; y <= 3; ++y) {
        fileController_->ChooseWordCell(x, y);
        fileController_->SetWordCellBoardStyle(false, false, false, true,
                          board_style::DOT_STYLE);
        switch (x % 3) {
        case 1:
          fileController_->SetWordCellHeight(h1 * kMM);
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                          board_style::NULL_STYLE);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                          board_style::NULL_STYLE);
          break;
        case 2:
          fileController_->SetWordCellHeight(h2txt * kMM);
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                          board_style::NULL_STYLE);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                          board_style::DOT_STYLE);
          break;
        case 0:
          fileController_->SetWordCellHeight(h3 * kMM);
          fileController_->SetWordCellBoardStyle(true, false, false, false,
                          board_style::DOT_STYLE);
          fileController_->SetWordCellBoardStyle(false, false, true, false,
                          board_style::NULL_STYLE);
          break;
        default:
          break;
        }
        switch (y) {
        case 1:
          fileController_->SetWordCellWidth(widthLeft * kMM);
          fileController_->SetWordCellBoardStyle(false, true, false, false,
                          board_style::NULL_STYLE);
          break;
        case 2:
          fileController_->SetWordCellWidth(widthCenter * kMM);
          break;
        case 3:
          fileController_->SetWordCellWidth(widthRight * kMM);
          break;
        default:
          break;
        }
        if ((counter < aData.size()) && (y == 2) &&((x % 3) == 2)){
          fileController_->SetWordCellText(this->MakeTitle(true, aData.at(counter)));
        }
        if ((counter < aData.size()) && (y == 2) && ((x % 3) == 0)){
          fileController_->SetWordCellText(dateTimeText);
          counter++;
        }
      }
    }
  }
  fileController_->SaveWordFile();
  fileController_->CloseWordFile();
  emit EndCreatingFile();
}

void report_maker::MakeReportExcel(const QList<QStringList> &aData){
  QString newFile = fileController_->CreateExcel();
  emit StartCreatingFile(newFile);
  fileController_->OpenExcelFile(newFile);
  fileController_->ChooseFirstCell(2, 1);
  fileController_->ChooseSecondCell(2 + aData.size() - 1,
                   1 + aData.at(0).size() - 1);
  fileController_->CreateExcelTable(aData);
  fileController_->SaveExcelFile();
  fileController_->CloseExcelFile();
  emit EndCreatingFile();
}

void report_maker::MakeReportLongStorage(const QList<QStringList> &aData){
  emit StartedCreateCoolReport();
  QString newFile = fileController_->CreateCool();
  emit StartCreatingFile(newFile);
  QList<QStringList> longStorageDataData;
  for (int i = 0; i < aData.size(); i++){
    QStringList temporaryStringList;
    temporaryStringList.push_back("105");
    temporaryStringList.push_back(aData.at(i).at(1));
    temporaryStringList.push_back(aData.at(i).at(4));
    temporaryStringList.push_back(aData.at(i).at(3));
    temporaryStringList.push_back(aData.at(i).at(2));
    temporaryStringList.push_back(aData.at(i).at(5));
    longStorageDataData.push_back(temporaryStringList);
  }
  fileController_->OpenExcelFile(newFile);
  fileController_->ChooseFirstCell(3, 2);
  fileController_->ChooseSecondCell(3 + longStorageDataData.size() - 1,
                   2 + longStorageDataData.at(0).size() - 1);
  fileController_->CreateExcelTable(longStorageDataData);
  fileController_->ReplaceCreatorName(longStorageDataData.size() + 3, 1);
  fileController_->SaveExcelFile();
  fileController_->CloseExcelFile();
  //
  emit EndCreatingFile();
  emit EndCreateReports();
}

report_maker::~report_maker(){
  delete fileController_;
}
