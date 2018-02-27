#include "file_controller.h"

file_controller::file_controller(){
}

QString file_controller::CreateLabel(){
  QString newFile = "D:\\";
  newFile += dateTime_.currentDateTime().toString("yyyy-MM-dd");
  newFile += "_" + dateTime_.currentDateTime().toString("hh-mm");
  newFile += "_label.doc";
  //
  QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();
  QString applicationDirPath;
  for(int i=0;i< temporaryApplicationDirPath.length();i++){
    if(temporaryApplicationDirPath[i] == '/'){
      applicationDirPath+="\\";
    } else{
      applicationDirPath+=temporaryApplicationDirPath[i];
    }
  }
  applicationDirPath+="\\template_label.doc";
  QFile::copy(applicationDirPath, newFile);
  return newFile;
}

QString file_controller::CreateEnvelope(){
  QString newFile = "D:\\";
  newFile += dateTime_.currentDateTime().toString("yyyy-MM-dd");
  newFile += "_" + dateTime_.currentDateTime().toString("hh-mm");
  newFile += "_envelope.doc";
  //
  QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();
  QString applicationDirPath;
  for(int i=0;i< temporaryApplicationDirPath.length();i++){
    if(temporaryApplicationDirPath[i] == '/'){
      applicationDirPath+="\\";
    } else{
      applicationDirPath+=temporaryApplicationDirPath[i];
    }
  }
  applicationDirPath+="\\template_envelope.doc";
  QFile::copy(applicationDirPath, newFile);
  return newFile;
}

QString file_controller::CreateEnvelopeA4(){
  QString newFile = "D:\\";
  newFile += dateTime_.currentDateTime().toString("yyyy-MM-dd");
  newFile += "_" + dateTime_.currentDateTime().toString("hh-mm");
  newFile += "_envelope_A4.doc";
  //
  QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();
  QString applicationDirPath;
  for(int i=0;i< temporaryApplicationDirPath.length();i++){
    if(temporaryApplicationDirPath[i] == '/'){
      applicationDirPath+="\\";
    } else{
      applicationDirPath+=temporaryApplicationDirPath[i];
    }
  }
  applicationDirPath+="\\template_envelope_A4.doc";
  QFile::copy(applicationDirPath, newFile);
  return newFile;
}

QString file_controller::CreateExcel(){
  QString newFile = "D:\\";
  newFile += dateTime_.currentDateTime().toString("yyyy-MM-dd");
  newFile += "_" + dateTime_.currentDateTime().toString("hh-mm");
  newFile += ".xls";
  //
  QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();
  QString applicationDirPath;
  for(int i=0;i< temporaryApplicationDirPath.length();i++){
    if(temporaryApplicationDirPath[i] == '/'){
      applicationDirPath+="\\";
    } else{
      applicationDirPath+=temporaryApplicationDirPath[i];
    }
  }
  applicationDirPath+="\\template_excel.xls";
  QFile::copy(applicationDirPath, newFile);
  return newFile;
}

QString file_controller::CreateCool(){
  QString newFile = "D:\\";
  newFile += dateTime_.currentDateTime().toString("yyyy-MM-dd");
  newFile += "_" + dateTime_.currentDateTime().toString("hh-mm");
  newFile += "_cool.xls";
  //
  QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();
  QString applicationDirPath;
  for(int i=0;i< temporaryApplicationDirPath.length();i++){
    if(temporaryApplicationDirPath[i] == '/'){
      applicationDirPath+="\\";
    } else{
      applicationDirPath+=temporaryApplicationDirPath[i];
    }
  }
  applicationDirPath+="\\template_cool.xls";
  QFile::copy(applicationDirPath, newFile);
  return newFile;
}

void file_controller::OpenWordFile(QString aFileName){
  word_ = new QAxObject( "Word.Application");
  word_->querySubObject("Documents")->querySubObject( "Open(const QString&)", aFileName);
  word_->querySubObject("Selection")->querySubObject("Start", 0);
  word_->querySubObject("Selection")->querySubObject("End", 0);
}

void file_controller::CreateWordTable(int aRows, int aCols){
  table_ = word_->querySubObject("ActiveDocument")->querySubObject("Tables")->
          querySubObject("Add(Range, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                         word_->querySubObject("ActiveDocument")->dynamicCall("Range()"),
                         aRows, aCols, 1, 1);
}

void file_controller::ChooseWordCell(int aRow, int aCol){
  cell_ = table_->querySubObject("Cell(Row, Column)", QVariant(aRow), QVariant(aCol)) ;
}

void file_controller::SetWordCellWidth(int aWidth){
  cell_->setProperty("Width", aWidth);
}

void file_controller::SetWordCellHeight(int aHeight){
  cell_->setProperty("Height", aHeight);
}

void file_controller::SetWordCellBoardStyle(bool aUp, bool aLeft, bool aDown, bool aRight,
                       board_style aStyle){
  if (aUp){
    cell_->querySubObject("Borders(xlEdge)", 1)
        ->setProperty("LineStyle", aStyle);
  }
  if (aLeft){
    cell_->querySubObject("Borders(xlEdge)", 2)
        ->setProperty("LineStyle", aStyle);
  }
  if (aDown){
    cell_->querySubObject("Borders(xlEdge)", 3)
        ->setProperty("LineStyle", aStyle);
  }
  if (aRight){
    cell_->querySubObject("Borders(xlEdge)", 4)
        ->setProperty("LineStyle", aStyle);
  }
}

void file_controller::SetWordCellText(QString aText){
  cell_->querySubObject("Range")->querySubObject("Text", aText);
}

void file_controller::SaveWordFile(){
  word_->querySubObject("Documents")->dynamicCall("Save()");
}

void file_controller::CloseWordFile(){
  word_->querySubObject("Documents")->dynamicCall("Close()");
  word_->dynamicCall("SetDisplayAlerts(bool)", FALSE);
  word_->dynamicCall("Quit()");
}

void file_controller::OpenExcelFile(QString aFileName){
  excel_ = new QAxObject( "Excel.Application");
  workbook_ = excel_->querySubObject( "Workbooks" )
      ->querySubObject( "Open(const QString&)", aFileName);
  sheet_ = workbook_->querySubObject( "Sheets" )
      ->querySubObject("Item(const QVariant&)", QVariant("P1") );
}

void file_controller::ChooseFirstCell(int aRow, int aCol){
  firstCell_ = sheet_->querySubObject("Cells(QVariant&,QVariant&)", aRow, aCol);
}

void file_controller::ChooseSecondCell(int aRow, int aCol){
  secondCell_ = sheet_->querySubObject("Cells(QVariant&,QVariant&)", aRow, aCol);
}

void file_controller::CreateExcelTable(const QList<QStringList> &aData){
  range_ = sheet_->querySubObject("Range(const QVariant&,const QVariant&)",
                     firstCell_->asVariant(), secondCell_->asVariant());
  QList<QVariant> rowsList;
  for (int i = 0; i < aData.size(); i++){
    rowsList << QVariant(aData.at(i));
  }
  range_->setProperty("Value", QVariant(rowsList) );
}

void file_controller::SaveExcelFile(){
  workbook_->dynamicCall("Save()");
}

void file_controller::CloseExcelFile(){
  workbook_->dynamicCall("Close()");
  excel_->dynamicCall("SetDisplayAlerts(bool)", FALSE);
  excel_->dynamicCall("Quit()");
}

void file_controller::ReplaceCreatorName(int aRow, int aCol){
  cell_ = sheet_->querySubObject("Cells(QVariant,QVariant)", 4, 1);
  QVariant sign = cell_->property("Value");
  cell_->setProperty("Value", QVariant(""));
  cell_ = sheet_->querySubObject("Cells(QVariant,QVariant)", aRow, aCol);
  cell_->setProperty("Value", QVariant(sign));
}

file_controller::~file_controller(){
  delete word_;
  delete table_;
  delete cell_;
  delete excel_;
  delete workbook_;
  delete sheet_;
  delete range_;
  delete firstCell_;
  delete secondCell_;
}

