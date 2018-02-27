#include "file_processing.h"

QStringList passport_excel_model::GetGeneralData(){
  QStringList generalData;
  generalData.push_back(this->culture_);
  generalData.push_back(this->numberVIR_);
  generalData.push_back(this->countryOfOrigin_);
  generalData.push_back(this->sampleName_);
  generalData.push_back(this->fertilityLevel_);
  generalData.push_back(this->developmentType_);
  return generalData;
}

void passport_excel_model::SetPassportExcelModelData(const QStringList &aData){
  culture_ = aData.at(0);
  numberVIR_ = aData.at(1);
  introductionNumber_ = aData.at(2);
  countryOfOrigin_ = aData.at(3);
  contryCodeOfOrigin_ = aData.at(4);
  sampleName_ = aData.at(5);
  fertilityLevel_ = aData.at(6);
  developmentType_ = aData.at(7);
  pedigree_ = aData.at(8);
  pedigreeSource_ = aData.at(9);
  sampleStatus_ = aData.at(10);
  seedsAvailability_ = aData.at(11);
  registrationYearInWheatDepartment_ = aData.at(12);
  donorCountry_ = aData.at(13);
  donorCountryCode_ = aData.at(14);
  donorCity_ = aData.at(15);
  donorAgency_ = aData.at(16);
  agencyAbbreviatedName_ = aData.at(17);
  donorName_ = aData.at(18);
  sampleNumberByDonor_ = aData.at(19);
  numberInWheatCatalog_ = aData.at(20);
  preregistrationBookNumber_ = aData.at(21);
  sampleNumberInPreregistrationBook_ = aData.at(22);
}

file_processing::file_processing(QObject *parent) : QObject(parent){
}

void file_processing::ParsePassportExcelFile(QString aPath){
  QString applicationDirPath;
  if (aPath == ""){
    QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();;

    for(int i=0;i< temporaryApplicationDirPath.length();i++){
      if(temporaryApplicationDirPath[i] == '/')  applicationDirPath+="\\";
      else applicationDirPath+=temporaryApplicationDirPath[i];
    }
    applicationDirPath+="\\passport.xls";
  } else{
    applicationDirPath = aPath;
  }
  emit CalculatedApplicationDirPath(applicationDirPath);
  excel_ = new QAxObject( "Excel.Application");
  workbooks_ = excel_->querySubObject( "Workbooks" );
  //
  passport_excel_model temporaryPassportExcelMode;
  QStringList temporaryStringList;
  //
  workbook_ = workbooks_->querySubObject( "Open(const QString&)", applicationDirPath);
  workbook_->setProperty("Save", true);
  sheets_ = workbook_->querySubObject( "Sheets" );
  sheet_ = sheets_->querySubObject ("Item(const QVariant&)", QVariant("P1") );
  //
  usedRange_ = sheet_->querySubObject("UsedRange");
  countOfRowsFromPassportExcelFile_ = usedRange_->querySubObject("Rows")->property("Count").toInt();
  emit CalculatedCountOfRowsFromPassportExcelFile(countOfRowsFromPassportExcelFile_);
  //
  countOfColsFromPassportExcelFile_ = usedRange_->querySubObject("Columns")->property("Count").toInt();
  emit CalculatedCountOfColsFromPassportExcelFile(countOfColsFromPassportExcelFile_);
  //
  emit StartDataProcessing();
  //
  int row = 2;
  while (row <= countOfRowsFromPassportExcelFile_){
    for (size_t col = 1; col < 24; col++){
      cell_ = sheet_->querySubObject("Cells(QVariant,QVariant)", row, col);
      temporaryStringList.push_back(cell_->property("Value").toString());
    }
    temporaryPassportExcelMode.SetPassportExcelModelData(temporaryStringList);
    passportTable_.push_back(temporaryPassportExcelMode);
    //
    temporaryStringList.clear();
    row++;
    emit ReadSomeData();
    qApp->processEvents();
  }
  emit ReadSomeData();
  //
  workbook_->dynamicCall("Save()");
  workbook_->dynamicCall("Close()");
  excel_->dynamicCall("SetDisplayAlerts(bool)", FALSE);
  excel_->dynamicCall("Quit()");
  //
  emit EndDataProcessing();
}

int file_processing::GetCountOfRowsFromPassportExcelFile(){
  return countOfRowsFromPassportExcelFile_;
}

void file_processing::FindPassportExcelModel(int aIndex){
  emit FoundPassportExcelModel(passportTable_[aIndex - 1]);
}


