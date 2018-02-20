#include "fileprocessing.h"

QStringList PassportExcelModel::getGeneralData(){
    QStringList generalData;

    generalData.push_back(this->culture);
    generalData.push_back(this->numberVIR);
    generalData.push_back(this->countryOfOrigin);
    generalData.push_back(this->sampleName);
    generalData.push_back(this->fertilityLevel);
    generalData.push_back(this->developmentType);

    return generalData;
}

void PassportExcelModel::setPassportExcelModelData(const QStringList &data){
   // if (right.size() != 20) return ERROR::INVALID_SIZE
    culture = data.at(0);
    numberVIR = data.at(1);
    introductionNumber = data.at(2);
    countryOfOrigin = data.at(3);
    contryCodeOfOrigin = data.at(4);
    sampleName = data.at(5);
    fertilityLevel = data.at(6);
    developmentType = data.at(7);
    pedigree = data.at(8);
    pedigreeSource = data.at(9);
    sampleStatus = data.at(10);
    seedsAvailability = data.at(11);
    registrationYearInWheatDepartment = data.at(12);
    donorCountry = data.at(13);
    donorCountryCode = data.at(14);
    donorCity = data.at(15);
    donorAgency = data.at(16);
    agencyAbbreviatedName = data.at(17);
    donorName = data.at(18);
    sampleNumberByDonor = data.at(19);
    numberInWheatCatalog = data.at(20);
    preregistrationBookNumber = data.at(21);
    sampleNumberInPreregistrationBook = data.at(22);
}

FileProcessing::FileProcessing(QObject *parent) : QObject(parent){

}

void FileProcessing::parsePassportExcelFile(){

    QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();;
    QString applicationDirPath;

    for(int i=0;i< temporaryApplicationDirPath.length();i++){
        if(temporaryApplicationDirPath[i] == '/')  applicationDirPath+="\\";
        else applicationDirPath+=temporaryApplicationDirPath[i];
    }
    applicationDirPath+="\\passport.xlsx";

    emit calculatedApplicationDirPath(applicationDirPath);

    mExcel = new QAxObject( "Excel.Application");
    mWorkbooks = mExcel->querySubObject( "Workbooks" );

    PassportExcelModel temporaryPassportExcelMode;
    QStringList temporaryStringList;

    mWorkbook = mWorkbooks->querySubObject( "Open(const QString&)", applicationDirPath);
    mWorkbook->setProperty("Save", true);
    mSheets = mWorkbook->querySubObject( "Sheets" );
    mStatSheet = mSheets->querySubObject ("Item(const QVariant&)", QVariant("P1") );
    //TODO: Сделать исключения, если не получается открывать различные отделы (книгу, файл и тд)

    mUsedRangeR = mStatSheet->querySubObject("UsedRange");

    mCountOfRowsFromPassportExcelFile = mUsedRangeR->querySubObject("Rows")->property("Count").toInt();
    emit calculatedCountOfRowsFromPassportExcelFile(mCountOfRowsFromPassportExcelFile);

    mCountOfColsFromPassportExcelFile = mUsedRangeR->querySubObject("Columns")->property("Count").toInt();
    emit calculatedCountOfColsFromPassportExcelFile(mCountOfColsFromPassportExcelFile);

    emit startDataProcessing();

    int row = 2;
    while (row <= mCountOfRowsFromPassportExcelFile){
        for (size_t col = 1; col < 24; col++){ // CONST = 24
            mCell = mStatSheet->querySubObject("Cells(QVariant,QVariant)", row, col);
            //if (mCell->property("Value").toString() == "") ERROR::NULL_VALUE IN ROW COL
            temporaryStringList.push_back(mCell->property("Value").toString());
        }
        temporaryPassportExcelMode.setPassportExcelModelData(temporaryStringList);
        passportTable.push_back(temporaryPassportExcelMode);

        temporaryStringList.clear();
        row++;
        emit readSomeData();
        qApp->processEvents();
    }
    emit readSomeData();

    //TODO: ПРОИЗВОДИТЬ ЗАКРЫТИЕ ФАЙЛА ПРИ РАЗЛИЧНЫХ ЗАВЕРШЕНИЯХ ПРОГРАММЫ
    mWorkbook->dynamicCall("Save()");
    mWorkbook->dynamicCall("Close()");
    mExcel->dynamicCall("SetDisplayAlerts(bool)", FALSE);
    mExcel->dynamicCall("Quit()");

    emit endDataProcessing();
}

int FileProcessing::getCountOfRowsFromPassportExcelFile(){
    return mCountOfRowsFromPassportExcelFile;
}

void FileProcessing::findPassportExcelModel(int index){
    emit foundPassportExcelModel(passportTable[index - 1]);
}


