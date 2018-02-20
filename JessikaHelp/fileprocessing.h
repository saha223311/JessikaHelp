#ifndef FILEPROCESSING_H
#define FILEPROCESSING_H

#include <QObject>
#include <QApplication>
#include <QAxObject>
#include <QString>
#include <QList>
#include <QStringList>

#include <QDebug>

struct PassportExcelModel{
    QString culture; //Культура (разновидность) (forTable)
    QString numberVIR; //Номер в основном каталоге ВИР (forTable)
    QString introductionNumber; //Интродукционный номер
    QString countryOfOrigin ; //Страна происхождения (forTable)
    QString contryCodeOfOrigin; //Код страны происхождения
    QString sampleName ; //Название образца (forTable)
    QString fertilityLevel ; //Уровень плоидности (forTable)
    QString developmentType; //Тип развития (forTable)
    QString pedigree; //Родословная
    QString pedigreeSource; //Источние сведений о родословной
    QString sampleStatus; //Статус образца
    QString seedsAvailability; //Доступность семян
    QString registrationYearInWheatDepartment; //Год регистрации в отделе пшениц
    QString donorCountry ; //Страна-донор
    QString donorCountryCode; //Код страны-донора
    QString donorCity; //Город-донор
    QString donorAgency; //Учреждение-донор
    QString agencyAbbreviatedName; //Сокращенное название учреждения
    QString donorName; //Имя донора
    QString sampleNumberByDonor; // Номер, присвоенный образцу учреждением-донором
    QString numberInWheatCatalog; // Номер в "пшеничном каталоге"
    QString preregistrationBookNumber; // Номер книги предварительной регистрации
    QString sampleNumberInPreregistrationBook; //Порядковый номер образца в книге предварительной регистрации

    QStringList getGeneralData();
    void setPassportExcelModelData(const QStringList& data);
};

class FileProcessing : public QObject
{
    Q_OBJECT
public:
    explicit FileProcessing(QObject *parent = 0);
    void parsePassportExcelFile();
    int getCountOfRowsFromPassportExcelFile();

public slots:
    void findPassportExcelModel(int index);

signals:
    void readSomeData();
    void startDataProcessing();
    void endDataProcessing();
    void calculatedApplicationDirPath(QString path);
    void calculatedCountOfRowsFromPassportExcelFile(int);
    void calculatedCountOfColsFromPassportExcelFile(int);

    void foundPassportExcelModel(PassportExcelModel data);

private:
    //мб не занимать ими память, а работать непосредственно через методы
    QAxObject *mExcel;
    QAxObject *mWorkbooks;
    QAxObject *mWorkbook;
    QAxObject *mSheets;
    QAxObject *mStatSheet;
    QAxObject *mCell;
    QAxObject *mUsedRangeR;

    //МБ И НЕ НУЖНЫ
    // ИМ ЛУЧШЕ МЕСТО В РЕСИВД ДАТА ДИСПЛЕЙ - !?
    int mCountOfRowsFromPassportExcelFile;
    int mCountOfColsFromPassportExcelFile;

    QList<PassportExcelModel> passportTable;

};

#endif // FILEPROCESSING_H
