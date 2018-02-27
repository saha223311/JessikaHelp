#ifndef FILEPROCESSING_H
#define FILEPROCESSING_H

#include <QObject>
#include <QApplication>
#include <QAxObject>
#include <QString>
#include <QStringList>

struct passport_excel_model{
  // Культура (разновидность) (forTable)
  QString culture_;
  // Номер в основном каталоге ВИР (forTable)
  QString numberVIR_;
  // Интродукционный номер
  QString introductionNumber_;
  // Страна происхождения (forTable)
  QString countryOfOrigin_ ;
  // Код страны происхождения
  QString contryCodeOfOrigin_;
  // Название образца (forTable)
  QString sampleName_ ;
  // Уровень плоидности (forTable)
  QString fertilityLevel_ ;
  // Тип развития (forTable)
  QString developmentType_;
  // Родословная
  QString pedigree_;
  // Источние сведений о родословной
  QString pedigreeSource_;
  // Статус образца
  QString sampleStatus_;
  // Доступность семян
  QString seedsAvailability_;
  // Год регистрации в отделе пшениц
  QString registrationYearInWheatDepartment_;
  // Страна-донор
  QString donorCountry_ ;
  // Код страны-донора
  QString donorCountryCode_;
  // Город-донор
  QString donorCity_;
  // Учреждение-донор
  QString donorAgency_;
  // Сокращенное название учреждения
  QString agencyAbbreviatedName_;
  // Имя донора
  QString donorName_;
  // Номер, присвоенный образцу учреждением-донором
  QString sampleNumberByDonor_;
  // Номер в "пшеничном каталоге"
  QString numberInWheatCatalog_;
  // Номер книги предварительной регистрации
  QString preregistrationBookNumber_;
  // Порядковый номер образца в книге предварительной регистрации
  QString sampleNumberInPreregistrationBook_;
  //
  QStringList GetGeneralData();
  void SetPassportExcelModelData(const QStringList& aData);
};

class file_processing : public QObject{
  Q_OBJECT
public:
  explicit file_processing(QObject *parent = 0);
  void ParsePassportExcelFile(QString aPath = "");
  int GetCountOfRowsFromPassportExcelFile();
public slots:
  void FindPassportExcelModel(int aIndex);
signals:
  void ReadSomeData();
  void StartDataProcessing();
  void EndDataProcessing();
  void CalculatedApplicationDirPath(QString);
  void CalculatedCountOfRowsFromPassportExcelFile(int);
  void CalculatedCountOfColsFromPassportExcelFile(int);
  void FoundPassportExcelModel(passport_excel_model);
private:
  QAxObject *excel_;
  QAxObject *workbooks_;
  QAxObject *workbook_;
  QAxObject *sheets_;
  QAxObject *sheet_;
  QAxObject *cell_;
  QAxObject *usedRange_;
  //
  int countOfRowsFromPassportExcelFile_;
  int countOfColsFromPassportExcelFile_;
  //
  QList<passport_excel_model> passportTable_;
};

#endif // FILEPROCESSING_H
