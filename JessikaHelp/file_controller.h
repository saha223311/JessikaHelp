#ifndef FILECONTROLLER_H
#define FILECONTROLLER_H

#include <QApplication>
#include <QString>
#include <QFile>
#include <QDateTime>
#include <QAxObject>
#include <QStringList>

enum board_style {
  NULL_STYLE = 0,
  LINE_STYLE = 1,
  DOT_STYLE = 4
};

class file_controller{
public:
  file_controller();
  ~file_controller();
  //
  QString CreateLabel();
  QString CreateEnvelope();
  QString CreateEnvelopeA4();
  QString CreateExcel();
  QString CreateCool();
  //
  void OpenWordFile(QString aFileName);
  void CreateWordTable(int aRows, int aCols);
  void ChooseWordCell(int aRow, int aCol);
  void SetWordCellWidth(int aWidth);
  void SetWordCellHeight(int aHeight);
  void SetWordCellBoardStyle(bool aUp, bool aLeft, bool aDown, bool aRight,
                               board_style aStyle);
  //
  void SetWordCellText(QString aText);
  void SaveWordFile();
  void CloseWordFile();
  //
  void OpenExcelFile(QString aFileName);
  void ChooseFirstCell(int aRow, int aCol);
  void ChooseSecondCell(int aRow, int aCol);
  void CreateExcelTable(const QList<QStringList>& aData);
  void SaveExcelFile();
  void CloseExcelFile();
  void ReplaceCreatorName(int aRow, int aCol);
private:
  QDateTime dateTime_;
  //
  QAxObject* word_;
  QAxObject* table_;
  QAxObject* cell_;
  //
  QAxObject* excel_;
  QAxObject* workbook_;
  QAxObject* sheet_;
  QAxObject* range_;
  //
  QAxObject* firstCell_;
  QAxObject* secondCell_;
};

#endif // FILECONTROLLER_H
