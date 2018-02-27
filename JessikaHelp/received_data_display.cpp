#include "received_data_display.h"
#include "ui_receiveddatadisplay.h"

#include <QRegExpValidator>
#include <QMessageBox>

received_data_display::received_data_display(QWidget *parent) :
  QWidget(parent),
  ui(new Ui::ReceivedDataDisplay){
  ui->setupUi(this);
  ui->progressBar->setValue(0);
  ui->textBrowser->setReadOnly(true);
  ui->textBrowser->setFontWeight(10);
  //
  QRegExp rx("[1-9]\\d{0,3}");
  QValidator *validator = new QRegExpValidator (rx, 0);
  ui->itemNumberLineEdit->setValidator(validator);
  QFont font;
  font.setPointSize(12);
  font.setFamily("MS Shell Dlg 2");
  ui->textBrowser->setFont(font);
  this->setFixedSize(this->size());
  ui->tableWidget->horizontalHeader()->resizeSections(QHeaderView::ResizeToContents);
  //
  QObject::connect(ui->findButton, SIGNAL(clicked()),
           this, SLOT(FindButtonProcessing()));
  //
  QObject::connect(ui->deleteButton, SIGNAL(clicked()),
           this, SLOT(DeleteButtonProcessing()));
  //
  QObject::connect(ui->deleteAllButton, SIGNAL(clicked()),
           this, SLOT(DeleteAllButtonProcessing()));
  //
  QObject::connect(ui->reportButton, SIGNAL(clicked()),
           this, SIGNAL(ReportButtonTriggered()));
  //
  QObject::connect(ui->longStorageReportButton, SIGNAL(clicked()),
           this, SIGNAL(LongStorageReportButtonTriggered()));
}


void received_data_display::SetMaximumValueForProgressBar(int aValue){
  ui->progressBar->setMaximum(aValue);
}

void received_data_display::StartCreateReports(){
  ui->progressBar->setValue(0);
  ui->progressBar->setMaximum(4);
  ui->textBrowser->show();
}

void received_data_display::StartCreateCoolReport(){
  ui->progressBar->setValue(0);
  ui->progressBar->setMaximum(1);
  ui->textBrowser->show();
}

void received_data_display::AppendTextToTextBrowser(QString aText){
  ui->textBrowser->append(aText);
}

void received_data_display::FileProcessing(QString aFileName){
  ui->textBrowser->append(QString::fromUtf8("Идет запись в файл: ")
              + aFileName + "\n");
}

void received_data_display::EndFileProcessing(){
  IncreaseProgressBar();
}

void received_data_display::IncreaseProgressBar(){
  ui->progressBar->setValue(ui->progressBar->value() + 1);
}

void received_data_display::SetCountOfRowsFromPassportExcelFile(int aValue){
  countOfRowsFromPassportExcelFile_ = aValue;
  ui->textBrowser->append(QString::fromUtf8("Количество строк: ") +
               QString::number(countOfRowsFromPassportExcelFile_));
}

void received_data_display::SetCountOfColsFromPassportExcelFile(int aValue){
  countOfColsFromPassportExcelFile_ = aValue;
  ui->textBrowser->append(QString::fromUtf8("Количество столбцов: ") +
               QString::number(countOfColsFromPassportExcelFile_));
}

void received_data_display::DisplayStartDataProcessing(){
  ui->textBrowser->append(QString::fromUtf8("\nПроисходит загрузка данных"));
}


void received_data_display::DisplayEndDataProcessing(){
  ui->textBrowser->append(QString::fromUtf8("Данные успешно загружены"));
  ui->textBrowser->append(QString::fromUtf8("\nДля начала работы нажмите любую клавишу.."));
}


void received_data_display::SetApplicationDirPath(QString aPath){
  ui->textBrowser->append(QString::fromUtf8("Открытие excel-файла: ") + aPath);
}

void received_data_display::KeyPressEvent(QKeyEvent*){
  if (ui->progressBar->value() == ui->progressBar->maximum()){
    ui->textBrowser->clear();
    ui->textBrowser->hide();
  }
}

void received_data_display::CloseTextBrowser(){
  ui->textBrowser->clear();
  ui->textBrowser->hide();
}

void received_data_display::FindButtonProcessing(int object){
  if (object == 0){
    if (!ui->itemNumberLineEdit->text().isEmpty()){
      int itemNumber = ui->itemNumberLineEdit->text().toInt();
      if (itemNumber < countOfRowsFromPassportExcelFile_){
      emit NeedToGetPassportExcelModel(itemNumber);
      } else{
        QMessageBox::warning(this, QString::fromUtf8("Внимание"),
                  QString::fromUtf8( "Не удалось найти заданный объект."));
      }
      ui->itemNumberLineEdit->clear();
    }
  }
  else{
    if (object < countOfRowsFromPassportExcelFile_){
    emit NeedToGetPassportExcelModel(object);
    }
    else{
      QMessageBox::warning(this, QString::fromUtf8("Внимание"),
                QString::fromUtf8( "Не удалось найти заданный объект."));
    }
    ui->itemNumberLineEdit->clear();
  }
}

void received_data_display::AddPassportExcelModel(passport_excel_model aData){
  QFont font;
  font.setPointSize(11);
  font.setFamily("MS Shell Dlg 2");
  ui->tableWidget->setRowCount(ui->tableWidget->rowCount() + 1);
  for (int i = 0; i < aData.GetGeneralData().size(); i++){
    QTableWidgetItem* item = new QTableWidgetItem;
    item->setFont(font);
    item->setText(aData.GetGeneralData().at(i));
    ui->tableWidget->setItem(ui->tableWidget->rowCount() - 1, i, item);
  }
  QTableWidgetItem* item = new QTableWidgetItem;
  ui->tableWidget->setItem(ui->tableWidget->rowCount() - 1,
               aData.GetGeneralData().size(), item);
  this->UpdateCountLabel();
}

void received_data_display::DeleteButtonProcessing(){
  ui->tableWidget->removeRow(ui->tableWidget->currentRow());
  this->UpdateCountLabel();
}

void received_data_display::DeleteAllButtonProcessing(){
  int countOfRow = ui->tableWidget->rowCount();
  for (int i = 0; i < countOfRow; i++){
    ui->tableWidget->removeRow(0);
  }
  this->UpdateCountLabel();
}

void received_data_display::UpdateCountLabel(){
  ui->countLabel->setText(QString::fromUtf8("Количество: ")
              + QString::number(ui->tableWidget->rowCount()));
}

QList<QStringList> received_data_display::GetAllPassportExcelModels(){
  QList<QStringList> models;
  for (int i = 0; i < ui->tableWidget->rowCount(); i++){
    QStringList column;
    column << ui->tableWidget->item(i, 0)->text() <<
              ui->tableWidget->item(i, 1)->text() <<
              ui->tableWidget->item(i, 2)->text() <<
              ui->tableWidget->item(i, 3)->text() <<
              ui->tableWidget->item(i, 4)->text() <<
              ui->tableWidget->item(i, 5)->text() <<
              ui->tableWidget->item(i, 6)->text();
    models.push_back(column);
  }
  return models;
}


received_data_display::~received_data_display(){
  DeleteAllButtonProcessing();
  delete ui;
}
