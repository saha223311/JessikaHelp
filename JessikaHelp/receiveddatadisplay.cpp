#include "receiveddatadisplay.h"
#include "ui_receiveddatadisplay.h"

ReceivedDataDisplay::ReceivedDataDisplay(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::ReceivedDataDisplay){
    ui->setupUi(this);

    ui->progressBar->setValue(0);
    ui->textBrowser->setReadOnly(true);
    ui->textBrowser->setFontWeight(10);

    QFont font;
    font.setPointSize(12);
    font.setFamily("MS Shell Dlg 2");
    ui->textBrowser->setFont(font);

    ui->tableWidget->horizontalHeader()->resizeSections(QHeaderView::ResizeToContents);

    QObject::connect(ui->findButton, SIGNAL(clicked()),
                     this, SLOT(findButtonProcessing()));

    QObject::connect(ui->deleteButton, SIGNAL(clicked()),
                     this, SLOT(deleteButtonProcessing()));

    QObject::connect(ui->deleteAllButton, SIGNAL(clicked()),
                     this, SLOT(deleteAllButtonProcessing()));

    QObject::connect(ui->reportButton, SIGNAL(clicked()),
                     this, SIGNAL(reportButtonTriggered()));

    QObject::connect(ui->longStorageReportButton, SIGNAL(clicked()),
                     this, SIGNAL(longStorageReportButtonTriggered()));
}


void ReceivedDataDisplay::setMaximumValueForProgressBar(int value){
    ui->progressBar->setMaximum(value);
}

void ReceivedDataDisplay::increaseProgressBar(){
    ui->progressBar->setValue(ui->progressBar->value() + 1);
}

void ReceivedDataDisplay::setCountOfRowsFromPassportExcelFile(int value){
    mCountOfRowsFromPassportExcelFile = value;
    ui->textBrowser->append(QString::fromUtf8("Количество строк: ") +
                             QString::number(mCountOfRowsFromPassportExcelFile));
}

void ReceivedDataDisplay::setCountOfColsFromPassportExcelFile(int value){
    mCountOfColsFromPassportExcelFile = value;
    ui->textBrowser->append(QString::fromUtf8("Количество столбцов: ") +
                             QString::number(mCountOfColsFromPassportExcelFile));
}

void ReceivedDataDisplay::displayStartDataProcessing(){
    ui->textBrowser->append(QString::fromUtf8("\nПроисходит загрузка данных"));
}

void ReceivedDataDisplay::displayEndDataProcessing(){
    ui->textBrowser->append(QString::fromUtf8("Данные успешно загружены"));
    ui->textBrowser->append(QString::fromUtf8("\nДля начала работы нажмите любую клавишу.."));
}

void ReceivedDataDisplay::setApplicationDirPath(QString path){
    ui->textBrowser->append(QString::fromUtf8("Открытие excel-файла: ") + path);
}

void ReceivedDataDisplay::keyPressEvent(QKeyEvent*){
    if (ui->progressBar->value() == mCountOfRowsFromPassportExcelFile){
        ui->textBrowser->hide();
    }
}

void ReceivedDataDisplay::findButtonProcessing(){

    emit needToGetPassportExcelModel(ui->itemNumberLineEdit->text().toInt());
    ui->itemNumberLineEdit->clear();
}

void ReceivedDataDisplay::addPassportExcelModel(PassportExcelModel data){
    QFont font;
    font.setPointSize(11);
    font.setFamily("MS Shell Dlg 2");
    ui->tableWidget->setRowCount(ui->tableWidget->rowCount() + 1);
    for (int i = 0; i < data.getGeneralData().size(); i++){
        //TODO: ОЧИЩАТЬ ПАМЯТЬ
        QTableWidgetItem* item = new QTableWidgetItem;
        item->setFont(font);
        item->setText(data.getGeneralData().at(i));
        ui->tableWidget->setItem(ui->tableWidget->rowCount() - 1, i, item);
  }
    QTableWidgetItem* item = new QTableWidgetItem;
    ui->tableWidget->setItem(ui->tableWidget->rowCount() - 1,
                             data.getGeneralData().size(), item);
    this->updateCountLabel();
}

void ReceivedDataDisplay::deleteButtonProcessing(){
    ui->tableWidget->removeRow(ui->tableWidget->currentRow());
    //TODO: ОЧИЩАТЬ ПАМЯТЬ ЯЧЕЕК
    this->updateCountLabel();
}

void ReceivedDataDisplay::deleteAllButtonProcessing(){
    int countOfRow = ui->tableWidget->rowCount();
    for (int i = 0; i < countOfRow; i++){
        ui->tableWidget->removeRow(0);
    }
    this->updateCountLabel();
}

void ReceivedDataDisplay::updateCountLabel(){
    ui->countLabel->setText(QString::fromUtf8("Количество: ")
                            + QString::number(ui->tableWidget->rowCount()));
}

QList<QStringList> ReceivedDataDisplay::getAllPassportExcelModels(){
    QList<QStringList> models;
    for (int i = 0; i < ui->tableWidget->rowCount(); i++){
        QStringList column;
        column << ui->tableWidget->item(i, 0)->text()
                  << ui->tableWidget->item(i, 1)->text()
                  << ui->tableWidget->item(i, 2)->text()
                  << ui->tableWidget->item(i, 3)->text()
                  << ui->tableWidget->item(i, 4)->text()
                  << ui->tableWidget->item(i, 5)->text() // сделать все в таком стиле
                  << ui->tableWidget->item(i, 6)->text();

        models.push_back(column);
    }
    return models;
}


ReceivedDataDisplay::~ReceivedDataDisplay(){
    delete ui;
}
