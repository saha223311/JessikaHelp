#ifndef RECEIVEDDATADISPLAY_H
#define RECEIVEDDATADISPLAY_H

#include <QWidget>
#include <QKeyEvent>

#include "fileprocessing.h" //ДЛЯ СТРУКТУРЫ, НАДО БУДЕТ СДЕЛАТЬ ОТДЕЛЬНЫЙ PassportExcelModel.h

namespace Ui {
class ReceivedDataDisplay;
}

class ReceivedDataDisplay : public QWidget
{
    Q_OBJECT

public:
    explicit ReceivedDataDisplay(QWidget *parent = 0);
    ~ReceivedDataDisplay();

protected:
    void keyPressEvent(QKeyEvent *);

public slots:
    void setMaximumValueForProgressBar(int value);
    void setCountOfRowsFromPassportExcelFile(int value);
    void setCountOfColsFromPassportExcelFile(int value);
    void displayStartDataProcessing();
    void displayEndDataProcessing();
    void setApplicationDirPath(QString path);
    void increaseProgressBar();

    void findButtonProcessing();
    void addPassportExcelModel(PassportExcelModel data);

    void deleteButtonProcessing();
    void deleteAllButtonProcessing();

signals:
    void needToGetPassportExcelModel(int index);

private:
    Ui::ReceivedDataDisplay *ui;

    int mCountOfRowsFromPassportExcelFile;
    int mCountOfColsFromPassportExcelFile;

    void updateCountLabel();
};

#endif // RECEIVEDDATADISPLAY_H
