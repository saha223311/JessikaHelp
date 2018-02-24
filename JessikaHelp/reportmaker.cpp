#include "reportmaker.h"
#include <QDebug>
#include <QString>

ReportMaker::ReportMaker(QObject *parent) : QObject(parent)
{

}

//REWRITE!
void ReportMaker::makeReportWordLabel(QList<QStringList> data){

    // Поменять путь на нормальный!
    //как отдельную функцию в "каком-нибудь классе"

    QString newFile = "D:\\";
    newFile += dateTime.currentDateTime().toString("yyyy-MM-dd");
    newFile += "_" + dateTime.currentDateTime().toString("hh-mm");
    newFile += "_label.doc";

    QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();;
    QString applicationDirPath;

    for(int i=0;i< temporaryApplicationDirPath.length();i++){
        if(temporaryApplicationDirPath[i] == '/')  applicationDirPath+="\\";
        else applicationDirPath+=temporaryApplicationDirPath[i];
    }
    applicationDirPath+="\\template_label.doc";
    QFile::copy(applicationDirPath, newFile);


    QAxObject* mWord = new QAxObject( "Word.Application");
    mWord->querySubObject("Documents")->querySubObject( "Open(const QString&)", newFile);
    mWord->querySubObject("Selection")->querySubObject("Start", 0);
    mWord->querySubObject("Selection")->querySubObject("End", 0);

    int n = data.size(); // QTABLEWIDGET ROWS
     const int colsLabel = 4;
     const int rowsLabel = (n + colsLabel - 1) / colsLabel;
     const int pageWidth = 210 - 2;
     const int pageHeight = 297 - 2; // HACK "-2" работает
     const int marginLeft = 1;
     const int marginRight = 1;
     const int marginTop = 1;
     const int marginBottom = 1;
     const int widthCenter = (pageWidth - marginLeft - marginRight) / colsLabel;

     int t = 0;
     const double kMM = 110.0 / 38.8; // коэффициент для перевода мм в единицы ворда

     QAxObject* sel = mWord->querySubObject("Selection");
     QAxObject* tables = mWord->querySubObject("ActiveDocument")->querySubObject("Tables");
     QAxObject* newTable = tables->querySubObject("Add(Range, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                              mWord->querySubObject("ActiveDocument")->dynamicCall("Range()"), rowsLabel, colsLabel, 1, 1);

     QAxObject* cell;
     QAxObject* boarder;
     QString text;

     const int dotStyle = 4;

     int tempN = 0;
     for (int y = 1; y <=rowsLabel; y++)
     for(int x = 1; x <= colsLabel; ++x) {
         cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(y), QVariant(x)) ;
         cell->setProperty("Width",widthCenter * kMM);
         boarder = cell->querySubObject("Borders(xlEdge)", 1);
         boarder->setProperty("LineStyle",dotStyle);
         boarder = cell->querySubObject("Borders(xlEdge)", 2);
         boarder->setProperty("LineStyle",dotStyle);
         boarder = cell->querySubObject("Borders(xlEdge)", 3);
         boarder->setProperty("LineStyle",dotStyle);
         boarder = cell->querySubObject("Borders(xlEdge)", 4);
         boarder->setProperty("LineStyle",dotStyle);

         text = "";
        /* if (tempN < n){
            for(int i = 0; i < data.at(tempN).size(); i++){
                QString s = data.at(tempN).at(i);
                if (i == 1) s = "K-" + s;
                QString enter = s.isEmpty() ? "" : "\n";
                text = text + enter + s;
            }
         }*/
         if (tempN < n) text = "\n" + data.at(tempN).at(0) +
                 "\nK-" + data.at(tempN).at(1) +
                 "\n" + data.at(tempN).at(2) +
                 "\n" + data.at(tempN).at(3) +
                 "\n" + data.at(tempN).at(4) +
                 "\n" + data.at(tempN).at(5) +
                 "\n" + data.at(tempN).at(6);
        tempN++;
         cell->querySubObject("Range")->querySubObject("Text", text);
     }

     mWord->querySubObject("Documents")->dynamicCall("Save()");
     mWord->querySubObject("Documents")->dynamicCall("Close()");
     mWord->dynamicCall("SetDisplayAlerts(bool)", FALSE);
     mWord->dynamicCall("Quit()");
}
