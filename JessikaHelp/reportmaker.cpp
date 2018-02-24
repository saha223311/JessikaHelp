#include "reportmaker.h"
#include <QDebug>
#include <QString>

ReportMaker::ReportMaker(QObject *parent) : QObject(parent)
{

}

void ReportMaker::makeAllReports(const QList<QStringList> &data){
    makeReportWordLabel(data);
    makeReportWordEnvelope(data);
    //makeReportWordEnvelopeA4(data);
    //makeReportExcel(data);
}

//REWRITE!
void ReportMaker::makeReportWordLabel(const QList<QStringList> &data){

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
    mWord->querySubObject("Selection")->querySubObject("Start", 0); // ???????????
    mWord->querySubObject("Selection")->querySubObject("End", 0); // ?????????

    int n = data.size();
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

//REWRITE!
void ReportMaker::makeReportWordEnvelope(const QList<QStringList> &data){
    // Поменять путь на нормальный!
    //как отдельную функцию в "каком-нибудь классе"

    QString newFile = "D:\\";
    newFile += dateTime.currentDateTime().toString("yyyy-MM-dd");
    newFile += "_" + dateTime.currentDateTime().toString("hh-mm");
    newFile += "_envelope.doc";

    QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();;
    QString applicationDirPath;

    for(int i=0;i< temporaryApplicationDirPath.length();i++){
        if(temporaryApplicationDirPath[i] == '/')  applicationDirPath+="\\";
        else applicationDirPath+=temporaryApplicationDirPath[i];
    }
    applicationDirPath+="\\template_envelope.doc";
    QFile::copy(applicationDirPath, newFile);

    QAxObject* mWord = new QAxObject( "Word.Application");
    mWord->querySubObject("Documents")->querySubObject( "Open(const QString&)", newFile);
    mWord->querySubObject("Selection")->querySubObject("Start", 0);
    mWord->querySubObject("Selection")->querySubObject("End", 0);
    //
     const int pageWidth = 210 - 2, pageHeight = 297 - 2; // HACK "-2" работает
     const int marginLeft = 15, marginRight = 15;
     const int marginTop = 15, marginBottom = 15;
     //
     const int widthOverlap = 25;
     const int widthCenter = (pageWidth - widthOverlap) / 2;
     const int widthLeft = (pageWidth - widthCenter) / 2 - marginLeft;
     const int widthRight = (pageWidth - widthCenter) / 2 - widthOverlap;
     //
     const int heightOverlap = 20;
     const int heightSpace = 55;
     const int heightHalf = pageHeight / 2 ;
     const int heightText = heightHalf - heightSpace - heightOverlap;
     const int h1 = heightSpace - marginTop;
     const int h2txt = heightText;
     const int h3 = heightOverlap;
     const int h4 = heightSpace;
     const int h5txt = heightText;
     const int h6 = heightOverlap - marginBottom;
     //
     const int recordPerPage = 2;
     const int cellPerRecord = 3;
     const int cellPerPage = 6;
     //
     int t = 0;
     int n = data.size();

     const double kMM = 110.0 / 38.8; // коэффициент для перевода мм в единицы ворда

     QAxObject* sel = mWord->querySubObject("Selection");
     QAxObject* tables = mWord->querySubObject("ActiveDocument")->querySubObject("Tables");
     QAxObject* newTable = tables->querySubObject("Add(Range, NumRows, NumColumns, DefaultTableBehavior, AutoFitBehavior)",
                              mWord->querySubObject("ActiveDocument")->dynamicCall("Range()"),
                                                  cellPerPage * ((n + 1) / 2), 3, 1, 1);



       QAxObject* cell;
       QAxObject* boarder;
       QString text;

       const int dotStyle = 4;

       int tempN = 0;
       for (int row = 1; row <= (n + 1) / 2; row++){
        int ind = (row - 1)  * cellPerPage;

        for (int y = ind + 1; y <= ind + 6; ++y){
         for(int x = 1; x <= 3; ++x) {
           cell = newTable->querySubObject("Cell(Row, Column)" ,
                                           QVariant(y), QVariant(x));

           boarder = cell->querySubObject("Borders(xlEdge)", 4);
           boarder->setProperty("LineStyle",dotStyle);

           switch (y % 6) {
           case 1:
               cell->setProperty("Height", h1 * kMM);
               boarder = cell->querySubObject("Borders(xlEdge)", 1);
               boarder->setProperty("LineStyle",0);
               boarder = cell->querySubObject("Borders(xlEdge)", 3);
               boarder->setProperty("LineStyle",0);

               break;
           case 2:
               cell->setProperty("Height", h2txt * kMM);
               boarder = cell->querySubObject("Borders(xlEdge)", 1);
               boarder->setProperty("LineStyle",0);
               boarder = cell->querySubObject("Borders(xlEdge)", 3);
               boarder->setProperty("LineStyle",dotStyle);

               break;
           case 3:
               cell->setProperty("Height", h3 * kMM);
               boarder = cell->querySubObject("Borders(xlEdge)", 1);
               boarder->setProperty("LineStyle",dotStyle);
               boarder = cell->querySubObject("Borders(xlEdge)", 3);
               boarder->setProperty("LineStyle",1);
               break;
           case 4:
               cell->setProperty("Height", h4 * kMM);
               boarder = cell->querySubObject("Borders(xlEdge)", 3);
               boarder->setProperty("LineStyle",0);
               break;
           case 5:
               cell->setProperty("Height", h5txt * kMM);
               boarder = cell->querySubObject("Borders(xlEdge)", 1);
               boarder->setProperty("LineStyle",0);
               boarder = cell->querySubObject("Borders(xlEdge)", 3);
               boarder->setProperty("LineStyle",dotStyle);
               break;
           case 0:
               cell->setProperty("Height", h6 * kMM);
               boarder = cell->querySubObject("Borders(xlEdge)", 1);
               boarder->setProperty("LineStyle",dotStyle);
               boarder = cell->querySubObject("Borders(xlEdge)", 3);
               boarder->setProperty("LineStyle",0);
               break;

           default:
               break;
           }

           switch (x) {
           case 1:
               cell->setProperty("Width", widthLeft * kMM);
               boarder = cell->querySubObject("Borders(xlEdge)", 2);
               boarder->setProperty("LineStyle",0);

               break;
           case 2:
               cell->setProperty("Width", widthCenter * kMM);

               break;
           case 3:
               cell->setProperty("Width", widthRight * kMM);

               break;
           default:
               break;
           }

        if ((tempN < n) && (x == 2) &&(((y % 6) == 2) || ((y % 6) == 5))){

                   text = "\n" + data.at(tempN).at(0) +
                           "\nK-" + data.at(tempN).at(1) +
                           "\n" + data.at(tempN).at(2) +
                           "\n" + data.at(tempN).at(3) +
                           "\n" + data.at(tempN).at(4) +
                           "\n" + data.at(tempN).at(5) +
                           "\n" + data.at(tempN).at(6);
                   cell->querySubObject("Range")->querySubObject("Text", text);

                // SetTextToTable(t, y1 + ind + 1, x, FormatDateTime("yyyy-mm-dd hh:nn", iDateTime));
        tempN++;
        }



         }
        }
       }

       mWord->querySubObject("Documents")->dynamicCall("Save()");
       mWord->querySubObject("Documents")->dynamicCall("Close()");
       mWord->dynamicCall("SetDisplayAlerts(bool)", FALSE);
       mWord->dynamicCall("Quit()");

}
