#include "reportmaker.h"
#include <QDebug>

ReportMaker::ReportMaker(QObject *parent) : QObject(parent)
{

}

void ReportMaker::makeReportWordLabel(QList<QStringList> data){

    QString temporaryApplicationDirPath = QCoreApplication::applicationDirPath();;
    QString applicationDirPath;

    for(int i=0;i< temporaryApplicationDirPath.length();i++){
        if(temporaryApplicationDirPath[i] == '/')  applicationDirPath+="\\";
        else applicationDirPath+=temporaryApplicationDirPath[i];
    }
    applicationDirPath+="\\template_label.doc";
    QFile::copy(applicationDirPath, "D:\\label.doc");

    QAxObject* mWord = new QAxObject( "Word.Application");
    mWord->querySubObject("Documents")->querySubObject( "Open(const QString&)", "D:\\label.doc");
    mWord->querySubObject("Selection")->querySubObject("Start", 0);
    mWord->querySubObject("Selection")->querySubObject("End", 0);

    int n = 5; // QTABLEWIDGET ROWS
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
     /*cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(1), QVariant(1)) ;
     cell->setProperty("Width",widthCenter * kMM);
     cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(1), QVariant(2)) ;
     cell->setProperty("Width",widthCenter * kMM);
     cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(1), QVariant(3)) ;
     cell->setProperty("Width",widthCenter * kMM);
     cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(1), QVariant(4)) ;
     cell->setProperty("Width",widthCenter * kMM);
     cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(2), QVariant(1)) ;
     cell->setProperty("Width",widthCenter * kMM);
     cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(2), QVariant(2)) ;
     cell->setProperty("Width",widthCenter * kMM);
     cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(2), QVariant(3)) ;
     cell->setProperty("Width",widthCenter * kMM);
     cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(2), QVariant(4)) ;
     cell->setProperty("Width",widthCenter * kMM);*/
/*
     const double kMM = 110.0 / 38.8; // коэффициент для перевода мм в единицы ворда
     for(int x = 1; x <= colsLabel; ++x) {
         mWord->querySubObject("ActiveDocument")->querySubObject("Tables")
                 ->querySubObject("Item(Table)", t)->querySubObject("Columns")
                 ->querySubObject("Item(Column)", x)->querySubObject("Width",  widthCenter * kMM);
       }
*/
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
         if (tempN < n) text = data.at(tempN).at(0) +
                 "\n" + data.at(tempN).at(1) +
                 "\n" + data.at(tempN).at(2) +
                 "\n" + data.at(tempN).at(3) +
                 "\n" + data.at(tempN).at(4) +
                 "\n" + data.at(tempN).at(5) +
                 "\n" + data.at(tempN).at(6);
        tempN++;
         cell->querySubObject("Range")->querySubObject("Text", text);
     }


   /* QAxObject* boarder;

     for(int row = 1; row <= n; ++row) {
         int x = ((row - 1) % colsLabel) + 1;
         int y = ((row - 1) / colsLabel) + 1;
         const int dotStyle = 4;


         cell = newTable->querySubObject("Cell(Row, Column)" , QVariant(y), QVariant(x)) ;
         boarder = cell->querySubObject("Borders(xlEdge)", 1);
         boarder->setProperty("LineStyle",dotStyle);
         boarder = cell->querySubObject("Borders(xlEdge)", 2);
         boarder->setProperty("LineStyle",dotStyle);
         boarder = cell->querySubObject("Borders(xlEdge)", 3);
         boarder->setProperty("LineStyle",dotStyle);
         boarder = cell->querySubObject("Borders(xlEdge)", 4);
         boarder->setProperty("LineStyle",dotStyle);

        //SetLineStyleBorderTable(t, y, x, true, true, true, true, dotStyle);
        // SetTextToTable(t, y, x, GetInfo(row, false));
       }
*/

     mWord->querySubObject("Documents")->dynamicCall("Save()");
     mWord->querySubObject("Documents")->dynamicCall("Close()");
     mWord->dynamicCall("SetDisplayAlerts(bool)", FALSE);
     mWord->dynamicCall("Quit()");


}
