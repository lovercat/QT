#include "savetoexcel.h"
#include<QDir>
#include<QDebug>
#include<QMessageBox>

SaveToExcel::SaveToExcel()
{


}

SaveToExcel::~SaveToExcel()
{

}

void SaveToExcel::saveFile(QStringList str)
{
    if (!filepath.isEmpty()) {
        //qDebug("OK");
        QAxObject *excel = new QAxObject(this);
        excel->setControl("Excel.Application");//连接Excel控件
        excel->dynamicCall("SetVisible (bool Visible)", "false"); //不显示窗体
        excel->setProperty("DisplayAlerts",
                false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示

        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
        workbooks->dynamicCall("Add");//新建一个工作簿
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
        QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
        QAxObject *worksheet = worksheets->querySubObject("Item(int)", 1); //获取工作表集合的工作表1，即sheet1

        QAxObject *cellA, *cellB, *cellC, *cellD, *cellE, *cellF, *cellG, *cellH;

        //设置标题
        int cellrow = 1;
        QString A = "A" + QString::number(cellrow); //设置要操作的单元格，如A1
        QString B = "B" + QString::number(cellrow);
        QString C = "C" + QString::number(cellrow);
        QString D = "D" + QString::number(cellrow);
        QString E = "E" + QString::number(cellrow);
        QString F = "F" + QString::number(cellrow);
        QString G = "G" + QString::number(cellrow);
        QString H = "H" + QString::number(cellrow);
        cellA = worksheet->querySubObject("Range(QVariant, QVariant)", A); //获取单元格
        cellB = worksheet->querySubObject("Range(QVariant, QVariant)", B);
        cellC = worksheet->querySubObject("Range(QVariant, QVariant)", C);
        cellD = worksheet->querySubObject("Range(QVariant, QVariant)", D);
        cellE = worksheet->querySubObject("Range(QVariant, QVariant)", E); //获取单元格
        cellF = worksheet->querySubObject("Range(QVariant, QVariant)", F);
        cellG = worksheet->querySubObject("Range(QVariant, QVariant)", G);
        cellH = worksheet->querySubObject("Range(QVariant, QVariant)", H);
        cellA->dynamicCall("SetValue(const QVariant&)", QVariant("Name")); //设置单元格的值
        cellB->dynamicCall("SetValue(const QVariant&)", QVariant("Time(ms)"));
        cellC->dynamicCall("SetValue(const QVariant&)", QVariant("S0(mw)"));
        cellD->dynamicCall("SetValue(const QVariant&)", QVariant("S1(mw)"));
        cellE->dynamicCall("SetValue(const QVariant&)", QVariant("S2(mw)")); //设置单元格的值
        cellF->dynamicCall("SetValue(const QVariant&)", QVariant("S3(mw)"));
        cellG->dynamicCall("SetValue(const QVariant&)", QVariant("DOP(%)"));
        cellH->dynamicCall("SetValue(const QVariant&)", QVariant("Average_Dop(%)"));

        for (int i = 0; i < str.size(); i++) {
            QString A = "A" + QString::number(i + 2); //设置要操作的单元格，如A1
            QString B = "B" + QString::number(i + 2);
            QString C = "C" + QString::number(i + 2);
            QString D = "D" + QString::number(i + 2);
            QString E = "E" + QString::number(i + 2);
            QString F = "F" + QString::number(i + 2);
            QString G = "G" + QString::number(i + 2);
            QString H = "H" + QString::number(i + 2);
            cellA = worksheet->querySubObject("Range(QVariant, QVariant)", A); //获取单元格
            cellB = worksheet->querySubObject("Range(QVariant, QVariant)", B);
            cellC = worksheet->querySubObject("Range(QVariant, QVariant)", C);
            cellD = worksheet->querySubObject("Range(QVariant, QVariant)", D);
            cellE = worksheet->querySubObject("Range(QVariant, QVariant)", E); //获取单元格
            cellF = worksheet->querySubObject("Range(QVariant, QVariant)", F);
            cellG = worksheet->querySubObject("Range(QVariant, QVariant)", G);
            cellH = worksheet->querySubObject("Range(QVariant, QVariant)", H);
            cellA->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 1, 1))); //设置单元格的值
            cellB->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 2, 2)));
            cellC->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 3, 3)));
            cellD->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 4, 4)));
            cellE->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 5, 5))); //设置单元格的值
            cellF->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 6, 6)));
            cellG->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 7, 7)));
            cellH->dynamicCall("SetValue(const QVariant&)", QVariant(str.at(i).section(",", 8, 8)));
        }



        workbook->dynamicCall("SaveAs(const QString&)",
                  QDir::toNativeSeparators(
                        filepath)); //保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
        workbook->dynamicCall("Close()");//关闭工作簿
        excel->dynamicCall("Quit()");//关闭excel
        delete excel;
        excel = NULL;
        emit successful();
        //QMessageBox::about(this, "", "存入excel成功");
    } else {
        emit faile();
        //QMessageBox::about(this, "", "存入excel失败");
    }

}
