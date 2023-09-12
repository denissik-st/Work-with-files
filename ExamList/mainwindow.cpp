#include "mainwindow.h"
#include "ui_mainwindow.h"


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->setWindowTitle("Работа с различными видами файлов.");
    //Создание БД при первом запуске
    db =QSqlDatabase::addDatabase("QODBC","ConnectAb");
    db.setDatabaseName("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};FIL={MS Access};DBQ=./Abiturients.accdb");
    if(!db.open()){
        QMessageBox msgError;
        msgError.setText("Ошибка открытия базы данных");
        msgError.exec();
    }
    //получение пути до файла
    QFileInfo fileExcel("Sort.xlsm");
    excelFile= fileExcel.absoluteFilePath();
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_bd_abityr_btn_clicked()
{
    // Кнопка перехода в режим работы с БД
    bd_abityr = new Bd_abityr;
    bd_abityr->setModal(true);
    bd_abityr->exec();
}

void MainWindow::on_sort_btn_clicked()
{
    //Подключение к файлу Excel, работа ведется на листе №1
    QAxObject *excel = new QAxObject("Excel.Application",this);
    QAxObject *workbooks = excel->querySubObject("Workbooks");
    QAxObject *workbook = workbooks->querySubObject( "Open(const QString&)",excelFile);
    QAxObject *mSheets = workbook->querySubObject( "Sheets" );
    QAxObject *sheet =mSheets->querySubObject("Item(int)",1);
    //Заполнение заголовка
    QAxObject *cell =sheet->querySubObject("Cells(int,int)",1,1);
    cell->dynamicCall("SetValue(const QString&)","ФИО Абитуриента");
    cell =sheet->querySubObject("Cells(int,int)",1,2);
    cell->dynamicCall("SetValue(const QString&)","Средний балл аттестата");
    cell =sheet->querySubObject("Cells(int,int)",1,3);
    cell->dynamicCall("SetValue(const QString&)","Номер экзаменационного листа");
    cell =sheet->querySubObject("Cells(int,int)",1,4);
    cell->dynamicCall("SetValue(const QString&)","Наличие льготы");
    cell =sheet->querySubObject("Cells(int,int)",1,5);
    cell->dynamicCall("SetValue(const QString&)","Оценка за экзамен №1");
    cell =sheet->querySubObject("Cells(int,int)",1,6);
    cell->dynamicCall("SetValue(const QString&)","Оценка за экзамен №2");

    //Запонение данными
    QSqlQuery query(db);
    db.open();
    if(query.exec("SELECT * FROM Abit_exam")){
        int row=2;
        while (query.next()) {
            cell =sheet->querySubObject("Cells(int,int)",row,1);
            cell->dynamicCall("SetValue(const QString&)",query.value(0).toString());
            cell =sheet->querySubObject("Cells(int,int)",row,2);
            cell->dynamicCall("SetValue(const QVarirnt&)",query.value(1).toFloat());
            cell =sheet->querySubObject("Cells(int,int)",row,3);
            cell->dynamicCall("SetValue(const QVarirnt&)",query.value(2).toInt());
            cell =sheet->querySubObject("Cells(int,int)",row,4);
            if (query.value(3).toBool()){
                cell->dynamicCall("SetValue(const QString&)","Есть");
            }
            else {
                cell->dynamicCall("SetValue(const QString&)","Нет");
            }
            cell =sheet->querySubObject("Cells(int,int)",row,5);
            cell->dynamicCall("SetValue(const QVarirnt&)",query.value(4).toInt());
            cell =sheet->querySubObject("Cells(int,int)",row,6);
            cell->dynamicCall("SetValue(const QVarirnt&)",query.value(5).toInt());
            row++;
        }
    }
    db.close();
    // Заполнение ячейки для выполнения макроса
    cell =sheet->querySubObject("Cells(int,int)",1,7);
    cell->dynamicCall("SetValue(const QVarirnt&)",1);
    //Чтение данных
    QVector<QString> FIO;
    QVector<float>GPA;
    QVector<int>Num_exam_list;
    QVector<QString> Benefit;
    QVector<int> exam_1;
    QVector<int>exam_2;
    cell =sheet->querySubObject("Cells(int,int)",1,1);
    int row=2;
    while (cell->dynamicCall("Value()").toString()!="") {
        cell =sheet->querySubObject("Cells(int,int)",row,1);
        FIO.append(cell->dynamicCall("Value()").toString());
        cell =sheet->querySubObject("Cells(int,int)",row,2);
        GPA.append(cell->dynamicCall("Value()").toFloat());
        cell =sheet->querySubObject("Cells(int,int)",row,3);
        Num_exam_list.append(cell->dynamicCall("Value()").toInt());
        cell =sheet->querySubObject("Cells(int,int)",row,4);
        Benefit.append(cell->dynamicCall("Value()").toString());
        cell =sheet->querySubObject("Cells(int,int)",row,5);
        exam_1.append(cell->dynamicCall("Value()").toInt());
        cell =sheet->querySubObject("Cells(int,int)",row,6);
        exam_2.append(cell->dynamicCall("Value()").toInt());
        row++;
    }

    //Сохранение файла и закрытие Excel (Освобождение памяти)
    delete cell;
    delete sheet;
    delete mSheets;
    workbook->dynamicCall("Save()");
    delete workbook;
    delete workbooks;
    excel->dynamicCall("Quit()");
    delete excel;


    //Отрытие файла Word
    QAxObject *word = new QAxObject("Word.Application");
    QAxObject *documents = word->querySubObject("Documents");
    QAxObject *document=documents->querySubObject("Add()");
    //Заголовок таблицы
    QAxObject *prangeZ= document->querySubObject("Range()");
    prangeZ->dynamicCall("SetRange(int, int)",0,70);
    prangeZ->setProperty("Text","Список абитуриентов, отсортированный по номеру экзаменационного билета");
    QAxObject *pfont = prangeZ->querySubObject("Font");
    pfont->setProperty("Name","Times New Roman");
    pfont->setProperty("Size",14);
    QAxObject *pformat = prangeZ->querySubObject("ParagraphFormat");
    pformat->setProperty("Alignment","wdAlignParagraphCenter");
    //Отрисовка таблицы
    QAxObject *prangeT =document->querySubObject("Range()");
    prangeT->dynamicCall("SetRange(int,int)",71,71);
    QAxObject *ptable = document->querySubObject("Tables()");
    QAxObject *pptable=ptable->querySubObject("Add(Range, NumRows As Long, NumColumns As Long,DefaultTableBehavior,AutoFitBehavior)",
                                              prangeT->asVariant(),row-2,6,1,2);
    //Заполнение заголовка таблицы
    QAxObject *wcell=pptable->querySubObject("Cell(Row, Column)",1,1);
    QAxObject *rangewcell = wcell->querySubObject("Range()");
    rangewcell->dynamicCall("InsertAfter(Text)", "ФИО Абитуриента");
    wcell=pptable->querySubObject("Cell(Row, Column)",1,2);
    rangewcell = wcell->querySubObject("Range()");
    rangewcell->dynamicCall("InsertAfter(Text)", "Средний балл аттестата");
    wcell=pptable->querySubObject("Cell(Row, Column)",1,3);
    rangewcell = wcell->querySubObject("Range()");
    rangewcell->dynamicCall("InsertAfter(Text)", "Номер экзаменационного листа");
    wcell=pptable->querySubObject("Cell(Row, Column)",1,4);
    rangewcell = wcell->querySubObject("Range()");
    rangewcell->dynamicCall("InsertAfter(Text)", "Наличие льготы");
    wcell=pptable->querySubObject("Cell(Row, Column)",1,5);
    rangewcell = wcell->querySubObject("Range()");
    rangewcell->dynamicCall("InsertAfter(Text)", "Оценка за экзамен №1");
    wcell=pptable->querySubObject("Cell(Row, Column)",1,6);
    rangewcell = wcell->querySubObject("Range()");
    rangewcell->dynamicCall("InsertAfter(Text)", "Оценка за экзамен №2");

    //Заполнение таблицы
    for (int i=2;i<=row-2;i++) {
        wcell=pptable->querySubObject("Cell(Row, Column)",i,1);
        rangewcell = wcell->querySubObject("Range()");
        rangewcell->dynamicCall("InsertAfter(Text)", FIO[i-2]);
        wcell=pptable->querySubObject("Cell(Row, Column)",i,2);
        rangewcell = wcell->querySubObject("Range()");
        rangewcell->dynamicCall("InsertAfter(Text)", GPA[i-2]);
        wcell=pptable->querySubObject("Cell(Row, Column)",i,3);
        rangewcell = wcell->querySubObject("Range()");
        rangewcell->dynamicCall("InsertAfter(Text)",Num_exam_list[i-2]);
        wcell=pptable->querySubObject("Cell(Row, Column)",i,4);
        rangewcell = wcell->querySubObject("Range()");
        rangewcell->dynamicCall("InsertAfter(Text)", Benefit[i-2]);
        wcell=pptable->querySubObject("Cell(Row, Column)",i,5);
        rangewcell = wcell->querySubObject("Range()");
        rangewcell->dynamicCall("InsertAfter(Text)", exam_1[i-2]);
        wcell=pptable->querySubObject("Cell(Row, Column)",i,6);
        rangewcell = wcell->querySubObject("Range()");
        rangewcell->dynamicCall("InsertAfter(Text)", exam_2[i-2]);
    }

    word->setProperty("Visible",true);
    //Освобождение памяти
    delete rangewcell;
    delete wcell;
    delete pfont;
    delete pformat;
    delete prangeZ;
    delete prangeT;
    delete pptable;
    delete ptable;
    QMessageBox::StandardButton reply = QMessageBox::question(this,"Печать файла","Расспечатать результаты сортировки?",
                                                              QMessageBox::Yes|QMessageBox::No);
    if(reply==QMessageBox::Yes){
        document->dynamicCall("PrintOut()");
    }
    delete document;
    delete documents;
    delete word;



}
