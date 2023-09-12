#include "bd_abityr.h"
#include "ui_bd_abityr.h"
#include <QDebug>

Bd_abityr::Bd_abityr(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Bd_abityr)
{
    ui->setupUi(this);
    //создание соединения
    this->setWindowTitle("База данных абитуриентов");
    dbb =QSqlDatabase::database("ConnectAb");
    if (!dbb.open()){
        QMessageBox msg;
        msg.setText("Ошибка отрытия базы данных");
        msg.exec();
    }
    //Отображение базы данных (Таблица)
    model = new QSqlTableModel(this,dbb);
    model->setTable("Abit_exam");
    model->select();
    ui->tableView->setModel(model);
    //Отображение заголовков
    model->setHeaderData(0,Qt::Horizontal,"ФИО \n абитуриента");
    model->setHeaderData(1,Qt::Horizontal,"Средний балл \n аттестата");
    model->setHeaderData(2,Qt::Horizontal,"Номер экзаменационного\nлиста");
    model->setHeaderData(3,Qt::Horizontal,"Наличие\nльгот");
    model->setHeaderData(4,Qt::Horizontal,"Оценка за\nэкзамен №1");
    model->setHeaderData(5,Qt::Horizontal,"Оценка за\nэкзамен №2");
    ui->tableView->resizeColumnsToContents();
}
//Дейструктор с закрытием подключения к БД
Bd_abityr::~Bd_abityr()
{
    dbb.close();
    delete ui;
}

void Bd_abityr::on_add_ab_btn_clicked()
{
    //Кнопка добавить
    model->insertRow(model->rowCount());
}

void Bd_abityr::on_del_ab_btn_clicked()
{
    //Кнопка удалить
     model->removeRow(row);
}

void Bd_abityr::on_tableView_clicked(const QModelIndex &index)
{
    //Получение номера удаляемой строки
    row = index.row();
}
