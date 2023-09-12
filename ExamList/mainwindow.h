#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include<bd_abityr.h>
#include<QSqlDatabase>
#include<QSqlQuery>
#include <QFileInfo>
#include<QAxObject>
#include<QMessageBox>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    QSqlDatabase db;
    Bd_abityr *bd_abityr;
    QString excelFile;
    ~MainWindow();

private slots:
    void on_bd_abityr_btn_clicked();

    void on_sort_btn_clicked();

private:
    Ui::MainWindow *ui;
};
#endif // MAINWINDOW_H
