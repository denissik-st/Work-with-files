#include "mainwindow.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    // Функция запуска приложения
    QApplication a(argc, argv);
    MainWindow w;
    w.show();
    return a.exec();
}
