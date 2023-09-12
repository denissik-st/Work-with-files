#ifndef BD_ABITYR_H
#define BD_ABITYR_H

#include <QDialog>
#include<QtSql/QSqlDatabase>
#include<QSqlTableModel>
#include<QMessageBox>

namespace Ui {
class Bd_abityr;
}

class Bd_abityr : public QDialog
{
    Q_OBJECT

public:
    explicit Bd_abityr(QWidget *parent = nullptr);
    ~Bd_abityr();

private slots:
    void on_add_ab_btn_clicked();

    void on_del_ab_btn_clicked();

    void on_tableView_clicked(const QModelIndex &index);

private:
    Ui::Bd_abityr *ui;
    QSqlDatabase dbb;
    QSqlTableModel *model;
    int row;
};

#endif // BD_ABITYR_H
