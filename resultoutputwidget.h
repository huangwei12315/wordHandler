#ifndef RESULTOUTPUTWIDGET_H
#define RESULTOUTPUTWIDGET_H
#include <QTreeWidget>
#include <QWidget>
#include <QApplication>
#include <QPushButton>
#include <QString>
#include <QMouseEvent>
#include <QTableWidgetItem>
#include <QTableWidget>
#include <qnamespace.h>
#include <QHeaderView>
#include <QListWidget>
#include <QCheckBox>
#include <QFileDialog>
#include <QMessageBox>
#include <QFileInfo>
#include <QDebug>
#include <QCheckBox>
#include <QFont>
#include <QQueue>
#include <QVector>
#include <QDir>
#include <QScrollBar>
#include <QDesktopServices>
#include <QProcess>
#include "form.h"

namespace Ui {
class ResultOutputWidget;
}

class Form;

class ResultOutputWidget : public QWidget
{
    Q_OBJECT

public:
    void deletemsg(int delete_line);
    void recvmsg(int row,int column,QString msg,int row_sum,int column_sum);
    //void deletemsg(QList<QTableWidgetItem*> delete_row,int item_count);
    virtual void mouseDoubleClickEvent(QMouseEvent *p);
    virtual void mousePressEvent(QMouseEvent *p);
    virtual void mouseMoveEvent(QMouseEvent *p);
    virtual void mouseReleaseEvent(QMouseEvent *p);
    explicit ResultOutputWidget(QWidget *parent = 0);
    ~ResultOutputWidget();
signals:
    void sendToMainWidget();
public slots:
    void dealSignals(bool);
    void openFile(int row,int column);
private:
    Ui::ResultOutputWidget *ui;
    QPoint m_mouseStart;
    QPoint m_windowStart;
    bool m_move;

};

#endif // RESULTOUTPUTWIDGET_H
