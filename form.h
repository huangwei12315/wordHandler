#ifndef FORM_H
#define FORM_H
#include "wordHander.h"
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
#include "resultoutputwidget.h"
#include <QDesktopWidget>
#include "resutoutput.h"
#include <QThread>
using namespace std;
namespace Ui {
class Form;
}
class ResultOutputWidget;

class Form : public QWidget
{
    Q_OBJECT

public:
    void thread();
    virtual void mouseDoubleClickEvent(QMouseEvent *);
    virtual void mousePressEvent(QMouseEvent *p);
    virtual void mouseMoveEvent(QMouseEvent *p);
    virtual void mouseReleaseEvent(QMouseEvent *p);
    explicit Form(QWidget *parent = 0);
    ~Form();
signals:

public slots:
    void GetMotherWord();
    void GetTemplateWord();
    void DeleteTableWidget2Line();
    void DeleteTableWidget3Line();
    void isSortNumber(bool check);
    void isFavour(bool check);
    void isVoteRate(bool check);
    void OpenFile2(int row,int column);
    void OpenFile3(int row,int column);
    void ShowResultWidget();
    void dealResultSignal();
private:
    resutOutput simple_output;
    WordHandler word;
    QCheckBox *check1;
    QCheckBox *check2;
    QCheckBox *check3;
    QPoint m_mouseStart;
    QPoint m_windowStart;
    bool m_move;
    bool sort_number = false;
    bool favour = false;
    bool vote_rate = false;
    Ui::Form *ui;
    ResultOutputWidget  *result;
};

class Thread : public QThread{
public:
    void virtual run();

};
#endif // FORM_H
