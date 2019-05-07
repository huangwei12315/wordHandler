#include "resultoutputwidget.h"
#include "ui_resultoutputwidget.h"

ResultOutputWidget::ResultOutputWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::ResultOutputWidget)
{
    ui->setupUi(this);
    this->update();
    this->resize(1353, 745);
    this->setWindowFlag(Qt::FramelessWindowHint);
    QObject::connect(ui->pushButton,SIGNAL(clicked(bool)),this,SLOT(dealSignals(bool)));
    QFont f = QFont("Ubuntu",16);
    ui->tableWidget_2->setFont(f);
    ui->tableWidget_2->setColumnWidth(1,250);
    ui->tableWidget_2->resizeColumnToContents(2);
}

void ResultOutputWidget::mousePressEvent(QMouseEvent *p){
    if (p->button() == Qt::LeftButton){
        m_move = true;
        m_mouseStart=p->globalPos();
        m_windowStart = this->frameGeometry().topLeft();
    }
}

void ResultOutputWidget::mouseMoveEvent(QMouseEvent *p){
    if (p->buttons() & Qt::LeftButton) {
        QPoint relativePos = p->globalPos() - m_mouseStart;
        this->move(m_windowStart + relativePos );
    }
}

void ResultOutputWidget::mouseReleaseEvent(QMouseEvent *p){
    if (p->button() == Qt::LeftButton) {
        m_move = false;
    }
}

void ResultOutputWidget::mouseDoubleClickEvent(QMouseEvent *p){
    this ->showFullScreen();
}

void ResultOutputWidget::dealSignals(bool){
    emit sendToMainWidget();
}

void ResultOutputWidget::recvmsg(int row,int column,QString msg,int row_sum,int column_sum){
   ui->tableWidget_2->setRowCount(row_sum);
   ui->tableWidget_2->setColumnCount(column_sum);
   ui->tableWidget_2->setItem(row,column,new QTableWidgetItem(msg));
   ui->tableWidget_2->item(row,column)->setToolTip(msg);
}

void ResultOutputWidget::openFile(int row,int column){
    QDesktopServices::openUrl(QUrl::fromLocalFile(ui->tableWidget_2->item(row,2)->text()));
}

void ResultOutputWidget::deletemsg(int delete_line){
   // qDebug()<<delete_line;
    ui->tableWidget_2->removeRow(delete_line);
}

ResultOutputWidget::~ResultOutputWidget()
{
    delete ui;
}
