#include "form.h"
#include "ui_form.h"

Form::Form(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Form)
{
    ui->setupUi(this);
    this->update();
    this->resize(1353, 745);
    this->setWindowFlag(Qt::FramelessWindowHint);
    check1 = new QCheckBox();
    check2 = new QCheckBox();
    check3 = new QCheckBox();
    ui->tableWidget->setCellWidget(0,1,check1);
    ui->tableWidget->setCellWidget(1,1,check2);
    ui->tableWidget->setCellWidget(2,1,check3);
    check1->setStyleSheet(QLatin1String("QCheckBox::indicator { \n"
                                        "width: 21px;\n"
                                        "height: 21px;\n"
                                        "}\n"
                                        "QCheckBox{\n"
                                        "background-color: rgb(255, 255, 255);\n"
                                        "}"));
    check2->setStyleSheet(QLatin1String("QCheckBox::indicator { \n"
                                        "width: 21px;\n"
                                        "height: 21px;\n"
                                        "}\n"
                                        "QCheckBox{\n"
                                        "background-color: rgb(255, 255, 255);\n"
                                        "}"));
    check3->setStyleSheet(QLatin1String("QCheckBox::indicator { \n"
                                        "width: 21px;\n"
                                        "height: 21px;\n"
                                        "}\n"
                                        "QCheckBox{\n"
                                        "background-color: rgb(255, 255, 255);\n"
                                        "}"));
    ui->tableWidget->setColumnWidth(0,175);
    QObject::connect(check1,SIGNAL(clicked(bool)),this,SLOT(isSortNumber(bool)));
    QObject::connect(check2,SIGNAL(clicked(bool)),this,SLOT(isVoteRate(bool)));
    QObject::connect(check3,SIGNAL(clicked(bool)),this,SLOT(isFavour(bool)));
    QFont f = QFont("Ubuntu",16);
    ui->tableWidget_2->setFont(f);
    ui->tableWidget_2->setColumnWidth(1,250);
    ui->tableWidget_2->resizeColumnToContents(2);
    ui->tableWidget_3->setFont(f);
    ui->tableWidget_3->setColumnWidth(1,250);
    ui->tableWidget_3->resizeColumnToContents(2);
    QDesktopWidget *pDesk1 = QApplication::desktop();
    result = new ResultOutputWidget(pDesk1);
    result->move((pDesk1->width() - result->width()) / 2, (pDesk1->height() - result->height()) / 2);
    QObject::connect(result,SIGNAL(sendToMainWidget()),this,SLOT(dealResultSignal()));
}

void Form::isSortNumber(bool check){
    sort_number = check;
}

void Form::isVoteRate(bool check){
    vote_rate = check;
}

void Form::isFavour(bool check){
    favour = check;
}

void Form::mousePressEvent(QMouseEvent *p){
    if (p->button() == Qt::LeftButton){
        m_move = true;
        m_mouseStart=p->globalPos();
        m_windowStart = this->frameGeometry().topLeft();
    }
}
void Form::mouseMoveEvent(QMouseEvent *p){
    if (p->buttons() & Qt::LeftButton) {
        QPoint relativePos = p->globalPos() - m_mouseStart;
        this->move(m_windowStart + relativePos );
    }
}

void Form::mouseReleaseEvent(QMouseEvent *p){
    if (p->button() == Qt::LeftButton) {
        m_move = false;
    }
}
void Form::mouseDoubleClickEvent(QMouseEvent *){
    this ->showFullScreen();
}


void Form::GetTemplateWord(){
    QList<QTableWidgetItem*> items = ui->tableWidget_2->selectedItems();
    int item_count = items.count();
    QVector<int> line;
    QStringList mother_path;
    for(int i = 0;i < item_count;i = i+2){
        int row = ui->tableWidget_2->row(items.at(i));
        line.append(row);
    }
    for(int i = 0;i < line.count();i++){
        mother_path << ui->tableWidget_2->item(line[i],2)->text();
    }
    for(int i = 0;i < mother_path.count();i++){
        QFileInfo fileInfo(mother_path[i]);
        qDebug()<<fileInfo.fileName();
        if(fileInfo.fileName() == "多子类母版.docx"){
            QString mompath = fileInfo.filePath();
            QDir dir(fileInfo.path());
            dir.cdUp();
            QString tempath = dir.path() +  "//" +"templateWord"+"/多子类模板.docx";
            word.CopyFile(mompath,tempath);
            simple_output.creatMutiTemplate(sort_number,vote_rate,favour,tempath.toStdString());
        }
        else if(fileInfo.fileName() == "单子类母版.docx"){
            QString mompath = fileInfo.filePath();
            QDir dir(fileInfo.path());
            dir.cdUp();
            QString tempath = dir.path() +  "//" +"templateWord"+"/单子类模板.docx";
            word.CopyFile(mompath,tempath);
            simple_output.creatSimpleTemplate(sort_number,vote_rate,favour,tempath.toStdString());
        }
    }
    if(!mother_path.empty()){
        QFileInfo tempInfo(mother_path[0]);
        QDir temporary(tempInfo.path());
        if(!temporary.cdUp()){
            qDebug()<< "file not exist";
        }
        QString temp_path = temporary.path() +"/templateWord";
        QDir traversalFile(temp_path);
        //qDebug()<<traversalFile.path();
        QStringList filters;
        filters <<QString("*.docx");
        traversalFile.setNameFilters(filters);
        QList<QFileInfo> fileInfo(traversalFile.entryInfoList(filters));
        for(int i = 0;i < fileInfo.count();i++){
            ui->tableWidget_3->setRowCount(fileInfo.count());
            ui->tableWidget_3->setItem(i,1,new QTableWidgetItem(fileInfo[i].fileName()));
            ui->tableWidget_3->item(i,1)->setToolTip(fileInfo[i].fileName());
            ui->tableWidget_3->setItem(i,2,new QTableWidgetItem(fileInfo[i].filePath()));
            ui->tableWidget_3->item(i,2)->setToolTip(fileInfo[i].filePath());
        }
        int row_sum = ui->tableWidget_3->rowCount();
        int column_sum = ui->tableWidget_3->columnCount();
        for(int i = 0 ; i < row_sum; i++){
            for(int j = 0 ; j < column_sum; j++){
                if(ui->tableWidget_3->item(i,j) ==NULL)
                    continue;
                else
                    result->recvmsg(i,j,ui->tableWidget_3->item(i,j)->text(),row_sum,column_sum);
            }
        }
    }
}

void Form::GetMotherWord(){
    QStringList header;
    header << QStringLiteral("类型") << QStringLiteral("母版名") << QStringLiteral("母板路径");
    ui->tableWidget_2->setHorizontalHeaderLabels(header);
    QFileDialog *fileDialog = new QFileDialog(this);//创建一个QFileDialog对象，构造函数中的参数可以有所添加。
    fileDialog->setWindowTitle("SD Card");//设置对话框的标题
    fileDialog->setAcceptMode(QFileDialog::AcceptOpen);
    fileDialog->setFileMode(QFileDialog::ExistingFiles);//设置文件对话框弹出的时候显示任何文件，不论是文件夹还是文件
    fileDialog->setViewMode(QFileDialog::Detail);//文件以详细的形式显示，显示文件名，大小，创建日期等信息；
    fileDialog->setDirectory(".");//设置文件对话框打开时初始打开的位置
    fileDialog->setNameFilter(QString("所有文件(*.*);;.docx(*.docx*)"));
    fileDialog->setFixedSize(800,600);
    fileDialog->show();
    int m = ui->tableWidget_2->rowCount();
    if(fileDialog->exec() == QDialog::Accepted){
        QStringList list = fileDialog->selectedFiles();
        ui->tableWidget_2->setRowCount(list.size()+m);
        ui->tableWidget_2->setColumnCount(3);
        for(int i = 0;i < list.size();i++){
            QString path = list[i];
            ui->tableWidget_2->setItem(m+i,2,new QTableWidgetItem(path));
            ui->tableWidget_2->item(m+i,2)->setToolTip(path);
            QFileInfo fi = QFileInfo(path);
            ui->tableWidget_2->setItem(m+i,1,new QTableWidgetItem(fi.fileName()));
            ui->tableWidget_2->item(m+i,1)->setToolTip(fi.fileName());
        }
    }
    ui->widget->update();
}

void Form::DeleteTableWidget2Line(){
    QList<QTableWidgetItem*> delete_row = ui->tableWidget_2->selectedItems();
    int item_count = delete_row.count();
    QVector<int> line;
    for(int i = 0;i < item_count;i = i+2){
        line.append(ui->tableWidget_2->row(delete_row.at(i)));
    }
    for(int i = 0;i < line.count();i++){
        QFile delete_momword;
        delete_momword.remove(ui->tableWidget_2->item(line[i],2)->text());
    }
    for(int i = 0;i < item_count;i = i+2){
        ui->tableWidget_2->removeRow(ui->tableWidget_2->row(delete_row.at(i)));
    }
}
void Form::DeleteTableWidget3Line(){
    QList<QTableWidgetItem*> delete_row = ui->tableWidget_3->selectedItems(); //存储的每个单元格的地址
    int item_count = delete_row.count();
    QVector<int> line;
    for(int i = 0;i < item_count;i = i+2){
        line.append(ui->tableWidget_3->row(delete_row.at(i)));
    }
    /*QVector<QString> content;
    for(int i = 0;i < item_count;i = i+2){
        qDebug()<<delete_row.at(i)->text();
        content.append(delete_row.at(i)->text());
    }*/
    for(int i = 0;i < line.count();i++){
        QFile delete_momword;
        delete_momword.remove(ui->tableWidget_3->item(line[i],2)->text());
    }
    for(int i = 0;i < item_count;i = i+2){
        result->deletemsg(ui->tableWidget_3->row(delete_row.at(i)));   //需要将子窗口的删除函数放在父窗口的前面,否则父窗口删除数据会影响主窗口
        ui->tableWidget_3->removeRow(ui->tableWidget_3->row(delete_row.at(i)));
    }
}

void Form::OpenFile2(int row,int column){
    QDesktopServices::openUrl(QUrl::fromLocalFile(ui->tableWidget_2->item(row,2)->text()));
}

void Form::OpenFile3(int row,int column){
    QDesktopServices::openUrl(QUrl::fromLocalFile(ui->tableWidget_3->item(row,2)->text()));
}

void Form::ShowResultWidget(){
    this->hide();
    result->show();
}

void Form::dealResultSignal(){
    result->hide();
    this->show();
}


Form::~Form()
{
    delete ui;
    delete check1;
    delete check2;
    delete check3;
    delete result;
}
void Thread::run(){
    while(1){
        qDebug()<<"11111";
    }
}
