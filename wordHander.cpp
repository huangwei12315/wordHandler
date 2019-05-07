#include "wordHander.h"





WordHandler::WordHandler(){
    this->word_cur =NULL;
    this->cur_save =NULL;
    this->cur_p_start =NULL;
    this->cur_p_end =NULL;
    this->cur_copy_paraph =NULL;
    this->cur_paraph_index =NULL;
    this->cur_delete_r =NULL;
    this->word_doc =NULL;
    this->word_namespace =NULL;
    this->attr_id_hide =NULL;
    this->attr_id_value =NULL;
    this->attr_delete_book =NULL;
    this->attr_change_book =NULL;
    this->attr_add_content =NULL;
    this->paraph_number =0;
    this->table_row_number =0;
    this->table_column_number =0;
    this->counter_start = 0;
    this->counter_end = 0;
    this ->change_copy = 0;
    memset(old_id_attr,0,sizeof(cur_copy));
}
/*
 *  函数名：initxml(xNode &cur_temp, xDoc &doc_temp , xNs* &name_temp)
 *  函数说明： 初始化xml节点
 *  函数参数: 1 cur_temp 节点
 *           2 doc_temp 文档节点
 *           3 name_temp 命名空间
 */
void WordHandler::Initxml(xNode &cur_temp, xDoc &doc_temp){

    xmlKeepBlanksDefault(0);
    try{
        doc_temp = xmlParseFile("document.xml");
    }catch(exception ex){
        cout << ex.what()<<endl;
    }
    try{
        cur_temp = xmlDocGetRootElement(doc_temp);
    }catch(exception ex){
        cout << ex.what()<< endl;
    }
    xmlNsPtr* name_list = NULL;
    name_list = xmlGetNsList(doc_temp,cur_temp);
    int index = 0;
    for(index = 0;name_list[index]->next !=NULL;index++){
        if(!xmlStrcmp((const xChar)"http://schemas.openxmlformats.org/wordprocessingml/2006/main",(const xChar)name_list[index]->href)){
            memcpy(&word_namespace,&name_list[index],sizeof(word_namespace));
        }
    }


    memcpy(&word_cur,&cur_temp,sizeof(word_cur));
    memcpy(&word_doc,&doc_temp,sizeof(word_doc));
}

void WordHandler::CopyFile(QString Qoldpath,QString Qnewpath){
    QFileInfo fi(Qnewpath);
    QFile remvfile;
    if(fi.exists()){
        remvfile.remove(Qnewpath);
    }
    QFile::copy(Qoldpath,Qnewpath);

}

/*
 * 函数名：AddBookmarkColumn(int tbl_number,int row_number_Start,int column_number,vector<string> &bookname, vector<string> &book_content,xNode cur)
 * 函数作用：向表格加入列书签
 * 函数参数：1.tbl_number  表示表格的索引号 以1为起始
 * 2. row_number_Start   表示添加列的起始行数索引号 以1为起始
 * 3. column_number      表示添加书签所在列数索引号 以1为起始
 * 4. bookname           字符串数组，表示添加书签的名称
 * 5. book_content       字符串数组,表示添加书签的内容
 * 6. xNode              节点，后面的与此参数相同意义
 * 函数说明：该函数用于向表格添加列书签，起始位置以行作为标准，在列里添加书签
 * 如 row_number_Start = 3; column_number = 2;表明添加第二列第三行之
 * 后的书签。第四个参数为添加书签的名称，同时因为是字符串数组，也可以进行个数
 * 限定，第五个参数也可限定个数
 */
void WordHandler::AddBookmarkColumn(int tbl_number,int row_number_Start,int column_number,vector<string> &bookname, vector<string> &book_content,xNode cur){
    /* 遍历文档树 */
    int tbl_number_temp = tbl_number + 1;
    unsigned int counter = 0;
    cur = cur->xmlChildrenNode;                                           //递归
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){                      //找到节点名为tbl的节点
            tbl_number--;
            if(!tbl_number){                                                     //找到相应索引号的tbl节点
                xNode cur1 = NULL;
                cur1 = cur ->xmlChildrenNode;
                while(cur1 != NULL){
                    int column_number_temp = column_number;                      //使用临时变量保存列数索引的值
                    if(!xmlStrcmp(cur1 -> name,(const xChar)"tr")){           //tr节点指的是行， 找到行节点
                        row_number_Start --;
                        if(row_number_Start <=0){                                 //找到起始行后面所有行的节点
                            if(counter < bookname.size())                         //  创建一个计数器，若计数器值小于需要
                                counter ++;                                       //  创建书签数组的大小贼继续添加书签，若大
                            else if(counter >= bookname.size())                   //  于创建书签数组的大小则程序结束，防止书签
                                break;                                            //  书签数小于单元格数继续向单元格添加书签
                            xNode cur2 = NULL;
                            cur2 = cur1 ->xmlChildrenNode;
                            while(cur2 != NULL){
                                if(!xmlStrcmp(cur2 -> name,(const xChar)"tc")){   //tc节点指列 找到列节点 列节点为行节点子节点
                                    column_number_temp --;
                                    if(!column_number_temp){                         //找到每行列的索引
                                        AddColumnBook(column_number,tbl_number_temp,counter,bookname,book_content,cur2);  //加入书签操作
                                    }
                                }
                                cur2 = cur2 -> next;
                            }
                        }
                    }
                    cur1 = cur1 -> next;
                }
            }
        }
        AddBookmarkColumn(tbl_number,row_number_Start,column_number,bookname,book_content,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}



/*向指定单元格加入书签
 * 参数1 列索引
 * 参数2 表格索引
 * 参数3 计数器 表示第几个书签添加到表格
 */
void WordHandler::AddColumnBook(int column_number,int tbl_number_temp,unsigned int counter,vector<string> &bookname, vector<string> &book_content, xNode cur2){
    string buf_temp_name = bookname[counter-1];
    const char *buf_bookName = buf_temp_name.c_str();                //将vector容器里的字符串提取出来 并转换成char*
    string buf_temp_cotent = book_content[counter-1];
    const char *buf_bookContent = buf_temp_cotent.c_str();
    cur2 = cur2 -> children;
    while(cur2 != NULL){
        if(!xmlStrcmp(cur2 -> name,(const xChar)"pPr")){         //找到pPr节点
            xNode cur_bookstart = NULL;                        //创建书签
            xNode cur_bookend =NULL;
            xNode cur_rpr = NULL;
            xNode cur_copy = NULL;
            xNode cur_r = NULL;
            xNode cur_t = NULL;
            xChar content = NULL;
            cur_rpr = cur2 -> children;
            while(cur_rpr != NULL){
                if(!xmlStrcmp(cur_rpr -> name,(const xChar)"rPr"))
                    cur_copy = xmlCopyNode(cur_rpr,1);
                cur_rpr = cur_rpr->next;
            }
            if(cur2 -> next == NULL){
                cur_bookstart = xmlNewNode(word_namespace,(const xChar)"bookmarkStart");
                xmlAddNextSibling(cur2,cur_bookstart);
                xmlSetNsProp(cur_bookstart,word_namespace,(const xChar)"name",(const xChar)buf_bookName);
                xmlSetNsProp(cur_bookstart,word_namespace,(const xChar)"id",(const xChar)buf_bookName);
                content = xmlGetProp(cur_bookstart,(const xChar)"id");
                cur_r = xmlNewNode(word_namespace,(const xChar)"r");
                xmlAddNextSibling(cur_bookstart,cur_r);
                xmlAddChild(cur_r,cur_copy);
                cur_t = xmlNewNode(word_namespace,(const xChar)"t");
                xmlAddNextSibling(cur_copy,cur_t);
                xmlNodeSetContent(cur_t,(const xChar)buf_bookContent);
                cur_bookend = xmlNewNode(word_namespace,(const xChar)"bookmarkEnd");
                xmlAddNextSibling(cur_r,cur_bookend);
                xmlSetNsProp(cur_bookend,word_namespace,(const xChar)"id",(const xChar)buf_bookName);
            }

        }
        AddColumnBook(column_number, tbl_number_temp,counter, bookname,book_content, cur2);
        cur2= cur2 -> next;
    }
}





/*
 *  函数名：AddBookmarkForTable(int tbl_number,int row_number_Start,int column_number_Start,,vector<string> &bookname, vector<string> &book_content,xNode cur)
 *  函数说明：向表格单元格内加入书签
 *  函数参数: 1.tbl_number    表示表格的索引号 以1为起始
 * 2. row_number           表示起始行数索引号 以1为起始
 * 3. column_number_Start  表示起始列数索引号 以1为起始
 * 4. bookname             字符串数组，表示添加书签的名称
 * 5. book_content         字符串数组,表示添加书签的内容
*/

void WordHandler::AddBookmarkForTable(int tbl_number,int row_number_Start,int column_number_Start,const string &bookname, const string &book_content,xNode cur){
    int tbl_number_temp = tbl_number;
    cur = cur->xmlChildrenNode;                                                //递归
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){                      //找到节点名为tbl的节点
            tbl_number--;
            if(!tbl_number){
                xNode cur1 = NULL;
                cur1 = cur ->xmlChildrenNode;
                while(cur1 != NULL){
                    if(!xmlStrcmp(cur1 -> name,(const xChar)"tr")){        //tr节点指的是行， 找到行节点
                        row_number_Start --;
                        if(!row_number_Start){
                            xNode cur2;
                            cur2 = cur1 -> xmlChildrenNode;
                            while(cur2 != NULL){
                                if(!xmlStrcmp(cur2 -> name,(const xChar)"tc")){   //tc节点指列 找到列节点 列节点为行节点子节点
                                    column_number_Start --;
                                    if(!column_number_Start){
                                        AddTableBook(column_number_Start,tbl_number_temp,bookname,book_content,cur2);
                                    }
                                }
                                cur2 = cur2 ->next;
                            }
                        }
                    }
                    cur1 = cur1 -> next;
                }
            }
        }

        AddBookmarkForTable(tbl_number,row_number_Start,column_number_Start,bookname,book_content,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}



void WordHandler::AddTableBook(int column_number_Start,int tbl_number_temp,const string &bookname, const string &book_content, xNode cur){
    const char *buf_bookName = bookname.c_str();
    const char *buf_bookContent = book_content.c_str();
    cur = cur -> children;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"pPr")){
            xNode cur_bookstart = NULL;
            xNode cur_bookend =NULL;
            xNode cur_rpr = NULL;
            xNode cur_copy = NULL;
            xNode cur_r = NULL;
            xNode cur_t = NULL;
            xChar content = NULL;
            cur_rpr = cur -> children;
            while(cur_rpr != NULL){
                if(!xmlStrcmp(cur_rpr -> name,(const xChar)"rPr"))
                    cur_copy = xmlCopyNode(cur_rpr,1);
                cur_rpr = cur_rpr->next;
            }
            if(cur -> next == NULL){
                cur_bookstart = xmlNewNode(word_namespace,(const xChar)"bookmarkStart");                            //向指定单元格添加书签
                xmlAddNextSibling(cur,cur_bookstart);
                xmlSetNsProp(cur_bookstart,word_namespace,(const xChar)"name",(const xChar)buf_bookName);
                xmlSetNsProp(cur_bookstart,word_namespace,(const xChar)"id",(const xChar)buf_bookName);
                content = xmlGetProp(cur_bookstart,(const xChar)"id");
                cur_r = xmlNewNode(word_namespace,(const xChar)"r");
                xmlAddNextSibling(cur_bookstart,cur_r);
                xmlAddChild(cur_r,cur_copy);
                cur_t = xmlNewNode(word_namespace,(const xChar)"t");
                xmlAddNextSibling(cur_copy,cur_t);
                xmlNodeSetContent(cur_t,(const xChar)buf_bookContent);
                cur_bookend = xmlNewNode(word_namespace,(const xChar)"bookmarkEnd");
                xmlAddNextSibling(cur_r,cur_bookend);
                xmlSetNsProp(cur_bookend,word_namespace,(const xChar)"id",(const xChar)buf_bookName);
            }

        }
        AddTableBook(column_number_Start, tbl_number_temp, bookname,book_content, cur);
        cur= cur -> next;
    }
}

void WordHandler::DeleteRowInBookmarkTable(const string &bookmark_name,int delete_row,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){   //找到书签start所在段落
            attr_add_content = xmlGetProp(cur,(const xChar)"id");
            cur_p_start = cur;
            while(cur_p_start !=NULL){
                if(!xmlStrcmp(cur_p_start ->name,(const xChar)"p"))
                    break;
                cur_p_start = cur_p_start ->parent;
            }
        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_add_content,(const xChar)xmlGetProp(cur,(const xChar)"id"))){                             //找到书签end所在段落
            cur_p_end = cur;
            if(!xmlStrcmp(cur_p_end->prev->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->prev->name,(const xChar)"p")){
                xNode cur_tbl =NULL;
                cur_tbl = cur_p_start;
                while(cur_tbl!=cur_p_end){
                    if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                        xmlNodePtr cur_temp =NULL;
                        cur_temp = cur_tbl ->children;
                        while(cur_temp != NULL){
                            if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                delete_row --;
                                if(!delete_row){
                                    xmlUnlinkNode(cur_temp);
                                    xmlFreeNode(cur_temp);
                                }
                            }
                            cur_temp = cur_temp ->next;
                        }
                    }
                    cur_tbl = cur_tbl->next;

                }
            }else{
                while(cur_p_end !=NULL){
                    if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->name,(const xChar)"p"))
                        break;

                    cur_p_end = cur_p_end ->parent;
                }

                if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")){
                    xmlNodePtr cur_temp =NULL;
                    cur_temp = cur_p_end ->children;
                    while(cur_temp != NULL){
                        if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                            delete_row --;
                            if(!delete_row){
                                xmlUnlinkNode(cur_temp);
                                xmlFreeNode(cur_temp);
                            }
                        }
                        cur_temp = cur_temp ->next;
                    }
                }

                else if(!xmlStrcmp(cur_p_end->name,(const xChar)"p")){
                    xNode cur_tbl =NULL;
                    cur_tbl = cur_p_start;
                    while(cur_tbl!=cur_p_end){
                        if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                            xmlNodePtr cur_temp =NULL;
                            cur_temp = cur_tbl ->children;
                            while(cur_temp != NULL){
                                if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                    delete_row --;
                                    if(!delete_row){
                                        xmlUnlinkNode(cur_temp);
                                        xmlFreeNode(cur_temp);
                                    }
                                }
                                cur_temp = cur_temp ->next;
                            }
                        }
                        cur_tbl = cur_tbl->next;
                    }
                }
            }
        }
        DeleteRowInBookmarkTable(bookmark_name, delete_row, cur);
        cur = cur -> next;
    }
}



void WordHandler::DeleteCloumnInBookmarkTable(const string &bookmark_name,int delete_column,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){   //找到书签start所在段落
            attr_add_content = xmlGetProp(cur,(const xChar)"id");
            cur_p_start = cur;
            while(cur_p_start !=NULL){
                if(!xmlStrcmp(cur_p_start ->name,(const xChar)"p"))
                    break;
                cur_p_start = cur_p_start ->parent;
            }
        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_add_content,(const xChar)xmlGetProp(cur,(const xChar)"id"))){                             //找到书签end所在段落
            cur_p_end = cur;
            if(!xmlStrcmp(cur_p_end->prev->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->prev->name,(const xChar)"p")){
                xNode cur_tbl =NULL;
                cur_tbl = cur_p_start;
                while(cur_tbl!=cur_p_end){
                    if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                        xmlNodePtr cur_temp =NULL;
                        cur_temp = cur_tbl ->children;
                        while(cur_temp != NULL){
                            if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                int delete_column_temp = delete_column;
                                xmlNodePtr cur_tc = cur_temp ->children;
                                while(cur_tc !=NULL){
                                    if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                        delete_column_temp --;
                                        if(!delete_column_temp){
                                            xmlUnlinkNode(cur_tc);
                                            xmlFreeNode(cur_tc);
                                        }
                                    }
                                    cur_tc = cur_tc ->next;
                                }
                            }
                            cur_temp = cur_temp ->next;
                        }
                    }
                    cur_tbl = cur_tbl->next;

                }
            }else{
                while(cur_p_end !=NULL){
                    if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->name,(const xChar)"p"))
                        break;

                    cur_p_end = cur_p_end ->parent;
                }

                if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")){
                    xmlNodePtr cur_temp =NULL;
                    cur_temp = cur_p_end ->children;
                    while(cur_temp != NULL){
                        if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                            int delete_column_temp = delete_column;
                            xmlNodePtr cur_tc = cur_temp ->children;
                            while(cur_tc !=NULL){
                                if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                    delete_column_temp --;
                                    if(!delete_column_temp){
                                        xmlUnlinkNode(cur_tc);
                                        xmlFreeNode(cur_tc);
                                    }
                                }
                                cur_tc = cur_tc ->next;
                            }
                        }
                        cur_temp = cur_temp ->next;
                    }
                }

                else if(!xmlStrcmp(cur_p_end->name,(const xChar)"p")){
                    xNode cur_tbl =NULL;
                    cur_tbl = cur_p_start;
                    while(cur_tbl!=cur_p_end){
                        if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                            xmlNodePtr cur_temp =NULL;
                            xmlNodePtr cur_last=NULL;
                            cur_temp = cur_tbl ->children;
                            while(cur_temp != NULL){
                                if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                    int delete_column_temp = delete_column;
                                    xmlNodePtr cur_copy=NULL;
                                    xmlNodePtr cur_tc = cur_temp ->children;
                                    while(cur_tc !=NULL){
                                        if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                            delete_column_temp --;
                                            if(!delete_column_temp){
                                                xmlUnlinkNode(cur_tc);
                                                xmlFreeNode(cur_tc);
                                            }
                                        }
                                        cur_tc = cur_tc ->next;
                                    }
                                    cur_last = xmlGetLastChild(cur_temp);
                                    xmlAddNextSibling(cur_last,cur_copy);
                                }
                                cur_temp = cur_temp ->next;
                            }
                        }
                        cur_tbl = cur_tbl->next;
                    }
                }

            }
        }
        DeleteCloumnInBookmarkTable(bookmark_name,delete_column, cur);
        cur = cur -> next;
    }
}




/*  函数名：AddCloumnInBookmarkTable(const string &bookmark_name,int copy_column,xNode cur)
 *  函数说明：向在书签内的表格复制任意一行到最后一行
 *  函数参数: 1 bookmark_name   表格所在书签
 *           2 copy_column      复制的行数
 *           3 cur            节点
 */
void WordHandler::AddCloumnInBookmarkTable(const string &bookmark_name,int copy_column,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){   //找到书签start所在段落
            attr_add_content = xmlGetProp(cur,(const xChar)"id");
            cur_p_start = cur;
            while(cur_p_start !=NULL){
                if(!xmlStrcmp(cur_p_start ->name,(const xChar)"p"))
                    break;
                cur_p_start = cur_p_start ->parent;
            }
        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_add_content,(const xChar)xmlGetProp(cur,(const xChar)"id"))){                             //找到书签end所在段落
            cur_p_end = cur;
            if(!xmlStrcmp(cur_p_end->prev->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->prev->name,(const xChar)"p")){
                xNode cur_tbl =NULL;
                cur_tbl = cur_p_start;
                while(cur_tbl!=cur_p_end){
                    if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                        xmlNodePtr cur_temp =NULL;
                        xmlNodePtr cur_last=NULL;
                        cur_temp = cur_tbl ->children;
                        while(cur_temp != NULL){
                            if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                int copy_column_temp = copy_column;
                                xmlNodePtr cur_copy=NULL;
                                xmlNodePtr cur_tc = cur_temp ->children;
                                while(cur_tc !=NULL){
                                    if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                        copy_column_temp --;
                                        if(!copy_column_temp){
                                            cur_copy = xmlCopyNode(cur_tc,1);    //复制行节点
                                        }
                                    }
                                    cur_tc = cur_tc ->next;
                                }
                                cur_last = xmlGetLastChild(cur_temp);
                                xmlAddNextSibling(cur_last,cur_copy);
                            }
                            cur_temp = cur_temp ->next;
                        }
                    }
                    cur_tbl = cur_tbl->next;

                }
            }else{
                while(cur_p_end !=NULL){
                    if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->name,(const xChar)"p"))
                        break;

                    cur_p_end = cur_p_end ->parent;
                }

                if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")){
                    xmlNodePtr cur_temp =NULL;
                    xmlNodePtr cur_last=NULL;
                    cur_temp = cur_p_end ->children;
                    while(cur_temp != NULL){
                        if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                            int copy_column_temp = copy_column;
                            xmlNodePtr cur_copy=NULL;
                            xmlNodePtr cur_tc = cur_temp ->children;
                            while(cur_tc !=NULL){
                                if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                    copy_column_temp --;
                                    if(!copy_column_temp){
                                        cur_copy = xmlCopyNode(cur_tc,1);    //复制行节点
                                    }
                                }
                                cur_tc = cur_tc ->next;
                            }
                            cur_last = xmlGetLastChild(cur_temp);
                            xmlAddNextSibling(cur_last,cur_copy);
                        }
                        cur_temp = cur_temp ->next;
                    }
                }

                else if(!xmlStrcmp(cur_p_end->name,(const xChar)"p")){
                    xNode cur_tbl =NULL;
                    cur_tbl = cur_p_start;
                    while(cur_tbl!=cur_p_end){
                        if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                            xmlNodePtr cur_temp =NULL;
                            xmlNodePtr cur_last=NULL;
                            cur_temp = cur_tbl ->children;
                            while(cur_temp != NULL){
                                if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                    int copy_column_temp = copy_column;
                                    xmlNodePtr cur_copy=NULL;
                                    xmlNodePtr cur_tc = cur_temp ->children;
                                    while(cur_tc !=NULL){
                                        if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                            copy_column_temp --;
                                            if(!copy_column_temp){
                                                cur_copy = xmlCopyNode(cur_tc,1);    //复制行节点
                                            }
                                        }
                                        cur_tc = cur_tc ->next;
                                    }
                                    cur_last = xmlGetLastChild(cur_temp);
                                    xmlAddNextSibling(cur_last,cur_copy);
                                }
                                cur_temp = cur_temp ->next;
                            }
                        }
                        cur_tbl = cur_tbl->next;
                    }
                }

            }
        }
        AddCloumnInBookmarkTable(bookmark_name,copy_column, cur);
        cur = cur -> next;
    }
}
/*  函数名：AddRowInBookmarkTable(const string &bookmark_name,int copy_line,xNode cur)
 *  函数说明：向在书签内的表格复制任意一行到最后一行
 *  函数参数: 1 bookmark_name   表格所在书签
 *           2 copy_line      复制的行数
 *           3 cur            节点
 */
void WordHandler::AddRowInBookmarkTable(const string &bookmark_name,int copy_line,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){   //找到书签start所在段落
            attr_add_content = xmlGetProp(cur,(const xChar)"id");
            cur_p_start = cur;
            while(cur_p_start !=NULL){
                if(!xmlStrcmp(cur_p_start ->name,(const xChar)"p"))
                    break;
                cur_p_start = cur_p_start ->parent;
            }
        }

        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_add_content,(const xChar)xmlGetProp(cur,(const xChar)"id"))){                             //找到书签end所在段落
            //cout <<bookmark_name.data()<<endl;
            cur_p_end = cur;
            if(!xmlStrcmp(cur_p_end->prev->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->prev->name,(const xChar)"p")){
                xNode cur_tbl =NULL;
                cur_tbl = cur_p_start;
                while(cur_tbl!=cur_p_end){
                    if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                        xmlNodePtr cur_temp =NULL;
                        xmlNodePtr cur_last=NULL;
                        xmlNodePtr cur_copy=NULL;
                        cur_temp = cur_tbl ->children;
                        while(cur_temp != NULL){
                            if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                copy_line --;
                                if(!copy_line){
                                    cur_copy = xmlCopyNode(cur_temp,1);    //复制行节点
                                }
                            }
                            cur_temp = cur_temp ->next;
                        }
                        cur_last = xmlGetLastChild(cur_tbl);              //得到表格最后一个子节点也就是行接点
                        xmlAddPrevSibling(cur_last,cur_copy);
                    }
                    cur_tbl = cur_tbl->next;

                }
            }else{
                while(cur_p_end !=NULL){
                    if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->name,(const xChar)"p"))
                        break;

                    cur_p_end = cur_p_end ->parent;
                }

                if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")){
                    //printf("11111\n");
                    xmlNodePtr cur_temp =NULL;
                    xmlNodePtr cur_last=NULL;
                    xmlNodePtr cur_copy=NULL;
                    cur_temp = cur_p_end ->children;
                    while(cur_temp != NULL){
                        if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                            copy_line --;
                            if(!copy_line){
                                cur_copy = xmlCopyNode(cur_temp,1);    //复制行节点
                            }
                        }
                        cur_temp = cur_temp ->next;
                    }
                    cur_last = xmlGetLastChild(cur_p_end);              //得到表格最后一个子节点也就是行接点
                    xmlAddPrevSibling(cur_last,cur_copy);
                }
                else if(!xmlStrcmp(cur_p_end->name,(const xChar)"p")){
                    xNode cur_tbl =NULL;
                    cur_tbl = cur_p_start;
                    while(cur_tbl!=cur_p_end){
                        if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                            xmlNodePtr cur_temp =NULL;
                            xmlNodePtr cur_last=NULL;
                            xmlNodePtr cur_copy=NULL;
                            cur_temp = cur_tbl ->children;
                            while(cur_temp != NULL){
                                if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                                    copy_line --;
                                    if(!copy_line){
                                        cur_copy = xmlCopyNode(cur_temp,1);    //复制行节点
                                    }
                                }
                                cur_temp = cur_temp ->next;
                            }
                            cur_last = xmlGetLastChild(cur_tbl);              //得到表格最后一个子节点也就是行接点
                            xmlAddNextSibling(cur_last,cur_copy);
                        }
                        cur_tbl = cur_tbl->next;
                    }
                }

            }
        }
        AddRowInBookmarkTable(bookmark_name,copy_line,cur);
        cur = cur -> next;
    }
}



/*  函数名：AddContentInBookmarkTable(const string &bookmark_name,int row_number,int column_number,const string &book_content,xNode cur)
 *  函数说明：向在书签内的表格加入内容
 *  函数参数: 1 bookmark_name   表格所在书签
 *           2 row_number      行号
 *           3 column_number   列号
 *           4 book_content    书签内容
 */
void WordHandler::AddContentInBookmarkTable(const string &bookmark_name,int row_number,int column_number,const string &book_content,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){   //找到书签start所在段落
            attr_add_content = xmlGetProp(cur,(const xChar)"id");
            cur_p_start = cur;
            while(cur_p_start !=NULL){
                if(!xmlStrcmp(cur_p_start ->name,(const xChar)"p"))
                    break;
                cur_p_start = cur_p_start ->parent;
            }
        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_add_content,(const xChar)xmlGetProp(cur,(const xChar)"id"))){                             //找到书签end所在段落
            cur_p_end = cur;
            if(!xmlStrcmp(cur_p_end->prev->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->prev->name,(const xChar)"p")){
                xNode cur_tbl =NULL;
                cur_tbl = cur_p_start;
                while(cur_tbl!=cur_p_end){
                    if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                        xNode cur_tr = NULL;
                        cur_tr = cur_tbl ->children;
                        while(cur_tr != NULL){
                            if(!xmlStrcmp(cur_tr->name,(const xChar)"tr")){
                                row_number--;
                                if(!row_number){
                                    xNode cur_tc = NULL;
                                    cur_tc = cur_tr ->children;
                                    while(cur_tc !=NULL){
                                        if(!xmlStrcmp(cur_tc->name,(const xChar)"tc")){
                                            column_number--;
                                            if(!column_number){
                                                AddContent(book_content,cur_tc);
                                            }
                                        }
                                        cur_tc = cur_tc->next;
                                    }
                                }
                            }
                            cur_tr = cur_tr->next;
                        }
                    }
                    cur_tbl = cur_tbl->next;
                }
            }else{
                while(cur_p_end !=NULL){
                    if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->name,(const xChar)"p"))
                        break;

                    cur_p_end = cur_p_end ->parent;
                }
                if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")){
                    xNode cur_tr = NULL;
                    cur_tr = cur_p_end ->children;
                    while(cur_tr != NULL){
                        if(!xmlStrcmp(cur_tr->name,(const xChar)"tr")){
                            row_number--;
                            if(!row_number){
                                xNode cur_tc = NULL;
                                cur_tc = cur_tr ->children;
                                while(cur_tc !=NULL){
                                    if(!xmlStrcmp(cur_tc->name,(const xChar)"tc")){
                                        column_number--;
                                        if(!column_number){
                                            AddContent(book_content,cur_tc);
                                        }
                                    }
                                    cur_tc = cur_tc->next;
                                }
                            }
                        }
                        cur_tr = cur_tr->next;
                    }
                }

                else if(!xmlStrcmp(cur_p_end->name,(const xChar)"p")){
                    xNode cur_tbl =NULL;
                    cur_tbl = cur_p_start;
                    while(cur_tbl!=cur_p_end){
                        if(!xmlStrcmp(cur_tbl->name,(const xChar)"tbl")){
                            xNode cur_tr = NULL;
                            cur_tr = cur_tbl ->children;
                            while(cur_tr != NULL){
                                if(!xmlStrcmp(cur_tr->name,(const xChar)"tr")){
                                    row_number--;
                                    if(!row_number){
                                        xNode cur_tc = NULL;
                                        cur_tc = cur_tr ->children;
                                        while(cur_tc !=NULL){
                                            if(!xmlStrcmp(cur_tc->name,(const xChar)"tc")){
                                                column_number--;
                                                if(!column_number){
                                                    AddContent(book_content,cur_tc);
                                                }
                                            }
                                            cur_tc = cur_tc->next;
                                        }
                                    }
                                }
                                cur_tr = cur_tr->next;
                            }
                        }
                        cur_tbl = cur_tbl->next;
                    }
                }

            }
        }
        AddContentInBookmarkTable(bookmark_name,row_number,column_number,book_content,cur);
        cur = cur -> next;
    }
}

void WordHandler::AddContent(const string &book_content,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur->name,(const xChar)"pPr")){
            xNode cur_rpr = NULL;
            xNode cur_copy =NULL;
            xNode cur_r = NULL;
            xNode cur_t = NULL;
            cur_rpr = cur ->children;
            while(cur_rpr != NULL){
                if(!xmlStrcmp(cur_rpr->name,(const xChar)"rPr")){
                    cur_copy = xmlCopyNode(cur_rpr,1);
                }
                cur_rpr = cur_rpr ->next;
            }
            cur_r = xmlNewNode(word_namespace,(const xChar)"r");
            xmlAddNextSibling(cur,cur_r);
            xmlAddChild(cur_r,cur_copy);
            cur_t = xmlNewNode(word_namespace,(const xChar)"t");
            xmlAddNextSibling(cur_copy,cur_t);
            xmlNodeSetContent(cur_t,(const xChar)book_content.data());

        }
        AddContent(book_content,cur);
        cur =cur ->next;
    }
}



//获取表格总数
int WordHandler:: GetTablesum(xNode cur){
    cur = cur->xmlChildrenNode->xmlChildrenNode;
    int table_number_cur = 0;
    while(cur !=NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            table_number_cur ++;
        }
        cur = cur -> next;
    }
    return table_number_cur;
}


/*
 *  函数名：CopyTable(int tbl_copy,xNode cur)
 *  函数说明：复制表格及表格上面的段落直到另一个表格，粘贴到其后面表格的后面
 *  函数参数: 1 tbl_number 表格索引
 *           2 tbb_sum 表格总数
 */
void WordHandler::CopyTableAndParaph(int tbl_copy,int tbl_sum,xNode cur){
    /* 遍历文档树 */
    cur = cur->xmlChildrenNode;
    int copy_tag_number = 0;
    int tbl_number_temp = tbl_copy;
    xNode cur2[100];
    xNode cur3[100];
    xNode cur_copy;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            tbl_sum --;
            tbl_copy --;
            if(tbl_copy == 0){
                cur_copy = xmlCopyNode(cur,1);
            }

            if(tbl_copy == 1){
                xNode cur_number;
                cur_number = cur->next;
                while(xmlStrcmp(cur_number->name,(const xChar)"tbl")){
                    copy_tag_number ++;
                    cur_number = cur_number -> next;
                }
                copy_tag_number = copy_tag_number+1;
                int MAX_cur;
                cur2[0] = cur->next;
                cur3[0]=xmlCopyNode(cur2[0],1);
                for(MAX_cur = 1;MAX_cur < copy_tag_number;MAX_cur++){
                    cur2[MAX_cur] = cur2[MAX_cur - 1]->next;
                    cur3[MAX_cur]=xmlCopyNode(cur2[MAX_cur],1);
                }
            }
        }

        if(tbl_sum ==0&&!xmlStrcmp(cur -> name,(const xChar)"tbl")&&tbl_number_temp != 1){
            int MAX_copy_cur;
            xmlAddNextSibling(cur,cur3[0]);
            for(MAX_copy_cur = 1;MAX_copy_cur < copy_tag_number;MAX_copy_cur ++){
                xmlAddNextSibling(cur3[MAX_copy_cur -1],cur3[MAX_copy_cur]);
            }
        }
        if(tbl_sum ==0&&!xmlStrcmp(cur -> name,(const xChar)"tbl")&&tbl_number_temp ==1){
            xmlAddNextSibling(cur,cur_copy);
        }
        CopyTableAndParaph(tbl_copy,tbl_sum,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}

/*
 *  函数名：CopyTable(int tbl_number_copy,int tbl_number_paste,bool position,xNode cur)
 *  函数说明：复制某个表格到另外一个表格的前或者后
 *  函数参数: 1 tbl_number_copy 复制的表格索引
 *           2 tbl_number_paste 粘贴表格索引
 *           3 position 位置  0为前 1为后
 *           4 cur 节点
 */

void WordHandler::CopyTable(int tbl_number_copy,int tbl_number_paste,bool position,xNode cur){
    xNode cur_copy =NULL;
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            if(tbl_number_copy < tbl_number_paste){
                tbl_number_copy--;
                tbl_number_paste--;
                if(!tbl_number_copy){
                    cur_copy = xmlCopyNode(cur,1);
                }
                if(!tbl_number_paste){
                    position?xmlAddNextSibling(cur,cur_copy):xmlAddPrevSibling(cur,cur_copy);    //如果复制索引小于粘贴索引，粘贴索引-复制索引向下寻找，找到后粘贴
                }
            }
            else if(tbl_number_copy > tbl_number_paste){
                tbl_number_copy--;
                tbl_number_paste--;
                if(!tbl_number_paste)
                    cur_save = cur;
                if(!tbl_number_copy)
                    cur_copy = xmlCopyNode(cur,1);
                position?xmlAddNextSibling(cur_save,cur_copy):xmlAddPrevSibling(cur_save,cur_copy);
            }                                                                                 //如果复制索引大于粘贴索引，复制索引-粘贴索引向上寻找，找到后粘贴
        }
        CopyTable(tbl_number_copy,tbl_number_paste,position,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}



/*
 *  函数名：CopyParaphByBookmark(const string &old_bookmark_name,const string &new_bookmark_name,const string &book_content,int tbl_number_paste,bool position,xNode cur)
 *  函数说明：复制书签所在段落并且修改书签
 *  函数参数: 1 old_bookmark_name 原来书签名
 *           2 new_bookmark_name 修改后书签名
 *           3 book_content 修改后书签内容
 *           4 tbl_number_paste 粘贴表格的位置
 *           5 position 位置  0为前 1为后
 *           6 cur 节点
 */

void WordHandler::CopyParaphByBookmark(const string &old_bookmark_name,const string &new_bookmark_name,const string &book_content,int tbl_number_paste,bool position,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)old_bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){
            xNode cur_p = cur;
            while (cur_p != NULL) {
                if(!xmlStrcmp(cur_p->name,(const xChar)"p"))
                    break;
                cur_p = cur_p ->parent;
            }
            cur_copy_paraph = xmlCopyNode(cur_p,1);
            ChangeCopy(old_bookmark_name,new_bookmark_name,book_content,cur_copy_paraph);
        }

        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            tbl_number_paste --;
            if(!tbl_number_paste){
                position?xmlAddNextSibling(cur,cur_copy_paraph):xmlAddPrevSibling(cur,cur_copy_paraph);                    //粘贴在表格前或后
            }
        }
        CopyParaphByBookmark(old_bookmark_name,new_bookmark_name,book_content, tbl_number_paste,position,cur);
        cur = cur -> next;
    }
}
void WordHandler::ChangeCopy(const string &old_bookmark_name,const string &new_bookmark_name,const string &book_content,xmlNodePtr cur){
    cur = cur->children;
    xChar attr = NULL;
    while(cur !=NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"bookmarkStart")&&!xmlStrcmp((const xChar)old_bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){
            xmlNodePtr cur_r;
            xmlNodePtr cur_t;
            xmlNodePtr cur_delete[20];
            attr = xmlGetProp(cur,(const xChar)"id");
            xmlSetNsProp(cur,word_namespace,(const xChar)"name",(const xChar)new_bookmark_name.data());
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)new_bookmark_name.data());
            if(!book_content.empty()){
            cur_r = cur->next;
            cur_delete[0] = cur_r ->next;
            for(int index = 0;xmlStrcmp(cur_delete[index]->name,(const xChar)"bookmarkEnd");index++){
                cur_delete[index+1]=cur_delete[index]->next;
                xmlUnlinkNode(cur_delete[index]);
                xmlFreeNode(cur_delete[index]);
            }

            cur_t = cur_r ->children;
            while(cur_t !=NULL){
                if(!xmlStrcmp(cur_t -> name,(const xChar)"t")){
                    if(!book_content.empty()){
                    xmlNodePtr cur_text;
                    cur_text = cur_t -> children;
                    xmlNodeSetContent(cur_t,(const xChar)book_content.data());
                    }
                    else
                        break;
                }
                cur_t = cur_t ->next;
            }
            }
        }

        if(!xmlStrcmp(cur -> name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)xmlGetProp(cur,(const xChar)"id"),(const xChar)attr)){
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)new_bookmark_name.data());
        }
        ChangeCopy(old_bookmark_name,new_bookmark_name,book_content,cur);
        cur =cur->next;
    }
}


/*
 *  函数名：CopyParaphByBookmark(const string &old_bookmark_name,const vector<string> &new_bookmark_name,const vector<string> &book_content,int tbl_number_paste,bool position,xNode cur)
 *  函数说明：复制书签所在段落并且修改书签
 *  函数参数: 1 old_bookmark_name 原来书签名
 *           2 new_bookmark_name 修改后书签名
 *           3 book_content 修改后书签内容
 *           4 tbl_number_paste 粘贴表格的位置
 *           5 position 位置  0为前 1为后
 *           6 cur 节点
 */
void WordHandler::CopyParaphByBookmark(const string &old_bookmark_name,const vector<string> &new_bookmark_name,const vector<string> &book_content,int tbl_number_paste,bool position,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)old_bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){
            xNode cur_p = cur;
            while (cur_p != NULL) {
                if(!xmlStrcmp(cur_p->name,(const xChar)"p"))
                    break;
                cur_p = cur_p ->parent;
            }
            cur_copy_paraph = xmlCopyNode(cur_p,1);
            ChangeCopy(new_bookmark_name,book_content,cur_copy_paraph);
        }

        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            tbl_number_paste --;
            if(!tbl_number_paste){
                position?xmlAddNextSibling(cur,cur_copy_paraph):xmlAddPrevSibling(cur,cur_copy_paraph);                    //粘贴在表格前或后
            }
        }
        CopyParaphByBookmark(old_bookmark_name,new_bookmark_name,book_content, tbl_number_paste,position,cur);
        cur = cur -> next;
    }
}
void WordHandler::ChangeCopy(const vector<string> &new_bookmark_name,const vector<string> &book_content,xmlNodePtr cur){
    cur = cur->children;
    int counter = 0;
    xChar attr = NULL;
    while(cur !=NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"bookmarkStart")){
            counter ++;
            string buf_temp_name = new_bookmark_name[counter-1];
            string buf_temp_content;
            if(!book_content.empty())
            buf_temp_content = book_content[counter-1];
            xmlNodePtr cur_r;
            xmlNodePtr cur_t;
            xmlNodePtr cur_delete[20];
            attr = xmlGetProp(cur,(const xChar)"id");
            xmlSetNsProp(cur,word_namespace,(const xChar)"name",(const xChar)buf_temp_name.data());
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)buf_temp_name.data());
            cur_r = cur->next;
            cur_delete[0] = cur_r ->next;
            for(int index = 0;xmlStrcmp(cur_delete[index]->name,(const xChar)"bookmarkEnd");index++){
                cur_delete[index+1]=cur_delete[index]->next;
                xmlUnlinkNode(cur_delete[index]);
                xmlFreeNode(cur_delete[index]);
            }
            cur_t = cur_r ->children;
            while(cur_t !=NULL){
                if(!xmlStrcmp(cur_t -> name,(const xChar)"t")){
                    if(!buf_temp_content.empty()){
                        xmlNodePtr cur_text;
                        cur_text = cur_t -> children;
                        xmlNodeSetContent(cur_t,(const xChar)buf_temp_content.data());
                    }
                    else
                        break;
                }
                cur_t = cur_t ->next;
            }
        }

        if(!xmlStrcmp(cur -> name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)xmlGetProp(cur,(const xChar)"id"),(const xChar)attr)){
            string buf_temp_name = new_bookmark_name[counter-1];
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)buf_temp_name.data());
        }
        ChangeCopy(new_bookmark_name,book_content,cur);
        cur =cur->next;
    }
}


void WordHandler::CopyParaphByBookmark1(const string &old_bookmark_name,const vector<string> &new_bookmark_name,int tbl_number_paste,xmlNodePtr cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        CopyParaphByBookmark1(old_bookmark_name,new_bookmark_name,tbl_number_paste,cur);
        if(!xmlStrcmp((const xChar)old_bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){   //找到书签start所在段落
            attr_add_content = xmlGetProp(cur,(const xChar)"id");
            cur_p_start = cur;
            while(cur_p_start !=NULL){
                if(!xmlStrcmp(cur_p_start ->name,(const xChar)"p"))
                    break;
                cur_p_start = cur_p_start ->parent;
            }
        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_add_content,(const xChar)xmlGetProp(cur,(const xChar)"id"))){                             //找到书签end所在段落
            cur_p_end = cur;
            if(!xmlStrcmp(cur_p_end->prev->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->prev->name,(const xChar)"p")){
                xNode cur_copy_temp  = NULL;
                int copy_index = 0;
                cur_copy[0] = cur_p_start;
                cur_copy_temp = cur_p_start;
                while(cur_copy_temp ->prev != cur_p_end){
                    cur_copy[copy_index] = xmlCopyNode(cur_copy_temp,1);
                    ChangeCopy(new_bookmark_name,cur_copy[copy_index]);
                    copy_index++;
                    cur_copy_temp = cur_copy_temp ->next;
                }
            }else{
                while(cur_p_end !=NULL){
                    if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")||!xmlStrcmp(cur_p_end->name,(const xChar)"p"))
                        break;

                    cur_p_end = cur_p_end ->parent;
                }
                if(!xmlStrcmp(cur_p_end->name,(const xChar)"tbl")){
                    xNode cur_copy_temp  = NULL;
                    int copy_index = 0;
                    cur_copy[0] = cur_p_start;
                    cur_copy_temp = cur_p_start;
                    while(cur_copy_temp ->prev != cur_p_end){
                        cur_copy[copy_index] = xmlCopyNode(cur_copy_temp,1);
                        ChangeCopy(new_bookmark_name,cur_copy[copy_index]);
                        copy_index++;
                        cur_copy_temp = cur_copy_temp ->next;
                    }
                }

                else if(!xmlStrcmp(cur_p_end->name,(const xChar)"p")){
                    xNode cur_copy_temp  = NULL;
                    int copy_index = 0;
                    cur_copy[0] = cur_p_start;
                    cur_copy_temp = cur_p_start;
                    while(cur_copy_temp ->prev != cur_p_end){
                        cur_copy[copy_index] = xmlCopyNode(cur_copy_temp,1);
                        ChangeCopy(new_bookmark_name,cur_copy[copy_index]);
                        copy_index++;
                        cur_copy_temp = cur_copy_temp ->next;
                    }
                }
            }
        }
        if(!xmlStrcmp(cur->name,(const xChar)"tbl")){
            tbl_number_paste--;
            if(!tbl_number_paste){
                xmlAddNextSibling(cur,cur_copy[0]);
                for(int index = 1; cur_copy[index] != NULL;index ++){
                    xmlAddNextSibling(cur_copy[index -1],cur_copy[index]);
                }
            }
        }

        cur = cur -> next;
    }
}




void WordHandler::ChangeCopy(const vector<string> &new_bookmark_name,xmlNodePtr cur){
    cur = cur ->children;
    while(cur !=NULL){
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkStart")){
            string buf_temp_name = new_bookmark_name[change_copy];
            old_id_attr[change_copy] = xmlGetProp(cur,(const xChar)"id");
            xmlSetNsProp(cur,word_namespace,(const xChar)"name",(const xChar)buf_temp_name.data());
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)buf_temp_name.data());
            change_copy ++;

        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")){
            for(int buf_temp_end = 0;buf_temp_end < change_copy;buf_temp_end++){
                if(!xmlStrcmp((const xChar)old_id_attr[buf_temp_end],xmlGetProp(cur,(const xChar)"id"))){
                    string buf_temp_name = new_bookmark_name[buf_temp_end];
                    xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)buf_temp_name.data());
                }
            }
        }
        cur = cur ->next;
    }
}

/*
 * 函数名：AddBookmarkForRow(int tbl_number,int row_number,int column_number_Start,vector<string> &bookname, vector<string> &book_content)
 * 函数作用：向表格加入行书签
 * 函数参数：1.tbl_number    表示表格的索引号 以1为起始
 * 2. row_number           表示添加书签起始行数索引号 以1为起始
 * 3. column_number_Start  表示起始列数索引号 以1为起始
 * 4. bookname             字符串数组，表示添加书签的名称
 * 5. book_content         字符串数组,表示添加书签的内容
 * 函数说明：该函数用于向表格添加列书签，起始位置以行作为标准，在列里添加书签
 * 如 row_number = 3; column_number_Start = 2;表明添加第三行第二列之
 * 后的书签。第四个参数为添加书签的名称，同时因为是字符串数组，也可以进行个数
 * 限定，第五个参数也可限定个数
 */
void WordHandler::AddBookmarkForRow(int tbl_number,int row_number,int column_number_Start,vector<string> &bookname, vector<string> &book_content,xNode cur){
    int tbl_number_temp = tbl_number;
    unsigned int counter = 0;
    cur = cur->xmlChildrenNode;                                                //递归
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){                      //找到节点名为tbl的节点
            tbl_number--;
            if(!tbl_number){
                xNode cur1 = NULL;
                cur1 = cur ->xmlChildrenNode;
                while(cur1 != NULL){
                    if(!xmlStrcmp(cur1 -> name,(const xChar)"tr")){        //tr节点指的是行， 找到行节点
                        row_number --;
                        if(!row_number){
                            xNode cur2;
                            cur2 = cur1 -> xmlChildrenNode;
                            while(cur2 != NULL){
                                if(!xmlStrcmp(cur2 -> name,(const xChar)"tc")){   //tc节点指列 找到列节点 列节点为行节点子节点
                                    column_number_Start --;
                                    if(column_number_Start <= 0){
                                        if(counter < bookname.size())                         //    创建一个计数器，若计数器值小于需要
                                            counter ++;                                       //  创建书签数组的大小贼继续添加书签，若大
                                        else
                                            break;
                                        AddRowBook(column_number_Start,tbl_number_temp,counter,bookname,book_content,cur2);
                                    }
                                }
                                cur2 = cur2 ->next;
                            }
                        }
                    }
                    cur1 = cur1 -> next;
                }
            }
        }
        AddBookmarkForRow(tbl_number,row_number,column_number_Start,bookname,book_content,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}

void WordHandler::AddRowBook(int column_number_Start,int tbl_number_temp,unsigned int counter,vector<string> &bookname, vector<string> &book_content, xNode cur2){
    string buf_temp_name = bookname[counter-1];
    const char *buf_bookName = buf_temp_name.c_str();
    string buf_temp_cotent = book_content[counter-1];
    const char *buf_bookContent = buf_temp_cotent.c_str();
    cur2 = cur2 -> children;
    while(cur2 != NULL){
        if(!xmlStrcmp(cur2 -> name,(const xChar)"pPr")){
            xNode cur_bookstart = NULL;
            xNode cur_bookend =NULL;
            xNode cur_rpr = NULL;
            xNode cur_copy = NULL;
            xNode cur_r = NULL;
            xNode cur_t = NULL;
            xChar content = NULL;
            cur_rpr = cur2 -> children;
            while(cur_rpr != NULL){
                if(!xmlStrcmp(cur_rpr -> name,(const xChar)"rPr"))
                    cur_copy = xmlCopyNode(cur_rpr,1);
                cur_rpr = cur_rpr->next;
            }
            if(cur2 -> next == NULL){
                cur_bookstart = xmlNewNode(word_namespace,(const xChar)"bookmarkStart");                            //向指定单元格添加书签
                xmlAddNextSibling(cur2,cur_bookstart);
                xmlSetNsProp(cur_bookstart,word_namespace,(const xChar)"name",(const xChar)buf_bookName);
                xmlSetNsProp(cur_bookstart,word_namespace,(const xChar)"id",(const xChar)buf_bookName);
                content = xmlGetProp(cur_bookstart,(const xChar)"id");
                cur_r = xmlNewNode(word_namespace,(const xChar)"r");
                xmlAddNextSibling(cur_bookstart,cur_r);
                xmlAddChild(cur_r,cur_copy);
                cur_t = xmlNewNode(word_namespace,(const xChar)"t");
                xmlAddNextSibling(cur_copy,cur_t);
                xmlNodeSetContent(cur_t,(const xChar)buf_bookContent);
                cur_bookend = xmlNewNode(word_namespace,(const xChar)"bookmarkEnd");
                xmlAddNextSibling(cur_r,cur_bookend);
                xmlSetNsProp(cur_bookend,word_namespace,(const xChar)"id",(const xChar)buf_bookName);
            }

        }
        AddRowBook(column_number_Start, tbl_number_temp,counter, bookname,book_content, cur2);
        cur2= cur2 -> next;
    }
}



/*
 *  函数名：AddRow(int tbl_number, int copy_line,int paste_line,xNode cur)
 *  函数说明：将指定行复制到第几行 无法将其它行复制到最后一行
 *  函数参数: 1 tbl_number表格索引
 *           2 copy_line 复制行数
 *           3 paste_line 需要粘贴行数
*/
void WordHandler::AddRow(int tbl_number,int copy_line,int paste_line,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            xNode cur_copy=NULL;
            int copy_line_temp = copy_line;
            tbl_number--;
            if(!tbl_number){
                xNode cur_tr;
                cur_tr = cur ->children;
                while(cur_tr != NULL){
                    if(!xmlStrcmp(cur_tr -> name,(const xChar)"tr")){   //找到行节点
                        copy_line --;
                        if(!copy_line){
                            cur_copy = xmlCopyNode(cur_tr,1);               //对此行复制
                            xNode cur_paste;
                            cur_paste = cur_tr;
                            if(paste_line > copy_line_temp){                //如果复制的行数小于粘贴的行数，两行位置相减，以复制的行数节点为基础向后面节点寻找并添加行
                                int data = paste_line - copy_line_temp;
                                while(data--)
                                    cur_paste = cur_paste->next;
                                if(cur_paste!=NULL)
                                    xmlAddPrevSibling(cur_paste,cur_copy);     //如果复制的行数等于粘贴的行数，直接向复制的节点前面粘贴一行
                            }
                            else if(paste_line == copy_line_temp){
                                xmlAddPrevSibling(cur_paste,cur_copy);
                            }
                            else if(paste_line < copy_line_temp){           //如果复制的行数大于粘贴的行数，两行位置相减，以复制的行数节点为基础向前面节点寻找并添加行
                                int data = copy_line_temp - paste_line;
                                while(data--)
                                    cur_paste = cur_paste->prev;
                                xmlAddPrevSibling(cur_paste,cur_copy);
                            }
                        }
                    }
                    cur_tr = cur_tr ->next;
                }
            }
        }
        AddRow(tbl_number,copy_line,paste_line,cur);
        cur = cur->next; /* 下一个子节点 */
    }

}



/*
 *  函数名：AddRow(int tbl_number, int copy_line,xNode cur)
 *  函数说明：将某行复制到最后一行
 *  函数参数: 1 tbl_number表格索引
 *           2 copy_line复制的行数
 *           2 xNode cur 节点
*/

void WordHandler::AddRow(int tbl_number,int copy_line,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            tbl_number--;
            if(!tbl_number){
                xmlNodePtr cur_temp =NULL;
                xmlNodePtr cur_last=NULL;
                xmlNodePtr cur_copy=NULL;
                cur_temp = cur ->children;
                while(cur_temp != NULL){
                    if(!xmlStrcmp(cur_temp -> name,(const xChar)"tr")){
                        copy_line --;
                        if(!copy_line){
                            cur_copy = xmlCopyNode(cur_temp,1);    //复制行节点
                        }
                    }
                    cur_temp = cur_temp ->next;
                }
                cur_last = xmlGetLastChild(cur);              //得到表格最后一个子节点也就是行接点
                xmlAddNextSibling(cur_last,cur_copy);         //将行节点粘贴到最后一个节点的下个节点

            }
        }
        AddRow(tbl_number,copy_line,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}


/*
 *  函数名：AddCloumn(int tbl_number,int copy_cloumn,int paste_cloumn,xNode cur)
 *  函数说明：将指定列复制到第几列，无法复制到最后一列
 *  函数参数: 1 tbl_number表格索引
 *           2 copy_cloumn, 复制列数
 *           3  paste_cloumn 需要粘贴列数
*/
void WordHandler::AddCloumn(int tbl_number,int copy_cloumn,int paste_cloumn,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            tbl_number--;
            if(!tbl_number){
                xNode cur_tr;
                cur_tr = cur ->children;
                while(cur_tr != NULL){
                    if(!xmlStrcmp(cur_tr -> name,(const xChar)"tr")){   //找到行节点
                        int copy_cloumn_temp = copy_cloumn;
                        xNode cur_tc;
                        xNode cur_copy=NULL;
                        int copy_column_temp = copy_cloumn;
                        cur_tc = cur_tr ->children;
                        while(cur_tc != NULL){
                            if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                copy_cloumn_temp --;
                                if(!copy_cloumn_temp){
                                    xNode cur_paste;
                                    cur_copy = xmlCopyNode(cur_tc,1);
                                    cur_paste = cur_tc;
                                    if(paste_cloumn > copy_column_temp){                //如果复制的行数小于粘贴的行数，两行位置相减，以复制的行数节点为基础向后面节点寻找并添加行
                                        int data = paste_cloumn - copy_column_temp;
                                        while(data--)
                                            cur_paste = cur_paste->next;
                                        if(cur_paste!=NULL)
                                            xmlAddPrevSibling(cur_paste,cur_copy);     //如果复制的行数等于粘贴的行数，直接向复制的节点前面粘贴一行
                                    }
                                    else if(paste_cloumn == copy_column_temp){
                                        xmlAddPrevSibling(cur_paste,cur_copy);
                                    }
                                    else if(paste_cloumn < copy_column_temp){           //如果复制的行数大于粘贴的行数，两行位置相减，以复制的行数节点为基础向前面节点寻找并添加行
                                        int data = copy_column_temp - paste_cloumn;
                                        while(data--)
                                            cur_paste = cur_paste->prev;
                                        xmlAddPrevSibling(cur_paste,cur_copy);
                                    }

                                }

                            }

                            cur_tc = cur_tc -> next;
                        }
                    }
                    cur_tr = cur_tr ->next;
                }
            }
        }
        AddCloumn(tbl_number,copy_cloumn,paste_cloumn,cur);
        cur = cur->next; /* 下一个子节点 */
    }

}



/*
 *  函数名：AddCloumn(int tbl_number,int copy_cloumn,xNode cur)
 *  函数说明：将某列复制到最后一列
 *  函数参数: 1 tbl_number表格索引
 *           2 复制的列数
 *           3 xNode cur 节点
*/
void WordHandler::AddCloumn(int tbl_number,int copy_cloumn,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            tbl_number--;
            if(!tbl_number){
                xmlNodePtr cur_tr =NULL;
                xmlNodePtr cur_last=NULL;
                xmlNodePtr cur_copy=NULL;
                cur_tr = cur ->children;
                while(cur_tr != NULL){
                    if(!xmlStrcmp(cur_tr -> name,(const xChar)"tr")){
                        int copy_cloumn_temp = copy_cloumn;
                        xmlNodePtr cur_tc =NULL;
                        cur_tc = cur_tr ->children;
                        while(cur_tc !=NULL){
                            if(!xmlStrcmp(cur_tc -> name,(const xChar)"tc")){
                                copy_cloumn_temp --;
                                if(!copy_cloumn_temp){
                                    cur_copy = xmlCopyNode(cur_tc,1);
                                    cur_last = xmlGetLastChild(cur_tr);
                                    xmlAddNextSibling(cur_last,cur_copy);
                                }
                            }
                            cur_tc = cur_tc -> next;
                        }
                    }
                    cur_tr = cur_tr ->next;
                }
            }
        }
        AddCloumn(tbl_number,copy_cloumn,cur);
        cur = cur->next;
    }
}

void WordHandler::AddEnterForTable(int tbl_number,bool position,string font,size_t size,xNode cur){
    cur =cur->children;
    while(cur !=NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){
            tbl_number --;
            if(!tbl_number){
                stringstream ss;
                string Font_size;
                size *= 2;
                ss << size;
                Font_size = ss.str();
                xNode cur_p = NULL;
                xNode cur_pPr = NULL;
                xNode cur_rPr = NULL;
                xNode cur_rFonts = NULL;
                xNode cur_b = NULL;
                xNode cur_sz = NULL;
                xNode cur_szCs = NULL;
                cur_p = xmlNewNode(word_namespace,(const xChar)"p");
                cur_pPr = xmlNewNode(word_namespace,(const xChar)"pPr");
                xmlAddChild(cur_p,cur_pPr);
                cur_rPr = xmlNewNode(word_namespace,(const xChar)"rPr");
                xmlAddChild(cur_pPr,cur_rPr);
                cur_rFonts = xmlNewNode(word_namespace,(const xChar)"rFonts");
                xmlSetNsProp(cur_rFonts,word_namespace,(const xChar)"ascii",(const xChar)font.data());
                xmlSetNsProp(cur_rFonts,word_namespace,(const xChar)"hAnsi",(const xChar)font.data());
                xmlSetNsProp(cur_rFonts,word_namespace,(const xChar)"eastAsia",(const xChar)font.data());
                xmlSetNsProp(cur_rFonts,word_namespace,(const xChar)"cs",(const xChar)font.data());
                xmlAddChild(cur_rPr,cur_rFonts);
                cur_b = xmlNewNode(word_namespace,(const xChar)"b");
                xmlAddNextSibling(cur_rFonts,cur_b);
                cur_sz= xmlNewNode(word_namespace,(const xChar)"sz");
                xmlSetNsProp(cur_sz,word_namespace,(const xChar)"val",(const xChar)Font_size.data());
                xmlAddNextSibling(cur_b,cur_sz);
                cur_szCs= xmlNewNode(word_namespace,(const xChar)"szCs");
                xmlSetNsProp(cur_szCs,word_namespace,(const xChar)"val",(const xChar)Font_size.data());
                xmlAddNextSibling(cur_sz,cur_szCs);
                position?xmlAddNextSibling(cur,cur_p):xmlAddPrevSibling(cur,cur_p);

            }
        }
        AddEnterForTable(tbl_number,position,font,size,cur);
        cur = cur->next;
    }
}




/*
 *  函数名：ChangeBookmarkContent(const string &bookmark_name,const string &change_content,bool delete_bookmark,xNode cur)
 *  函数说明：改变书签内容 可以选择是否删掉书签
 *  函数参数: 1 bookmark_name  书签名字
 *           2 change_content  现书签内容
 *           3 delete_bookmark 是否需要删除书签 1为删除书签 0则不删除
 *           4 cur 节点
*/

void WordHandler::ChangeBookmarkContent(const string &bookmark_name,const string &change_content,bool delete_bookmark,xNode cur){
    xChar attr_name_value = NULL;
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        ChangeBookmarkContent(bookmark_name,change_content,delete_bookmark,cur);
        if(!xmlStrcmp(cur->name, (const xChar)"bookmarkStart")){                //找到bookmark名字为GoBack的节点删掉该书签
            attr_name_value = xmlGetProp(cur,(const xChar)"name");
            if(!xmlStrcmp((const xChar)attr_name_value,(const xChar)bookmark_name.data())){     //是否删除书签并修改内容
                if(!delete_bookmark){
                    xNode cur_delete[20];
                    cur = cur->next;
                    if(!change_content.empty()){
                        ChangeContent(change_content,cur);
                    }else{
                        break;
                    }
                    cur_delete[0] = cur->next;
                    int index;
                    for(index = 0;xmlStrcmp(cur_delete[index]->name,(const xChar)"bookmarkEnd");index++){
                        cur_delete[index+1]=cur_delete[index]->next;
                        xmlUnlinkNode(cur_delete[index]);
                        xmlFreeNode(cur_delete[index]);
                    }
                }
                else{
                    xNode cur_delete[20];
                    xNode cur_value_start;
                    attr_id_value = xmlGetProp(cur,(const xChar)"id");
                    cur_value_start = cur;
                    cur = cur -> next;
                    xmlUnlinkNode(cur_value_start);
                    xmlFreeNode(cur_value_start);
                    if(!change_content.empty()){
                        ChangeContent(change_content,cur);
                    }else{
                        break;
                    }
                    cur_delete[0] = cur->next;
                    int index;
                    for(index = 0;xmlStrcmp(cur_delete[index]->name,(const xChar)"bookmarkEnd");index++){
                        cur_delete[index+1]=cur_delete[index]->next;
                        xmlUnlinkNode(cur_delete[index]);
                        xmlFreeNode(cur_delete[index]);
                    }
                }
            }
        }

        if(!xmlStrcmp(cur->name, (const xChar)"bookmarkEnd")){
            if(attr_id_value!=NULL&&!xmlStrcmp(attr_id_value,xmlGetProp(cur,(const xChar)"id"))){
                if(!delete_bookmark){
                    cur = cur->next;
                    continue;
                }
                else{
                    xNode cur_value_end;
                    cur_value_end = cur;
                    if(cur -> next != NULL)
                        cur = cur -> next;
                    xmlUnlinkNode(cur_value_end);
                    xmlFreeNode(cur_value_end);
                }
            }
        }

        cur = cur->next; /* 下一个子节点 */
    }
}




//修改书签内容
void WordHandler::ChangeContent(const string &change_content, xNode cur){
    cur =cur ->children;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"bookmarkEnd")&&!xmlStrcmp(cur -> name,(const xChar)"bookmarkStart"))
            cur = cur -> next;
        else if(!xmlStrcmp(cur -> name,(const xChar)"text")){
            xmlNodeSetContent(cur,(const xChar)change_content.data());
        }
        ChangeContent(change_content,cur);
        cur = cur -> next;
    }
}



/*
 *  函数名：ChangeBookmarkContent(const string &bookmark_name,const string &change_content,string font,size_t size,bool delete_bookmark,xNode cur)
 *  函数说明：改变书签内容 可以选择是否删掉书签
 *  函数参数: 1 bookmark_name  书签名字
 *           2 change_content  现书签内容
 *           3 font            字体
 *           4 size            字体大小
 *           5 delete_bookmark 是否需要删除书签 1为删除书签 0则不删除
 *           6 cur 节点
*/

void WordHandler::ChangeBookmarkContent(const string &bookmark_name,const string &change_content,string font,size_t size,bool delete_bookmark,xNode cur){
    xChar attr_name_value = NULL;
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        ChangeBookmarkContent(bookmark_name,change_content,font, size,delete_bookmark,cur);
        if(!xmlStrcmp(cur->name, (const xChar)"bookmarkStart")){                //找到bookmark名字为GoBack的节点删掉该书签
            attr_name_value = xmlGetProp(cur,(const xChar)"name");
            if(!xmlStrcmp((const xChar)attr_name_value,(const xChar)bookmark_name.data())){     //是否删除书签并修改内容
                if(!delete_bookmark){
                    cur = cur->next;
                    ChangeContent(change_content,font, size,cur);
                }
                else{
                    xNode cur_value_start;
                    attr_id_value = xmlGetProp(cur,(const xChar)"id");
                    cur_value_start = cur;
                    cur = cur -> next;
                    xmlUnlinkNode(cur_value_start);
                    xmlFreeNode(cur_value_start);
                    ChangeContent(change_content,font, size,cur);
                }
            }
        }

        if(!xmlStrcmp(cur->name, (const xChar)"bookmarkEnd")){
            if(attr_id_value!=NULL&&!xmlStrcmp(attr_id_value,xmlGetProp(cur,(const xChar)"id"))){
                if(!delete_bookmark){
                    cur = cur->next;
                    continue;
                }
                else{
                    xNode cur_value_end;
                    cur_value_end = cur;
                    if(cur -> next != NULL)
                        cur = cur -> next;
                    xmlUnlinkNode(cur_value_end);
                    xmlFreeNode(cur_value_end);
                }
            }
        }

        cur = cur->next; /* 下一个子节点 */
    }
}

void WordHandler::ChangeContent(const string &change_content,string font, size_t size ,xNode cur){
    cur = cur -> children;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"bookmarkEnd")&&!xmlStrcmp(cur -> name,(const xChar)"bookmarkStart"))
            cur = cur -> next;
        else if(!xmlStrcmp(cur -> name,(const xChar)"rFonts")){
            xmlSetNsProp(cur,word_namespace,(const xChar)"ascii",(const xChar)font.data());
            xmlSetNsProp(cur,word_namespace,(const xChar)"hAnsi",(const xChar)font.data());
            xmlSetNsProp(cur,word_namespace,(const xChar)"eastAsia",(const xChar)font.data());
            xmlSetNsProp(cur,word_namespace,(const xChar)"cs",(const xChar)font.data());
        }
        else if(!xmlStrcmp(cur -> name,(const xChar)"sz")){
            stringstream ss;
            string Font_size;
            size *= 2;
            ss << size;
            Font_size = ss.str();
            xmlSetNsProp(cur,word_namespace,(const xChar)"val",(const xChar)Font_size.data());
        }
        else if(!xmlStrcmp(cur -> name,(const xChar)"szCs")){
            stringstream ss;
            string Font_size;
            size *= 2;
            ss << size;
            Font_size = ss.str();
            xmlSetNsProp(cur,word_namespace,(const xChar)"val",(const xChar)Font_size.data());
        }
        else if(!xmlStrcmp(cur -> name,(const xChar)"text")){
            xmlNodeSetContent(cur,(const xChar)change_content.data());
        }
        ChangeContent(change_content, font, size,cur);
        cur = cur -> next;
    }
}


/*
 *  函数名：ChangeBookmarkId(const string &bookmark_id,const string &change_id,xNode cur)
 *  函数说明：修改书签id
 *  函数参数: 1 bookmark_id  书签id
 *           2 change_id  需改变的书签id
 *           3 cur 节点
*/
void WordHandler::ChangeBookmarkId(const string &bookmark_id,const string &change_id,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur ->name,(const xChar)"bookmarkStart")&&!xmlStrcmp((const xChar)bookmark_id.data(),(const xChar)xmlGetProp(cur,(const xChar)"id"))){
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)change_id.data());
        }

        ChangeBookmarkId(bookmark_id,change_id,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}



/*
 *  函数名：ChangeBookmarkname(const string &bookmark_name,const string &change_name,xNode cur,xmlNsPtr *NameSpace)
 *  函数说明：修改书签名（注意书签名必须英文字母开头）
 *  函数参数: 1 bookmark_name  书签名
 *           2 change_name  需修改的书签名
 *           3 cur 节点
 *           4 NameSpace 不同命名空间的集合
*/
void WordHandler::ChangeBookmarkname(const string &bookmark_name,const string &change_name,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur ->name,(const xChar)"bookmarkStart")&&!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){
            attr_id_value = xmlGetProp(cur,(const xChar)"id");
            xmlSetNsProp(cur,word_namespace,(const xChar)"name",(const xChar)change_name.data());
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)change_name.data());
        }
        if(!xmlStrcmp(cur ->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_id_value,(const xChar)xmlGetProp(cur,(const xChar)"id"))){
            xmlSetNsProp(cur,word_namespace,(const xChar)"id",(const xChar)change_name.data());
        }

        ChangeBookmarkname(bookmark_name,change_name,cur);
        cur = cur->next; /* 下一个子节点 */
    }
}

/*
 *  函数名：ChangeTablecontent(int tbl_number, int row_number, int column_number, const string &change_content, xNode cur);
 *  函数说明：修改表格中的内容
 *  函数参数: 1 tbl_number      表格索引
 *           2 row_number      行号索引
 *           3 column_number   列号索引
 *           4 content         改变的内容
 *           5 cur             节点
 */
void WordHandler::ChangeTablecontent(int tbl_number, int row_number, int column_number, const string &change_content, xNode cur){
    cur = cur->xmlChildrenNode;                                                //递归
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){                      //找到节点名为tbl的节点
            tbl_number--;
            if(!tbl_number){
                xNode cur1 = NULL;
                cur1 = cur ->xmlChildrenNode;
                while(cur1 != NULL){
                    if(!xmlStrcmp(cur1 -> name,(const xChar)"tr")){        //tr节点指的是行， 找到行节点
                        row_number --;
                        if(!row_number){
                            xNode cur2;
                            cur2 = cur1 -> xmlChildrenNode;
                            while(cur2 != NULL){
                                if(!xmlStrcmp(cur2 -> name,(const xChar)"tc")){   //tc节点指列 找到列节点 列节点为行节点子节点
                                    column_number --;
                                    if(!column_number){
                                        ChangeContent1(change_content, cur2);
                                    }
                                }
                                cur2 = cur2 ->next;
                            }
                        }
                    }
                    cur1 = cur1 -> next;
                }
            }
        }

        ChangeTablecontent(tbl_number, row_number, column_number,change_content,cur);
        cur = cur->next;/* 下一个子节点 */
    }
}

void WordHandler::ChangeContent1(const string &change_content, xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"text")){
            xmlNodeSetContent(cur,(const xChar)change_content.data());
        }
        ChangeContent1(change_content, cur);
        cur = cur ->next;
    }
}
/*
 *  函数名：DeleteHideBookmark()
 *  函数说明：删除隐藏书签
 *  函数参数: 无
 */
void WordHandler::DeleteHideBookmark(xNode cur){
    xChar attr_name_value = NULL;
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur->name, (const xChar)"bookmarkStart")){                //找到bookmark名字为GoBack的节点删掉该书签
            attr_name_value = xmlGetProp(cur,(const xChar)"name");
            if(!xmlStrcmp(attr_name_value,(const xChar)"_GoBack")){
                xNode cur_hide_start;
                attr_id_hide = xmlGetProp(cur,(const xChar)"id");
                cur_hide_start = cur;
                cur = cur -> next;
                xmlUnlinkNode(cur_hide_start);
                xmlFreeNode(cur_hide_start);
            }

        }

        if(!xmlStrcmp(cur->name, (const xChar)"bookmarkEnd")){
            if(attr_id_hide!=NULL&&!xmlStrcmp(attr_id_hide,xmlGetProp(cur,(const xChar)"id"))){
                xNode cur_hide_end;
                cur_hide_end = cur;
                if(cur -> next != NULL)
                    cur = cur -> next;
                xmlUnlinkNode(cur_hide_end);
                xmlFreeNode(cur_hide_end);
            }

        }
        DeleteHideBookmark(cur);
        cur = cur->next; /* 下一个子节点 */
    }
}





/*
 *  函数名：DeleteBookmarkForId(const string ID_number,xNode cur)
 *  函数说明：用id删除书签
 *  函数参数: 1 ID_number ID号
 *           2 cur 节点
 */
void WordHandler::DeleteBookmarkForId(const string ID_number,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkStart")&&!xmlStrcmp((const xChar)ID_number.data(),(const xChar)xmlGetProp(cur,(const xChar)"id"))){
            xNode cur_start;
            cur_start = cur;
            if(cur -> next !=NULL)
                cur = cur ->next;
            xmlUnlinkNode(cur_start);
            xmlFreeNode(cur_start);
        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)ID_number.data(),(const xChar)xmlGetProp(cur,(const xChar)"id"))){
            xNode cur_end;
            cur_end = cur;
            if(cur -> next !=NULL)
                cur = cur ->next;
            xmlUnlinkNode(cur_end);
            xmlFreeNode(cur_end);
        }
        DeleteBookmarkForId(ID_number,cur);
        cur = cur->next;

    }
}


/*
 *  函数名：DeleteBookmarkInParaph(const string &bookmark_name,xNode cur)
 *  函数说明：用书签删除段落
 *  函数参数: 1 bookmark_name 书签名
 *           2 cur 节点
 */
void WordHandler::DeleteBookmarkInParaph(const string &bookmark_name,xNode cur){
    /* 遍历文档树 */
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"p")){           //找到段落p
            DeleteParaph(bookmark_name,cur);
        }
        DeleteBookmarkInParaph(bookmark_name,cur);
        cur = cur -> next;
    }
}


//对段落进行删除
void WordHandler::DeleteParaph(const string &bookmark_name,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur ->name,(const xChar)"bookmarkStart")&&!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){
            xNode cur_p;
            cur_p = cur;
            while(cur_p != NULL){
                if(!xmlStrcmp(cur_p ->name,(const xChar)"p")){                     //用书签找到段落p并且删除此段落
                    xmlUnlinkNode(cur_p);
                    xmlFreeNode(cur_p);
                }
                cur_p = cur_p ->parent;
            }
        }
        DeleteParaph(bookmark_name,cur);
        cur = cur->next;
    }
}



/*
 *  函数名：DeleteBookmarkInParaph1(const string &bookmark_name,xNode cur)
 *  函数说明：用书签删除段落 跨行一并删除
 *  函数参数: 1 bookmark_name 书签名
 *           2 cur 节点
 */
void WordHandler::DeleteBookmarkInParaph1(const string &bookmark_name,xNode cur){
    xNode cur_delete[100];
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp((const xChar)bookmark_name.data(),(const xChar)xmlGetProp(cur,(const xChar)"name"))){   //找到书签start所咋段落
            attr_delete_book = xmlGetProp(cur,(const xChar)"id");
            cur_p_start = cur;
            while(cur_p_start !=NULL){
                if(!xmlStrcmp(cur_p_start ->name,(const xChar)"p"))
                    break;
                cur_p_start = cur_p_start ->parent;
            }
        }
        if(!xmlStrcmp(cur->name,(const xChar)"bookmarkEnd")&&!xmlStrcmp((const xChar)attr_delete_book,(const xChar)xmlGetProp(cur,(const xChar)"id"))){
            cur_p_end = cur;                                                                                          //找到书签end所在段落
            if(cur_p_end -> parent !=NULL){
                if(!xmlStrcmp(cur_p_end ->prev->name,(const xChar)"p"))
                    cur_p_end = cur;
                else{
                    while(cur_p_end != NULL){
                        if(!xmlStrcmp(cur_p_end ->name,(const xChar)"p")||!xmlStrcmp(cur_p_end ->name,(const xChar)"tbl"))
                            break;
                        cur_p_end = cur_p_end ->parent;
                    }
                }
            }
            int counter = 0;                                                                                 //找到两段落后循环删除
            cur_delete[0] = cur_p_start;
            for(counter = 0;cur_delete[counter]!=cur_p_end;counter++){
                cur_delete[counter+1] = cur_delete[counter] ->next;
                xmlUnlinkNode(cur_delete[counter]);
                xmlFreeNode(cur_delete[counter]);
            }
            if(cur_delete[counter] == cur_p_end){
                xmlUnlinkNode(cur_delete[counter]);
                xmlFreeNode(cur_delete[counter]);
            }
        }

        DeleteBookmarkInParaph1(bookmark_name,cur);
        cur = cur -> next;
    }
}



/*
 *  函数名：DeleteColumn(int tbl_number,int column_number,xNode cur)
 *  函数说明：删除指定列
 *  函数参数: 1 tbl_number 表格索引
 *           2 column_number 列索引
 *           3 cur 节点
 */
void WordHandler::DeleteColumn(int tbl_number,int column_number,xNode cur){
    /* 遍历文档树 */
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){                //找到表格节点
            tbl_number--;
            if(!tbl_number){
                DeleteC(column_number,cur);
            }
        }
        DeleteColumn(tbl_number,column_number,cur);
        cur = cur->next; /* 下一个子节点 */
    }

}




//删除表格列操作
void WordHandler::DeleteC(int column_number, xNode cur){
    cur = cur -> children;                                                         //在指定表格节点进行递归
    while(cur != NULL){
        if(cur != NULL&&! xmlStrcmp(cur -> name,(const xChar)"tr")){     //找到行节点
            int column_number_temporary;
            column_number_temporary = column_number;
            xNode cur1;
            cur1 = cur;
            cur1 = cur1 -> children;
            while(cur1 != NULL){
                if(!xmlStrcmp(cur1 -> name,(const xChar)"tc")){            //找到行节点的子节点列节点
                    column_number_temporary--;
                    if(!column_number_temporary&&cur1->next!=NULL){                              //删除每个行节点指定的列节点
                        xNode cur2;
                        cur2 = cur1;
                        cur1 = cur1 -> next;
                        xmlUnlinkNode(cur2);
                        xmlFreeNode(cur2);
                    }
                    if(!column_number_temporary&&cur1->next==NULL){
                        xNode cur2;
                        cur2 = cur1;
                        xmlUnlinkNode(cur2);
                        xmlFreeNode(cur2);
                    }
                }
                cur1 = cur1->next;
            }
        }
        DeleteC(column_number,cur);
        cur = cur -> next;
    }
}



/*
 *  函数名：DeleteRow(int tbl_number,int row_number,xNode cur)
 *  函数说明：删除指定行
 *  函数参数: 1 tbl_number 表格索引
 *           2 column_number 列索引
 *           3 cur 节点
 */
void WordHandler::DeleteRow(int tbl_number,int row_number,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){           //找到指定表格
            tbl_number--;
            if(!tbl_number){
                DeleteR(row_number,cur);
            }
        }
        DeleteRow(tbl_number,row_number,cur);
        cur = cur->next; /* 下一个子节点 */
    }

}


//删除列操作
void  WordHandler::DeleteR(int row_number, xNode cur){
    cur = cur -> children;                                  //对表格进行遍历
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tr")){   //对指定行节点进行删除
            row_number --;
            if(!row_number){
                xNode cur_tr;
                cur_tr = cur;
                cur = cur->next;
                xmlUnlinkNode(cur_tr);
                xmlFreeNode(cur_tr);
                continue;
            }
        }
        DeleteR(row_number,cur);
        cur = cur -> next;
    }
}


/*
 *  函数名：GetTableColumnNumber(int tbl_number,xNode cur)
 *  函数说明：得到指定表格的列数
 *  函数参数:1 tbl_number 表格索引
 *          2 cur        节点
 */
int WordHandler::GetTableColumnNumber(int tbl_number,xNode cur){
    cur = cur -> xmlChildrenNode;
    while(cur !=NULL){
        if(!xmlStrcmp(cur->name,(const xChar)"tbl")){
            tbl_number --;
            if(!tbl_number){
                table_column_number = 0;
                xNode cur_tr =NULL;
                cur_tr = cur->children;
                while(cur_tr !=NULL){
                    if(!xmlStrcmp(cur_tr->name,(const xChar)"tr")){
                        xNode cur_tc = NULL;
                        cur_tc = cur_tr->children;
                        while(cur_tc !=NULL){
                            if(!xmlStrcmp(cur_tc->name,(const xChar)"tc")){
                                table_column_number++;
                            }
                            cur_tc = cur_tc ->next;
                        }
                        break;
                    }
                    cur_tr = cur_tr ->next;
                }
            }
        }
        GetTableColumnNumber(tbl_number,cur);
        cur = cur->next;
    }
    return table_column_number;
}


/*
 *  函数名：GetAllParaphNumber(xNode cur)
 *  函数说明：得到段落总数
 *  函数参数:1 cur 节点
 */
int WordHandler::GetAllParaphNumber(xNode cur){
    cur = cur -> xmlChildrenNode;
    while(cur !=NULL){
        if(!xmlStrcmp(cur->name,(const xChar)"p")){
            paraph_number++;
        }
        GetAllParaphNumber(cur);
        cur = cur->next;
    }
    return paraph_number;
}



/*
 *  函数名：GetTableRowNumber(int tbl_number,xNode cur)
 *  函数说明：得到指定表格的行数
 *  函数参数:1 tbl_number 表格索引
 *          2 cur        节点
 */
int WordHandler::GetTableRowNumber(int tbl_number,xNode cur){
    cur = cur -> xmlChildrenNode;
    while(cur !=NULL){
        if(!xmlStrcmp(cur->name,(const xChar)"tbl")){
            tbl_number --;
            if(!tbl_number){
                table_row_number = 0;
                xNode cur_r =NULL;
                cur_r = cur->children;
                while(cur_r !=NULL){
                    if(!xmlStrcmp(cur_r->name,(const xChar)"tr")){
                        table_row_number++;

                    }
                    cur_r = cur_r ->next;
                }
            }
        }
        GetTableRowNumber(tbl_number,cur);
        cur = cur->next;
    }
    return table_row_number;
}


/*
 *  函数名：FontScale(int tbl_number, int row_number, int column_number, int value, xNode cur)
 *  函数说明：表格内字体缩放程度
 *  函数参数:1 tbl_number       表格索引
 *          2 row_number       行数
 *          3 column_number    列数
 *          4 value            缩放比例 1-100
 *          5 cur              节点
 */
void WordHandler::FontScale(int tbl_number, int row_number, int column_number, int value, xNode cur){
    cur = cur->xmlChildrenNode;                                                //递归
    while(cur != NULL){
        if(!xmlStrcmp(cur -> name,(const xChar)"tbl")){                      //找到节点名为tbl的节点
            tbl_number--;
            if(!tbl_number){
                xNode cur1 = NULL;
                cur1 = cur ->xmlChildrenNode;
                while(cur1 != NULL){
                    if(!xmlStrcmp(cur1 -> name,(const xChar)"tr")){        //tr节点指的是行， 找到行节点
                        row_number --;
                        if(!row_number){
                            xNode cur2;
                            cur2 = cur1 -> xmlChildrenNode;
                            while(cur2 != NULL){
                                if(!xmlStrcmp(cur2 -> name,(const xChar)"tc")){   //tc节点指列 找到列节点 列节点为行节点子节点
                                    column_number --;
                                    if(!column_number){
                                        ChangeFontscale(value,cur2);
                                    }
                                }
                                cur2 = cur2 ->next;
                            }
                        }
                    }
                    cur1 = cur1 -> next;
                }
            }
        }

        FontScale(tbl_number, row_number, column_number, value, cur);
        cur = cur->next; /* 下一个子节点 */
    }
}


//改变字体比例
void WordHandler::ChangeFontscale(int value,xNode cur){
    cur = cur->xmlChildrenNode;
    while(cur != NULL){
        int value_temp = value;
        if(!xmlStrcmp(cur -> name,(const xChar)"sz")){
            stringstream ss;
            string Font_size;
            value_temp *= 2;
            ss << value_temp;
            Font_size = ss.str();
            xmlSetNsProp(cur,word_namespace,(const xChar)"val",(const xChar)Font_size.data());
        }
        else if(!xmlStrcmp(cur -> name,(const xChar)"szCs")){
            stringstream ss;
            string Font_size;
            value_temp *= 2;
            ss << value_temp;
            Font_size = ss.str();
            xmlSetNsProp(cur,word_namespace,(const xChar)"val",(const xChar)Font_size.data());
        }
        ChangeFontscale(value, cur);
        cur = cur ->next;
    }
}




//将xml文件提取出来
void  WordHandler::ExtractXML(const string &path){

    opcInitLibrary();
    opcContainer *c=opcContainerOpen((const xChar)path.data(), OPC_OPEN_READ_WRITE, NULL, NULL);
    if (NULL!=c) {
        opcPart part=opcPartFind(c,(const xChar)"word/document.xml", NULL, 0);
        if (OPC_PART_INVALID!=part) {
            const xmlChar* type=opcPartGetType(c, part);
            opc_uint32_t type_len=xmlStrlen(type);
            opc_bool_t is_xml=NULL!=type && type_len>=3 && 'x'==type[type_len-3] && 'm'==type[type_len-2] && 'l'==type[type_len-1];
            fprintf(stderr, "type=%s is_xml=%i\n", type, is_xml);
            opcContainerInputStream* stream = opcContainerOpenInputStream(c, part);
            if (stream != NULL) {
                FILE *source = fopen("document.xml","w+");
                if (source != NULL){
                    //opc_uint32_t result = 0;
                    opc_uint32_t ret=0;
                    opc_uint8_t buf[100] = {0};
                    while((ret=opcContainerReadInputStream(stream, buf, sizeof(buf)))>0) {
                        fwrite(buf, sizeof(opc_uint8_t), ret, source);
                    }
                }
                fclose(source);
            }
            opcContainerCloseInputStream(stream);
        }
    }
}



//将xml文件放回opc容器里
void WordHandler::InputXML(const string &path){
    opcInitLibrary();
    opcContainer *c=opcContainerOpen((const xChar)path.data(), OPC_OPEN_READ_WRITE, NULL, NULL);
    if (NULL!=c) {
        opcPart part=opcPartFind(c,(const xChar)"word/document.xml", NULL, 0);
        if (OPC_PART_INVALID!=part) {
            const xmlChar* type=opcPartGetType(c, part);
            opc_uint32_t type_len=xmlStrlen(type);
            opc_bool_t is_xml=NULL!=type && type_len>=3 && 'x'==type[type_len-3] && 'm'==type[type_len-2] && 'l'==type[type_len-1];
            fprintf(stderr, "type=%s is_xml=%i\n", type, is_xml);
            opcContainerOutputStream* stream = opcContainerCreateOutputStream(c, part,OPC_COMPRESSIONOPTION_NORMAL);
            if (stream != NULL) {
                FILE *source = fopen("document.xml","r+");
                if (source != NULL){
                    opc_uint32_t result = 0;
                    //opc_uint32_t ret=0;
                    opc_uint8_t buf[100] = {0};
                    while((result = fread(buf, sizeof(opc_uint8_t), sizeof(buf), source))>0) {
                        opcContainerWriteOutputStream(stream, buf, result);
                    }
                }
                fclose(source);
            }
            opcContainerCloseOutputStream(stream);
            opcContainerClose(c, OPC_CLOSE_NOW);
        }
    }
}

/*void Widget::on_pushButton_clicked()
{
    xmlDocPtr doc = NULL;
    xmlNodePtr cur = NULL;
    WordHandler word;
    string oldname = "/home/huangwei/build-WordHander-Desktop_Qt_5_10_1_GCC_64bit-Debug/1.docx";
    string newname = "/home/huangwei/build-WordHander-Desktop_Qt_5_10_1_GCC_64bit-Debug/多子类模板.docx";
     word.CopyFile(oldname,newname);
    word.ExtractXML(newname);
    word.Initxml(cur,doc);
    word.DeleteHideBookmark(cur);
    clock_t start,end;
    start=clock();
    const int candidate_number[20] = {10,8,5,5,7,8,9,2,1,4,24,46,57,8,9,5,6,7,8,9};                                //每个子类候选人人数
    const int index_sum = 19;                                               //子类索引
    const int xml_index_sum = index_sum +1;                                 //子类的个数
    const int other_number[20] = {5,0,4,6,8,9,0,0,9,0,9,0,8,7,6,7,9,0,0,9};                                        //另选人人数
    int sub_index = 1;

    std::ostringstream bookname_oss;
    std::ostringstream bookcontent_oss;
    bookname_oss <<"mYB" << 1 << "S" << 1  << "C";
    word.ChangeBookmarkname("m应到人数", bookname_oss.str(),cur);
    word.ChangeBookmarkContent(bookname_oss.str(),"yb1",0,cur);
    bookname_oss.str("");
    bookname_oss <<"mRB" << 1 << "S" << 1 << "C";
    word.ChangeBookmarkname("m实到人数", bookname_oss.str(),cur);
    word.ChangeBookmarkContent(bookname_oss.str(),"rb1",0,cur);
    bookname_oss.str("");
    bookname_oss <<"mFB" << 1 << "S" << 1  << "C";
    word.ChangeBookmarkname("m发出选票", bookname_oss.str(),cur);
    word.ChangeBookmarkContent(bookname_oss.str(),"fb1",0,cur);
    bookname_oss.str("");
    bookname_oss <<"mSB" << 1 << "S" << 1  << "C";
    word.ChangeBookmarkname("m收回选票", bookname_oss.str(),cur);
    word.ChangeBookmarkContent(bookname_oss.str(),"sb1",0,cur);
    bookname_oss.str("");
    bookname_oss <<"mVB" << 1 << "S" << 1  << "C";
    bookcontent_oss <<"vb" << 1 << "s" << 1 ;
    word.ChangeBookmarkname("m有效选票", bookname_oss.str(),cur);
    word.ChangeBookmarkContent(bookname_oss.str(),bookcontent_oss.str(),0,cur);
    bookname_oss.str("");
    bookcontent_oss.str("");
    bookname_oss <<"mWB" << 1 << "S" << 1 << "C";
    bookcontent_oss <<"wb" << 1 << "s" << 1 ;
    word.ChangeBookmarkname("m无效选票", bookname_oss.str(),cur);
    word.ChangeBookmarkContent(bookname_oss.str(),bookcontent_oss.str(),0,cur);
    bookname_oss.str("");
    bookcontent_oss.str("");
    bookname_oss <<"mLB" << 1 << "S" << 1  << "C";
    word.ChangeBookmarkname("m另选区", bookname_oss.str(),cur);
    bookname_oss.str("");
    bookname_oss <<"mDB" << 1 << "S" << 1  << "C";
    word.ChangeBookmarkname("m另选排序区", bookname_oss.str(),cur);
    bookname_oss.str("");
    bookname_oss <<"mEB" << 1 << "S" << 1  << "C";
    word.ChangeBookmarkname("m候选排序待删", bookname_oss.str(),cur);
    bookname_oss.str("");
    bookname_oss <<"mHB" << 1 << "S" << 1  << "C";
    word.ChangeBookmarkname("m子类名", bookname_oss.str(),cur);
    bookname_oss.str("");
    bookname_oss <<"mCB" << 1 << "S";
    word.ChangeBookmarkname("m日期", bookname_oss.str(),cur);
    bookname_oss.str("");
    xmlSaveFormatFile("document.xml", doc, 0);
    for(sub_index;sub_index< xml_index_sum ;sub_index++){
        int other_table = 2*sub_index +1;
        WordHandler CopyOther;
        CopyOther.Initxml(cur,doc);
        vector<string> bookname;
        vector<string> bookcontent;
        std::ostringstream oldbookname_oss;
        std::ostringstream newbookname_oss;
        bookcontent.clear();
        oldbookname_oss <<"mLB" << 1 << "S" << sub_index  << "C";
        newbookname_oss <<"mLB" << 1 << "S" << sub_index+1 << "C";
        bookname.push_back(newbookname_oss.str());
        newbookname_oss.str("");
        newbookname_oss <<"mDB" << 1 << "S" << sub_index+1  << "C";
        bookname.push_back(newbookname_oss.str());
        CopyOther.CopyParaphByBookmark1(oldbookname_oss.str(),bookname,other_table,cur);
        xmlSaveFormatFile("document.xml", doc, 0);
        oldbookname_oss.str("");
        newbookname_oss.str("");
        bookname_oss.str("");
        bookname.clear();
        word.AddEnterForTable(other_table,1,"仿宋_GB2312",10,cur);
        word.CopyTable(2,other_table,1,cur);
        xmlSaveFormatFile("document.xml", doc, 0);
        bookname_oss <<"mEB" << 1 << "S" << sub_index << "C";
        newbookname_oss <<"mEB" << 1 << "S" << sub_index+1 << "C";
        word.CopyParaphByBookmark(bookname_oss.str(),newbookname_oss.str(),"",other_table,1,cur);
        xmlSaveFormatFile("document.xml", doc, 0);
        bookname_oss.str("");
        newbookname_oss.str("");
        bookname_oss <<"mVB" << 1 << "S" << sub_index << "C";
        newbookname_oss <<"mVB" << 1 << "S" << sub_index+1  << "C";
        bookname.push_back(newbookname_oss.str());
        newbookname_oss.str("");
        newbookname_oss <<"mWB" << 1 << "S" << sub_index+1 << "C";
        bookname.push_back(newbookname_oss.str());
        word.CopyParaphByBookmark(bookname_oss.str(),bookname,bookcontent,other_table,1,cur);
        newbookname_oss.str("");
        bookname_oss.str("");
        bookname.clear();
        bookname_oss <<"mHB" << 1 << "S" << sub_index << "C";
        newbookname_oss <<"mHB" << 1 << "S" << sub_index+1 << "C";
        word.CopyParaphByBookmark(bookname_oss.str(),newbookname_oss.str(),"",other_table,1,cur);
        xmlSaveFormatFile("document.xml", doc, 0);
        bookname_oss.str("");
        newbookname_oss.str("");
        word.AddEnterForTable(other_table,1,"仿宋_GB2312",10,cur);
        xmlSaveFormatFile("document.xml", doc, 0);
    }
    for(sub_index = 1; sub_index <= xml_index_sum;sub_index ++){
        for(int add_row = 1;add_row < candidate_number[sub_index -1];add_row ++){
            word.AddRow(2*sub_index,2,cur);
            xmlSaveFormatFile("document.xml", doc, 0);
        }
    }
    for(sub_index = 1; sub_index <= xml_index_sum;sub_index ++){
        int column_sum= word.GetTableColumnNumber(2*sub_index,cur);
        int row_sum = word.GetTableRowNumber(2*sub_index,cur);
        for(int row = 2;row <= row_sum;row ++){
            for(int column = 1;column <= column_sum; column ++){
                std::ostringstream name_oss;
                std::ostringstream content_oss;
                int candIndex = row -1;                               //候选人索引
                switch (column) {
                case 2:
                    name_oss << "mNB" <<1 << "S" << sub_index  << "C" << candIndex << "E";
                    content_oss <<"候" << sub_index  << "_" << candIndex;
                    word.AddBookmarkForTable(2*sub_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
                case 3:
                    name_oss << "mAB" << 1 << "S" << sub_index  << "C" << candIndex << "E";
                    content_oss <<"赞" << sub_index  << "_" << candIndex;
                    word.AddBookmarkForTable(2*sub_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
                case 4:
                    name_oss << "mUB" << 1 << "S" << sub_index << "C" << candIndex << "E";
                    content_oss <<"反" << sub_index  << "_" << candIndex;
                    word.AddBookmarkForTable(2*sub_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
                case 5:
                    name_oss << "mQB" << 1 << "S" << sub_index   << "C" << candIndex << "E";
                    content_oss <<"弃" << sub_index   << "_" << candIndex;
                    word.AddBookmarkForTable(2*sub_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
                case 6:
                    name_oss << "mOB" << 1 << "S" << sub_index  << "C" << candIndex << "E";
                    content_oss <<"废" << sub_index  << "_" << candIndex;
                    word.AddBookmarkForTable(2*sub_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
                case 7:
                    name_oss << "mZB" << 1 << "S" << sub_index  << "C" << candIndex << "E";
                    content_oss <<"率" << sub_index << "_" << candIndex;
                    word.AddBookmarkForTable(2*sub_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
                default:
                    break;
                }
            }
        }
        word.DeleteColumn(2*sub_index,1,cur);
        word.DeleteColumn(2*sub_index,5,cur);
        word.DeleteColumn(2*sub_index,6,cur);
        word.DeleteColumn(2*sub_index,7,cur);
        xmlSaveFormatFile("document.xml", doc, 0);
    }
    for(sub_index = 1; sub_index <= xml_index_sum;sub_index ++){
        std::ostringstream other;
        WordHandler Other;
        other <<"mLB" << 1 << "S" << sub_index  << "C";
        if(!other_number[sub_index-1]){
            word.DeleteBookmarkInParaph1(other.str(),cur);
            xmlSaveFormatFile("document.xml", doc, 0);
        }else{
            for(int other_add_row = 1;other_add_row < other_number[sub_index-1];other_add_row++ ){
                Other.AddRowInBookmarkTable(other.str(),2,cur);
                xmlSaveFormatFile("document.xml", doc, 0);
            }
        }
    }
    auto t = chrono::system_clock::to_time_t(std::chrono::system_clock::now());
    std::ostringstream data_oss;
    std::ostringstream data_name;
    data_oss<<std::put_time(std::localtime(&t), "%Y年%m月%d日");
    data_name<<"mCB" << 1 << "S";
    word.ChangeBookmarkContent(data_name.str(),data_oss.str(),1,cur);
    xmlSaveFormatFile("document.xml", doc, 0);
    word.InputXML(newname);
    end = clock();
    double endtime=(double)(end-start)/CLOCKS_PER_SEC;
    cout << endtime <<"s" <<endl;
}*/

/*void Widget::on_pushButton_2_clicked()
{
    xmlDocPtr doc = NULL;
    xmlNodePtr cur = NULL;
    WordHandler word;
    string oldname = "/home/huangwei/build-WordHander-Desktop_Qt_5_10_1_GCC_64bit-Debug/2.docx";
    string newname = "/home/huangwei/build-WordHander-Desktop_Qt_5_10_1_GCC_64bit-Debug/单子类模板.docx";
    word.CopyFile(oldname,newname);
    word.ExtractXML(newname);
    word.Initxml(cur,doc);
    word.DeleteHideBookmark(cur);
        const int candidate_number = 10;                            //候选人人数
        const int row_number = candidate_number - 1;                      //计算出的需要的表格数
        const int index = 0;                                               //子类索引
        const int xml_index = index + 1;                                   //xml文件的索引从1开始
        const int other_number = 5;                                        //另选人人数
        int tbl_number_index = (xml_index)*2+1;
        int row = 0;
        int column = 0;
        int other_row = 0;
        for(int add_row = 0;add_row < row_number;add_row++){
            word.AddRow(tbl_number_index,2,cur);
        }
        xmlSaveFormatFile("document.xml", doc, 0);
        int row_sum = word.GetTableRowNumber(tbl_number_index,cur);
        int column_sum = word.GetTableColumnNumber(tbl_number_index,cur);
        for(row = 2;row <=row_sum ; row ++){
            for(column = 1;column <= column_sum;column ++){
                std::ostringstream name_oss;
                std::ostringstream content_oss;
                int candIndex = row -1;                               //候选人索引
                switch (column) {
                case 2:
                    name_oss << "mNB" <<1 << "S" << xml_index  << "C" << candIndex << "E";
                    content_oss <<"候" << xml_index  << "_" << candIndex;
                    word.AddBookmarkForTable(tbl_number_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
               case 3:
                    name_oss << "mAB" << 1 << "S" << xml_index  << "C" << candIndex << "E";
                    content_oss <<"赞" << xml_index  << "_" << candIndex;
                    word.AddBookmarkForTable(tbl_number_index,row,column,name_oss.str(),content_oss.str(),cur);
                    break;
                case 4:
                     name_oss << "mUB" << 1 << "S" << xml_index  << "C" << candIndex << "E";
                     content_oss <<"反" << xml_index  << "_" << candIndex;
                     word.AddBookmarkForTable(tbl_number_index,row,column,name_oss.str(),content_oss.str(),cur);
                     break;
                case 5:
                     name_oss << "mQB" << 1 << "S" << xml_index  << "C" << candIndex << "E";
                     content_oss <<"弃" << xml_index  << "_" << candIndex;
                     word.AddBookmarkForTable(tbl_number_index,row,column,name_oss.str(),content_oss.str(),cur);
                     break;
                case 6:
                     name_oss << "mOB" << 1 << "S" << xml_index  << "C" << candIndex << "E";
                     content_oss <<"废" << xml_index  << "_" << candIndex;
                     word.AddBookmarkForTable(tbl_number_index,row,column,name_oss.str(),content_oss.str(),cur);
                     break;
                case 7:
                     name_oss << "mZB" << 1 << "S" << xml_index  << "C" << candIndex << "E";
                     content_oss <<"率" << xml_index  << "_" << candIndex;
                     word.AddBookmarkForTable(tbl_number_index,row,column,name_oss.str(),content_oss.str(),cur);
                     break;
                default:
                    break;
                }
            }
        }
        xmlSaveFormatFile("document.xml", doc, 0);
        word.DeleteColumn(tbl_number_index,1,cur);            //只剩下姓名赞成反对弃权
        word.DeleteColumn(tbl_number_index,5,cur);
        word.DeleteColumn(tbl_number_index,6,cur);
        word.DeleteColumn(tbl_number_index,7,cur);
        std::ostringstream bookname_oss;
        bookname_oss <<"mYB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m应到人数", bookname_oss.str(),cur);
        word.ChangeBookmarkContent(bookname_oss.str(),"yb1",0,cur);
        bookname_oss.str("");
        bookname_oss <<"mRB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m实到人数", bookname_oss.str(),cur);
        word.ChangeBookmarkContent(bookname_oss.str(),"rb1",0,cur);
        bookname_oss.str("");
        bookname_oss <<"mFB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m发出选票", bookname_oss.str(),cur);
        word.ChangeBookmarkContent(bookname_oss.str(),"fb1",0,cur);
        bookname_oss.str("");
        bookname_oss <<"mSB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m收回选票", bookname_oss.str(),cur);
        word.ChangeBookmarkContent(bookname_oss.str(),"sb1",0,cur);
        bookname_oss.str("");
        bookname_oss <<"mVB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m有效选票", bookname_oss.str(),cur);
        word.ChangeBookmarkContent(bookname_oss.str(),"vb1",0,cur);
        bookname_oss.str("");
        bookname_oss <<"mWB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m无效选票", bookname_oss.str(),cur);
        word.ChangeBookmarkContent(bookname_oss.str(),"wb1",0,cur);
        bookname_oss.str("");
        bookname_oss <<"mLB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m另选区", bookname_oss.str(),cur);
        bookname_oss.str("");
        bookname_oss <<"mDB" << 1 << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkname("m另选排序区", bookname_oss.str(),cur);
        bookname_oss.str("");
        bookname_oss <<"mCB" << 1 << "S";
        word.ChangeBookmarkname("m日期", bookname_oss.str(),cur);
        bookname_oss.str("");
        xmlSaveFormatFile("document.xml", doc, 0);
        for(other_row = 1;other_row < other_number;other_row ++){
        std::ostringstream row_oss;
        row_oss <<"mLB" << 1 << "S" << xml_index  << "C";
        word.AddRowInBookmarkTable(row_oss.str(),2,cur);
        }
        xmlSaveFormatFile("document.xml", doc, 0);
        auto t = chrono::system_clock::to_time_t(std::chrono::system_clock::now());
        std::ostringstream data_oss;
        std::ostringstream data_name;
        data_oss<<std::put_time(std::localtime(&t), "%Y年%m月%d日");
        data_name<<"mCB" << 1 << "S";
        word.ChangeBookmarkContent(data_name.str(),data_oss.str(),1,cur);
        word.ChangeBookmarkContent("m监票人","",1,cur);
        word.ChangeBookmarkContent("m计票","",1,cur);
        word.ChangeBookmarkContent("m会议名","",1,cur);
        word.ChangeBookmarkContent("m选票名","",1,cur);
        word.ChangeBookmarkContent("一","",1,cur);
        word.ChangeBookmarkContent("二","",1,cur);
        word.ChangeBookmarkContent("三","",1,cur);
        word.DeleteBookmarkInParaph("m候选排序待删",cur);
        std::ostringstream sort_oss;
        sort_oss <<"mDB" << 1 << "S" << xml_index  << "C";
        word.ChangeBookmarkContent(sort_oss.str(),"",1,cur);
        xmlSaveFormatFile("document.xml", doc, 0);
        std::ostringstream other_oss;
        other_oss<<"mLB" << 1 << "S" << xml_index  << "C";
        for(int other_index = 2;other_index <= other_number+1;other_index ++){
            word.AddContentInBookmarkTable(other_oss.str(),other_index,1,"huangwei",cur);
        }
        xmlSaveFormatFile("document.xml", doc, 0);
       word.InputXML(newname);
}*/
