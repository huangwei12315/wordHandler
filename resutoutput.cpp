#include "resutoutput.h"

resutOutput::resutOutput()
{

}

void resutOutput::creatMutiTemplate(bool sort_number,bool vote_rate,bool favour,const string &path){
    xmlDocPtr doc = NULL;
    xmlNodePtr cur = NULL;
    word.ExtractXML(path);
    word.Initxml(cur,doc);
    word.DeleteHideBookmark(cur);
    clock_t start,end;
    start=clock();
    const int candidate_number[20] = {10,8,5,5,7,8,9,2,1,4,2,4,7,8,9,5,6,7,8,9};                                //每个子类候选人人数
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
         qApp->processEvents();
    }
    for(sub_index = 1; sub_index <= xml_index_sum;sub_index ++){
        for(int add_row = 1;add_row < candidate_number[sub_index -1];add_row ++){
            word.AddRow(2*sub_index,2,cur);
            xmlSaveFormatFile("document.xml", doc, 0);
        }
        qApp->processEvents();
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
                qApp->processEvents();
            }
            qApp->processEvents();
        }
        if(sort_number){
            word.DeleteColumn(2*sub_index,1,cur);
        }
        if(vote_rate){
            if(sort_number)
                word.DeleteColumn(2*sub_index,6,cur);
            else
                word.DeleteColumn(2*sub_index,7,cur);
        }
        if(favour){
            word.DeleteColumn(2*sub_index,1,cur);
            word.DeleteColumn(2*sub_index,3,cur);
            word.DeleteColumn(2*sub_index,3,cur);
            word.DeleteColumn(2*sub_index,3,cur);
            word.DeleteColumn(2*sub_index,3,cur);
        }
        qApp->processEvents();
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
         qApp->processEvents();
    }
    auto t = chrono::system_clock::to_time_t(std::chrono::system_clock::now());
    std::ostringstream data_oss;
    std::ostringstream data_name;
    data_oss<<std::put_time(std::localtime(&t), "%Y年%m月%d日");
    data_name<<"mCB" << 1 << "S";
    word.ChangeBookmarkContent(data_name.str(),data_oss.str(),1,cur);
    xmlSaveFormatFile("document.xml", doc, 0);
    word.InputXML(path);
    end = clock();
    double endtime=(double)(end-start)/CLOCKS_PER_SEC;
    cout << endtime <<"s" <<endl;
}




void resutOutput::creatSimpleTemplate(bool sort_number,bool vote_rate,bool favour,const string &path){
    xmlDocPtr doc = NULL;
    xmlNodePtr cur = NULL;
    word.ExtractXML(path);
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
        qApp->processEvents();
    }
    xmlSaveFormatFile("document.xml", doc, 0);
    int row_sum = word.GetTableRowNumber(tbl_number_index,cur);
    int column_sum = word.GetTableColumnNumber(tbl_number_index,cur);
    QApplication::setOverrideCursor(Qt::WaitCursor);
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
            qApp->processEvents();
        }
        qApp->processEvents();
    }
    xmlSaveFormatFile("document.xml", doc, 0);
    if(sort_number){
        word.DeleteColumn(tbl_number_index,1,cur);
    }
    if(vote_rate){
        if(sort_number)
            word.DeleteColumn(tbl_number_index,6,cur);
        else
            word.DeleteColumn(tbl_number_index,7,cur);
    }
    if(favour){
        word.DeleteColumn(tbl_number_index,1,cur);
        word.DeleteColumn(tbl_number_index,3,cur);
        word.DeleteColumn(tbl_number_index,3,cur);
        word.DeleteColumn(tbl_number_index,3,cur);
        word.DeleteColumn(tbl_number_index,3,cur);
    }
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
        qApp->processEvents();
    }
    xmlSaveFormatFile("document.xml", doc, 0);
    word.InputXML(path);
    QApplication::restoreOverrideCursor();
}
