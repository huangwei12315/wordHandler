#ifndef WORDHANDER_H
#define WORDHANDER_H
#include<iostream>
#include<stdio.h>
#include <string.h>
#include <stdlib.h>
#include <stdexcept>
#include <vector>
#include <string>
#include <libxml/tree.h>
#include <libxml/parser.h>
#include<opc/opc.h>
#include <sstream>
#include <iosfwd>
#include <chrono>
#include <ctime>
#include <time.h>
#include<iomanip>
#include <QFileInfo>
#include <QString>
#include <QDir>
#include <QDebug>
#include <iostream>
using namespace std;
#ifdef __cplusplus
extern "C" {
#endif


class WordHandler
{
public:
    typedef xmlChar* xChar;
    typedef xmlNodePtr xNode;
    typedef xmlDocPtr xDoc;
    typedef xmlNsPtr xNs;

    WordHandler();

    void CopyFile(QString Qoldpath,QString Qnewpath);
    /*
     *  函数名：initxml(xNode &cur_temp, xDoc &doc_temp , xNs* &name_temp)
     *  函数说明： 初始化xml节点
     *  函数参数: 1 cur_temp 节点
     *           2 doc_temp 文档节点
     *           3 name_temp 命名空间
     */
    void Initxml(xNode &cur_temp, xDoc &doc_temp);
    /*
     *  函数名：CopyTable(int tbl_number,int tbl_sum,xNode cur)
     *  函数说明：复制表格及表格上面的段落，粘贴到最后一个表格的后面
     *  函数参数: 1 tbl_number 表格索引
     *           2 tbb_sum 表格总数
     */
    void CopyTableAndParaph(int tbl_copy,int tbl_sum,xNode cur);




    /*
     *  函数名：CopyTable(int tbl_number_copy,int tbl_number_paste,bool position,xNode cur)
     *  函数说明：复制某个表格到另外一个表格的前或者后
     *  函数参数: 1 tbl_number_copy 复制的表格索引
     *           2 tbl_number_paste 粘贴表格索引
     *           3 position 位置  0为前 1为后
     *           4 cur 节点
     */
    void CopyTable(int tbl_number_copy,int tbl_number_paste,bool position,xNode cur);



    /*
     *  函数名：CopyParaphByBookmark(const string &bookmark_name,int tbl_number_paste,bool position,xNode cur)
     *  函数说明：通过书签粘贴指定段落到表格的前面或者后面
     *  函数参数: 1 bookmark_name  书签名
     *           2 tbl_number_paste 粘贴的段落索引号
     *           3 position  位置，0为表格前，1为表格后
     *           4 cur 节点
    */
    void CopyParaphByBookmark(const string &old_bookmark_name,const string &new_bookmark_name,const string &book_content,int tbl_number_paste,bool position,xNode cur);



    void CopyParaphByBookmark(const string &old_bookmark_name,const vector<string> &new_bookmark_name,const vector<string> &book_content,int tbl_number_paste,bool position,xNode cur);

    void CopyParaphByBookmark1(const string &old_bookmark_name,const vector<string> &new_bookmark_name,int tbl_number_paste,xmlNodePtr cur);

    /*  函数名：AddCloumnInBookmarkTable(const string &bookmark_name,int copy_column,xNode cur)
     *  函数说明：向在书签内的表格复制任意一行到最后一行
     *  函数参数: 1 bookmark_name   表格所在书签
     *           2 copy_column      复制的行数
     *           3 cur            节点
     */
    void AddCloumnInBookmarkTable(const string &bookmark_name,int copy_column,xNode cur);


    void DeleteCloumnInBookmarkTable(const string &bookmark_name,int delete_column,xNode cur);

    void DeleteRowInBookmarkTable(const string &bookmark_name,int delete_row,xNode cur);

    /*
     * 函数名：AddBookmarkForColumn(int tbl_number,int row_number_Start,int column_number,vector<string> &bookname, vector<string> &book_content,xNode cur)
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
    void AddBookmarkColumn(int tbl_number,int row_number_Start,int column_number,vector<string> &bookname, vector<string> &book_content,xNode cur);



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
    void AddBookmarkForRow(int tbl_number,int row_number,int column_number_Start,vector<string> &bookname, vector<string> &book_content,xNode cur);



    void AddEnterForTable(int tbl_number, bool position,string font,size_t size,xNode cur);
    /*
     *  函数名：AddBookmarkForTable(int tbl_number,int row_number_Start,int column_number_Start,,vector<string> &bookname, vector<string> &book_content,xNode cur)
     *  函数说明：向表格单元格内加入书签
     *  函数参数: 1.tbl_number    表示表格的索引号 以1为起始
     * 2. row_number           表示起始行数索引号 以1为起始
     * 3. column_number_Start  表示起始列数索引号 以1为起始
     * 4. bookname             字符串数组，表示添加书签的名称
     * 5. book_content         字符串数组,表示添加书签的内容
    */

    void AddBookmarkForTable(int tbl_number,int row_number_Start,int column_number_Start,const string &bookname, const string &book_content,xNode cur);

    /*
     *  函数名：AddRow(int tbl_number, int copy_line,int paste_line,xNode cur)
     *  函数说明：将指定行复制到第几行，无法复制到最后一行
     *  函数参数: 1 tbl_number表格索引
     *           2 copy_line 复制行数
     *           3 paste_line 需要粘贴行数
    */
    void AddRow(int tbl_number,int copy_line,int paste_line,xNode cur);


    /*
     *  函数名：AddRow(int tbl_number, int copy_line,xNode cur)
     *  函数说明：将某行复制到最后一行
     *  函数参数: 1 tbl_number表格索引
     *           2 xNode cur 节点
    */
    void AddRow(int tbl_number,int copy_line,xNode cur);




    /*
     *  函数名：AddRow(int tbl_number, int copy_cloumn,int  paste_cloumn,xNode cur)
     *  函数说明：将指定列复制到第几列，无法复制到最后一列
     *  函数参数: 1 tbl_number表格索引
     *           2 copy_cloumn, 复制列数
     *           3  paste_cloumn 需要粘贴列数
    */
    void AddCloumn(int tbl_number,int copy_cloumn,int paste_cloumn,xNode cur);

    /*
     *  函数名：AddRow(int tbl_number, int copy_cloumn,xNode cur)
     *  函数说明：将某列复制到最后一列
     *  函数参数: 1 tbl_number表格索引
     *           2 复制的列数
     *           3 xNode cur 节点
    */
    void AddCloumn(int tbl_number,int copy_cloumn,xNode cur);


    /*  函数名：AddContentInBookmarkTable(const string &bookmark_name,int row_number,int column_number,const string &book_content,xNode cur)
    *  函数说明：向在书签内的表格加入内容
    *  函数参数: 1 bookmark_name   表格所在书签
    *           2 row_number      行号
    *           3 column_number   列号
    *           4 book_content    书签内容
    */
    void AddContentInBookmarkTable(const string &bookmark_name,int row_number,int column_number,const string &book_content,xNode cur);






    /*
     *  函数名：ChangeBookmarkCotent(const string &bookmark_name,const string &change_content,bool delete_bookmark,xNode cur)
     *  函数说明：改变书签内容 可以选择是否删掉书签
     *  函数参数: 1 bookmark_name  书签名字
     *           2 change_content  现书签内容
     *           3 delete_bookmark 是否需要删除书签 1为删除书签 0则不删除
     *           4 cur 节点
    */
    void ChangeBookmarkContent(const string &bookmark_name,const string &change_content,bool delete_bookmark,xNode cur);




    /*
     *  函数名：ChangeBookmarkCotent(const string &bookmark_name,const string &change_content,string font,size_t size,bool delete_bookmark,xNode cur)
     *  函数说明：改变书签内容 可以选择是否删掉书签
     *  函数参数: 1 bookmark_name  书签名字
     *           2 change_content  现书签内容
     *           3 font            字体
     *           4 size            字体大小
     *           5 delete_bookmark 是否需要删除书签 1为删除书签 0则不删除
     *           6 cur 节点
    */
    void ChangeBookmarkContent(const string &bookmark_name,const string &change_content,string font,size_t size,bool delete_bookmark,xNode cur);





    /*
     *  函数名：ChangeBookmarkId(const string &bookmark_id,const string &change_id,xNode cur,xmlNsPtr *NameSpace)
     *  函数说明：修改书签id
     *  函数参数: 1 bookmark_id  书签id
     *           2 change_id  需改变的书签id
     *           3 cur 节点
    */
    void ChangeBookmarkId(const string &bookmark_id,const string &change_id,xNode cur);


    /*
     *  函数名：ChangeBookmarkname(const string &bookmark_name,const string &change_name,xNode cur,xmlNsPtr *NameSpace)
     *  函数说明：修改书签名（注意书签名必须英文字母开头）
     *  函数参数: 1 bookmark_name  书签名
     *           2 change_name  需修改的书签名
     *           3 cur 节点
    */
    void ChangeBookmarkname(const string &bookmark_name,const string &change_name,xNode cur);



    /*
     *  函数名：ChangeTablecontent(int tbl_number, int row_number, int column_number, const string &change_content, xNode cur);
     *  函数说明：修改表格中的内容
     *  函数参数: 1 tbl_number      表格索引
     *           2 row_number      行号索引
     *           3 column_number   列号索引
     *           4 content         改变的内容
     *           5 cur             节点
     */
    void ChangeTablecontent(int tbl_number, int row_number, int column_number, const string &change_content, xNode cur);


    /*
     *  函数名：DeleteHideBookmark()
     *  函数说明：删除隐藏书签
     *  函数参数: 无
     */



    void DeleteHideBookmark(xNode cur);


    /*
     *  函数名：DeleteBookmarkForId(const string ID_number,xNode cur)
     *  函数说明：用id删除书签
     *  函数参数: 1 ID_number ID号
     *           2 cur 节点
     */

    void DeleteBookmarkForId(const string ID_number,xNode cur);



    /*
     *  函数名：DeleteBookmarkInParaph(const string &bookmark_name,xNode cur)
     *  函数说明：用书签删除段落
     *  函数参数: 1 bookmark_name 书签名
     *           2 cur 节点
     */
    void DeleteBookmarkInParaph(const string &bookmark_name,xNode cur);


    /*
     *  函数名：DeleteBookmarkInParaph(const string &bookmark_name,xNode cur)
     *  函数说明：用书签删除段落 跨行一并删除
     *  函数参数: 1 bookmark_name 书签名
     *           2 cur 节点
     */
    void DeleteBookmarkInParaph1(const string &bookmark_name,xNode cur);


    /*
     *  函数名：DeleteColumn(int tbl_number,int column_number,xNode cur)
     *  函数说明：删除指定列
     *  函数参数: 1 tbl_number 表格索引
     *           2 column_number 列索引
     *           3 cur 节点
     */
    void DeleteColumn(int tbl_number,int column_number,xNode cur);



    /*
     *  函数名：DeleteRow(int tbl_number,int row_number,xNode cur)
     *  函数说明：删除指定列
     *  函数参数: 1 tbl_number 表格索引
     *           2 column_number 列索引
     *           3 cur 节点
     */
    void DeleteRow(int tbl_number,int row_number,xNode cur);


    /*
     *  函数名：GetAllParaphNumber(xNode cur)
     *  函数说明：得到段落总数
     *  函数参数:1 cur 节点
     */
    int GetAllParaphNumber(xNode cur);



    /*
     *  函数名：GetTableRowNumber(int tbl_number,xNode cur)
     *  函数说明：得到指定表格的行数
     *  函数参数:1 tbl_number 表格索引
     *          2 cur        节点
     */
    int GetTableRowNumber(int tbl_number,xNode cur);



    /*
     *  函数名：GetTableRowNumber(int tbl_number,xNode cur)
     *  函数说明：得到指定表格的列数
     *  函数参数:1 tbl_number 表格索引
     *          2 cur        节点
     */
    int GetTableColumnNumber(int tbl_number,xNode cur);


    void AddRowInBookmarkTable(const string &bookmark_name,int copy_line,xNode cur);

    /*
     *  函数名：FontScale(int tbl_number, int row_number, int column_number, int value, xNode cur)
     *  函数说明：表格内字体缩放程度
     *  函数参数:1 tbl_number       表格索引
     *          2 row_number       行数
     *          3 column_number    列数
     *          4 value            缩放比例 1-100
     *          5 cur              节点
     */
    void FontScale(int tbl_number, int row_number, int column_number, int value, xNode cur);





    //获取表格总数
    int GetTablesum(xNode cur);

    void  ExtractXML(const string &path);

    void InputXML(const string &path);



private:
    void AddColumnBook(int column_number,int tbl_number_temp,unsigned int counter,vector<string> &bookname, vector<string> &book_content, xNode cur2);
    void AddRowBook(int column_number_Start,int tbl_number_temp,unsigned int counter,vector<string> &bookname, vector<string> &book_content, xNode cur2);
    void AddTableBook(int column_number_Start,int tbl_number_temp,const string &bookname, const string &book_content, xNode cur2);
    void AddContent(const string &book_content,xNode cur);
    void ChangeContent(const string &change_content, xNode cur);
    void ChangeContent(const string &change_content,string font, size_t size ,xNode cur);
    void ChangeContent1(const string &change_content, xNode cur);
    void ChangeFontscale(int value,xNode cur);
    void ChangeCopy(const string &old_bookmark_name,const string &new_bookmark_name,const string &book_content,xmlNodePtr cur);
    void ChangeCopy(const vector<string> &new_bookmark_name,const vector<string> &book_content,xmlNodePtr cur);
    void ChangeCopy(const vector<string> &new_bookmark_name,xmlNodePtr cur);
    void DeleteParaph(const string &bookmark_name,xNode cur);
    void DeleteC(int column_number, xNode cur);
    void DeleteR(int row_number, xNode cur);
    bool xmlFindContent(xNode cur);
    xNode word_cur;
    xDoc word_doc;
    xNode cur_save;
    xNode cur_p_start;
    xNode cur_p_end;
    xNode cur_copy_paraph;
    xNode cur_paraph_index;
    xNode cur_delete_r;
    xNs word_namespace;
    xChar attr_id_hide;
    xChar attr_id_value;
    xChar attr_delete_book;
    xChar attr_change_book;
    xChar attr_add_content;
    xChar old_id_attr[10];
    xNode cur_copy[20];
    int paraph_number;
    int change_copy;
    int table_row_number;
    int table_column_number;
    int counter_start;
    int counter_end;


};




#ifdef __cplusplus
}
#endif

#endif // MAINWINDOW_H
