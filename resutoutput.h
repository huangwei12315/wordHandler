#ifndef RESUTOUTPUT_H
#define RESUTOUTPUT_H
#include "wordHander.h"
#include <QApplication>


class resutOutput
{
public:
    resutOutput();
    void creatSimpleTemplate(bool sort_number,bool vote_rate,bool favour,const string &path);
    void creatMutiTemplate(bool sort_number,bool vote_rate,bool favour,const string &path);
private:
    WordHandler word;
};

#endif // RESUTOUTPUT_H
