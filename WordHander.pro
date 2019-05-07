#-------------------------------------------------
#
# Project created by QtCreator 2019-03-18T11:16:54
#
#-------------------------------------------------

#模块
QT       += core gui
CONFIG  +=c++11

#为了兼容QT4以前的版本
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets


#应用程序的名字
TARGET = WordHander
TEMPLATE = app

# The following define makes your compiler emit warnings if you use
# any feature of Qt which has been marked as deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

# You can also make your code fail to compile if you use deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0


SOURCES += \
        main.cpp \
    form.cpp \
    resultoutputwidget.cpp \
    wordHander.cpp \
    resutoutput.cpp

HEADERS += \
    wordhander.h \
    form.h \
    resultoutputwidget.h \
    wordHander.h \
    resutoutput.h

FORMS += \
    form.ui \
    resultoutputwidget.ui

win32:CONFIG(release, debug|release): LIBS += -L$$PWD/../libconfigure/zlib/lib/release/ -lz
else:win32:CONFIG(debug, debug|release): LIBS += -L$$PWD/../libconfigure/zlib/lib/debug/ -lz
else:unix: LIBS += -L$$PWD/../libconfigure/zlib/lib/ -lz

INCLUDEPATH += $$PWD/../libconfigure/zlib/include
DEPENDPATH += $$PWD/../libconfigure/zlib/include

win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/zlib/lib/release/libz.a
else:win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/zlib/lib/debug/libz.a
else:win32:!win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/zlib/lib/release/z.lib
else:win32:!win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/zlib/lib/debug/z.lib
else:unix: PRE_TARGETDEPS += $$PWD/../libconfigure/zlib/lib/libz.a

win32:CONFIG(release, debug|release): LIBS += -L$$PWD/../libconfigure/libxml2/lib/release/ -lxml2
else:win32:CONFIG(debug, debug|release): LIBS += -L$$PWD/../libconfigure/libxml2/lib/debug/ -lxml2
else:unix: LIBS += -L$$PWD/../libconfigure/libxml2/lib/ -lxml2

INCLUDEPATH += $$PWD/../libconfigure/libxml2/include
DEPENDPATH += $$PWD/../libconfigure/libxml2/include

win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libxml2/lib/release/libxml2.a
else:win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libxml2/lib/debug/libxml2.a
else:win32:!win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libxml2/lib/release/xml2.lib
else:win32:!win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libxml2/lib/debug/xml2.lib
else:unix: PRE_TARGETDEPS += $$PWD/../libconfigure/libxml2/lib/libxml2.a

win32:CONFIG(release, debug|release): LIBS += -L$$PWD/../libconfigure/libopc/release/ -lopc
else:win32:CONFIG(debug, debug|release): LIBS += -L$$PWD/../libconfigure/libopc/debug/ -lopc
else:unix: LIBS += -L$$PWD/../libconfigure/libopc/ -lopc

INCLUDEPATH += $$PWD/../libconfigure/libopc
DEPENDPATH += $$PWD/../libconfigure/libopc

win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/release/libopc.a
else:win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/debug/libopc.a
else:win32:!win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/release/opc.lib
else:win32:!win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/debug/opc.lib
else:unix: PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/libopc.a

win32:CONFIG(release, debug|release): LIBS += -L$$PWD/../libconfigure/libopc/release/ -lmce
else:win32:CONFIG(debug, debug|release): LIBS += -L$$PWD/../libconfigure/libopc/debug/ -lmce
else:unix: LIBS += -L$$PWD/../libconfigure/libopc/ -lmce

INCLUDEPATH += $$PWD/../libconfigure/libopc
DEPENDPATH += $$PWD/../libconfigure/libopc

win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/release/libmce.a
else:win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/debug/libmce.a
else:win32:!win32-g++:CONFIG(release, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/release/mce.lib
else:win32:!win32-g++:CONFIG(debug, debug|release): PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/debug/mce.lib
else:unix: PRE_TARGETDEPS += $$PWD/../libconfigure/libopc/libmce.a

RESOURCES += \
    imageresource.qrc
