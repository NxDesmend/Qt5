#ifndef EXPLUGININTERFACE__H
#define EXPLUGININTERFACE__H

#include <QObject>
#include <QWidget>
#include <QString>

class EXPluginInterface
{
public:
    virtual ~EXPluginInterface() {}
    virtual QString title() = 0;
    virtual QString version() = 0;
    virtual QString description() = 0;
    virtual QString help() = 0;
    virtual void init(QWidget* parent) = 0;
    virtual void processExcel() = 0;
};

#define EXPLUGININTERFACE_IID "com.excel.plugininterface"
Q_DECLARE_INTERFACE(EXPluginInterface, EXPLUGININTERFACE_IID)

#endif
