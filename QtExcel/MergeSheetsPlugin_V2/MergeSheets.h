#pragma once

#include <QDateTime>
#include <QFileDialog>
#include <QMessageBox>

#include "explugininterface.h" 
#include "mergeSheets_global.h"
#include "xlsxdocument.h"

struct SheetData {
    int rows{0};
    int columns{0};

    int destRows{ 1 };
};

namespace Ui {
    class PluginGui;
}

class MERGESHEETS_EXPORT MergeSheets : public QObject, public EXPluginInterface
{
    Q_OBJECT
        Q_PLUGIN_METADATA(IID "com.excel.cvplugininterface")
        Q_INTERFACES(EXPluginInterface)
public:
    MergeSheets();
    ~MergeSheets();

    QString title();
    QString version();
    QString description();
    QString help();
    void init(QWidget* parent);
    void processExcel();

    static int threadProcess(MergeSheets* ms);
    int processSheets();
private:
    int viewFileInfo();
    int writeToDestSheet(QXlsx::Worksheet* sourceSheet, QXlsx::Worksheet* destSheet, int& destRow);

    int createDestFileAndSave();
    int writeToDest();

signals:
    void updateNeeded(QString Info);
    void errorMessage(QString msg);
    void infoMessage(QString msg);

private slots:
    void selectFileEditClicked();
    void selectSaveDirEditClicked();
    void onViewInfoButtonPressed();
    void onMergeSheetsButtonPressed();

private:
    Ui::PluginGui* ui{nullptr};

    QString excelInfo{};
    QString mergeInfo{};
    QString destFilePath{};
    QString openFilePath{};
    QString saveFilePath{};

    QString m_version{ "2.0.0" };
};

