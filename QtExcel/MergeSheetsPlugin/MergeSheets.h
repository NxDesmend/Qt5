#pragma once

#include <QAxObject>
#include <QDateTime>
#include <QFileDialog>
#include <QMessageBox>

#include "mergeSheets_global.h"
#include "explugininterface.h" 

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
private:
    int openFile();
    int readSheetInfo(QAxObject* worksheet, int serial);
    int createExcelFile();
    int getRowAndColumn(QAxObject* worksheet, int& pRows, int& pColumns);
    int processSheets();
    int writeToDestExcel(QAxObject* worksheet, SheetData& rd);
    int lineWriteToDestExcel(QAxObject* worksheet, SheetData& rd);
    int rectWriteToDestExcel(QAxObject* worksheet, SheetData& rd);
    int initSourceQAxObj();
    int initDestQAxObj();
    int closeDestQAxObj();
    int closeQAxObj();

signals:
    void updateNeeded(QString Info);
    void errorMessage(QString msg);
    void infoMessage(QString msg);

private slots:
    void selectFileEditClicked();
    void selectSaveDirEditClicked();
    void onViewInfoButtonPressed();
    void onMergeSheetsButtonPressed();
    bool isDestPointerValid();

private:
    Ui::PluginGui* ui{nullptr};

    QString excelInfo{};
    QString mergeInfo{};
    QString destFilePath{};
    QString openFilePath{};
    QString saveFilePath{};

    QString m_version{ "1.0.0" };

    QAxObject* sourceExcel{ nullptr };
    QAxObject* sourceWorkbooks{ nullptr };
    QAxObject* sourceWorkbook{ nullptr };

    QAxObject* destExcel{ nullptr };
    QAxObject* destWorkbooks{ nullptr };
    QAxObject* destWorkbook{ nullptr };
    QAxObject* destWorksheets{ nullptr };
    QAxObject* destWorksheet{ nullptr };

    bool isPorcessSheets{ false };
};

