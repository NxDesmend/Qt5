#ifndef  CONVERTEXCELTOSQL_H
#define  CONVERTEXCELTOSQL_H

#include <QFileDialog>
#include <QMessageBox>

#include "explugininterface.h" 
#include "convertexceltosql_global.h"
#include "ui_plugin.h"
#include "xlsxdocument.h"

class CONVERTEXCELTOSQL_EXPORT ConvertExcelToSql : public QObject, public EXPluginInterface
{
	Q_OBJECT
		Q_PLUGIN_METADATA(IID "com.excel.cvplugininterface")
		Q_INTERFACES(EXPluginInterface)
public:
	ConvertExcelToSql();
	~ConvertExcelToSql();

	QString title();
	QString version();
	QString description();
	QString help();
	void init(QWidget* parent);
	void processExcel();

private:
	int createDestFileAndSave();
	int writeToDestTxt();

signals:
	void updateNeeded(QString Info);

private slots:
	void selectFileEditClicked();
	void selectDestDirClicked();
	void onViewInfoButtonPressed();

private:
	Ui::PluginGui* ui{ nullptr };

	QString openFilePath{};
	QString destFilePath{};

	QString convertInfo{};
};

#endif
