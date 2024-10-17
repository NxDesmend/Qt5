#include <QDateTime>

#include "ConvertExcelToSql.h"

ConvertExcelToSql::ConvertExcelToSql() {
}

ConvertExcelToSql::~ConvertExcelToSql() {
}

QString ConvertExcelToSql::title() {
	return "ConvertExcel2Sql";
}

QString ConvertExcelToSql::version() {
	return "1.0.0";
}

QString ConvertExcelToSql::description() {
	return "";
}

QString ConvertExcelToSql::help() {
	return "";
}

void ConvertExcelToSql::processExcel() {
	// Todo
}

void ConvertExcelToSql::init(QWidget* parent) {
	ui = new Ui::PluginGui;
	ui->setupUi(parent);

	connect(ui->excelFileLineEdit, SIGNAL(clicked()), this, SLOT(selectFileEditClicked()));
	connect(ui->destDirEdit, SIGNAL(clicked()), this, SLOT(selectDestDirClicked()));
	connect(ui->convertButton, SIGNAL(pressed()), this, SLOT(onViewInfoButtonPressed()));
}

void ConvertExcelToSql::selectFileEditClicked() {
	if (openFilePath.isEmpty()) {
		openFilePath = QDir::currentPath();
	}
	QString fileName = QFileDialog::getOpenFileName(nullptr,
		tr("Open Excel"),
		openFilePath,
		tr("Excel") + " (*.xlsx *.csv)");
	ui->excelFileLineEdit->setText(fileName);
	openFilePath = fileName;
}

void ConvertExcelToSql::selectDestDirClicked() {
	if (destFilePath.isEmpty()) {
		destFilePath = QDir::currentPath();
	}
	QString folderPath = QFileDialog::getExistingDirectory(nullptr,
		QObject::tr("Open Directory"), destFilePath);
	ui->destDirEdit->setText(folderPath);
	destFilePath = folderPath;
}

int ConvertExcelToSql::createDestFileAndSave() {
	QString dir = ui->destDirEdit->text();
	time_t rawtime = QDateTime::currentDateTime().toTime_t();
	destFilePath = dir + "/" + QString::number(rawtime) + ".sql";
	QFile file(destFilePath);
	if (!file.open(QIODevice::WriteOnly)) {
		QMessageBox::warning(nullptr, tr("Warning"),
			QString(tr("Create %1 error!")).arg(destFilePath));
	}
	return 0;
}

int ConvertExcelToSql::writeToDestTxt() {
	//QXlsx::Document sourceXlsx(openFilePath);
	using namespace QXlsx;
	Document sourceXlsx(openFilePath);
	QStringList sheetlist = sourceXlsx.sheetNames();
	if (sheetlist.size() == 0) {
		QMessageBox::warning(nullptr, tr("Warning"),
			QString(tr("Open %1 error or the file is void!")).arg(openFilePath));
		return -1;
	}

	QXlsx::Worksheet* sourceWorksheet = sourceXlsx.currentWorksheet();
	QXlsx::CellRange sourceRange = sourceWorksheet->dimension();

	int iRow = 0;
	int iColumn = 0;
	QFile file(destFilePath);
	if (!file.open(QIODevice::WriteOnly)) {
		QMessageBox::warning(nullptr, tr("Warning"),
			QString(tr("Open %1 error!")).arg(destFilePath));
		return -1;
	}
	QString result{};
	for (iRow = sourceRange.firstRow(); iRow <= sourceRange.lastRow(); iRow++) {
		if (iRow == 1) {
			result += "INSERT INTO " + ui->tableNameEdit->text() + " (";
		} else {
			result += "(";
		}
		for (iColumn = sourceRange.firstColumn(); iColumn <= sourceRange.lastColumn(); iColumn++) {
			QVariant sourceVar = sourceWorksheet->read(iRow, iColumn);
			if (iRow == 1) {
				if (iColumn == sourceRange.lastColumn()) {
					result += sourceVar.toString() + ")";
				} else {
					result += sourceVar.toString() + ", ";
				}
			} else {
				QString tmpStr = sourceVar.toString();
				if (tmpStr == "") {
					tmpStr = "NULL";
					if (iColumn == sourceRange.lastColumn()) {
						result += tmpStr + ")";
					} else {
						result += tmpStr + ", ";
					}
				} else {
					if (iColumn == sourceRange.lastColumn()) {
						result += "\'" + tmpStr + "\'" + ")";
					} else {
						result += "\'" + tmpStr + "\'" + ", ";
					}
				}

			}
		}
		if (iRow == 1) {
			result += " VALUES\n";
		} else if (iRow == sourceRange.lastRow()) {
			result += ";\n";
		} else {
			result += ",\n";
		}
		file.write(result.toUtf8());
		result.clear();
	}
	convertInfo += "Excel lines : " + QString::number(sourceRange.lastRow());
}

void ConvertExcelToSql::onViewInfoButtonPressed() {
	createDestFileAndSave();
	writeToDestTxt();
	convertInfo += "\nDest : " + destFilePath + "\nConvert completed";
	emit updateNeeded(convertInfo);
}
