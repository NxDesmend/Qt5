#include <thread>

#include "MergeSheets.h"
#include "ui_plugin.h"

using namespace QXlsx;

const QString DEST_SHEET_NAME = "myMergeSheet";

MergeSheets::MergeSheets()
{
	
}

MergeSheets::~MergeSheets() {
}

QString MergeSheets::title() {
	return " MergeSheetsEx";
}

QString MergeSheets::version() {
	return m_version;
}

QString MergeSheets::description() {
	return "";
}

QString MergeSheets::help() {
	return "";
}

void MergeSheets::processExcel() {
	// Todo
}

void MergeSheets::init(QWidget* parent)
{
	ui = new Ui::PluginGui;
	ui->setupUi(parent);

	connect(ui->selectFileEdit, SIGNAL(clicked()), this, SLOT(selectFileEditClicked()));
	connect(ui->selectSaveDirEdit, SIGNAL(clicked()), this, SLOT(selectSaveDirEditClicked()));
	connect(ui->viewInfoButton, SIGNAL(pressed()), this, SLOT(onViewInfoButtonPressed()));
	connect(ui->mergeSheetsButton, SIGNAL(pressed()), this, SLOT(onMergeSheetsButtonPressed()));
}

void MergeSheets::selectFileEditClicked() {
	if (openFilePath.isEmpty()) {
		openFilePath = QDir::currentPath();
	}
	QString fileName = QFileDialog::getOpenFileName(nullptr,
		tr("Open Excel"),
		openFilePath,
		tr("Excel") + " (*.xlsx *.csv)");
	ui->selectFileEdit->setText(fileName);
	openFilePath = fileName;
}

void MergeSheets::selectSaveDirEditClicked() {
	if (saveFilePath.isEmpty()) {
		saveFilePath = QDir::currentPath();
	}
	QString folderPath = QFileDialog::getExistingDirectory(nullptr,
		QObject::tr("Open Directory"), saveFilePath);
	ui->selectSaveDirEdit->setText(folderPath);
	saveFilePath = folderPath;
}

int MergeSheets::viewFileInfo() {
	Document sourceXlsx(ui->selectFileEdit->text());
	QStringList sheetlist = sourceXlsx.sheetNames();

	excelInfo += "Total numbe of sheets : " + QString::number(sheetlist.size());
	for (auto it : sheetlist) {
		if (it == DEST_SHEET_NAME) {
			continue;
		}
		sourceXlsx.selectSheet(it);
		Worksheet* tmpSheet = sourceXlsx.currentWorksheet();
		CellRange tmpRange = tmpSheet->dimension();

		excelInfo += "\n" + tmpSheet->sheetName() + ", " + QString("rows : ") +
			QString::number(tmpRange.lastRow()) + ", " + QString("columns : ") +
			QString::number(tmpRange.lastColumn());
	}
	return 0;
}

void MergeSheets::onViewInfoButtonPressed() {
	viewFileInfo();

	emit updateNeeded(excelInfo);
	excelInfo.clear();
}

int MergeSheets::threadProcess(MergeSheets* ms) {
	ms->processSheets();
	return 0;
}

void MergeSheets::onMergeSheetsButtonPressed() {
	std::thread pt(threadProcess, this);
	pt.detach();
}

int MergeSheets::processSheets() {
	mergeInfo += "Sheets are being merged ...";
	emit updateNeeded(mergeInfo);

	createDestFileAndSave();
	writeToDest();

	mergeInfo += "\nMerge Completed\nDest:" + destFilePath;
	emit updateNeeded(mergeInfo);
	mergeInfo.clear();
	return 0;
}

int MergeSheets::createDestFileAndSave() {
	QString dir = ui->selectSaveDirEdit->text();
	time_t rawtime = QDateTime::currentDateTime().toTime_t();
	destFilePath = dir + "/" + QString::number(rawtime) + ".xlsx";

	Document sourceXlsx(ui->selectFileEdit->text());
	sourceXlsx.addSheet(DEST_SHEET_NAME);
	sourceXlsx.saveAs(destFilePath);
	return 0;
}

int MergeSheets::writeToDest() {
	Document destXlsx(destFilePath);
	QStringList sheetlist = destXlsx.sheetNames();
	destXlsx.selectSheet(DEST_SHEET_NAME);
	Worksheet* destWorksheet = destXlsx.currentWorksheet();

	int destRow = 0;
	for (auto it : sheetlist) {
		if (it == DEST_SHEET_NAME) {
			continue;
		}
		destXlsx.selectSheet(it);
		Worksheet* tmpSheet = destXlsx.currentWorksheet();

		writeToDestSheet(tmpSheet, destWorksheet, destRow);
	}
	for (auto it : sheetlist) {
		if (it != DEST_SHEET_NAME) {
			destXlsx.deleteSheet(it);
		}
	}
	destXlsx.save();
	return 0;
}

int MergeSheets::writeToDestSheet(Worksheet* sourceSheet, 
	Worksheet* destSheet, int& destRow) {
	CellRange sourceRange = sourceSheet->dimension();

	mergeInfo += "\n" + sourceSheet->sheetName() + ", " + QString("rows : ") +
		QString::number(sourceRange.lastRow()) + ", " + QString("columns : ") +
		QString::number(sourceRange.lastColumn());

	int iRow = 0;
	int iColumn = 0;
	for (iRow = sourceRange.firstRow(); iRow <= sourceRange.lastRow(); iRow++) {
		destRow++;
		for (iColumn = sourceRange.firstColumn(); iColumn <= sourceRange.lastColumn(); iColumn++) {
			QVariant sourceVar = sourceSheet->read(iRow, iColumn);
			destSheet->write(destRow, iColumn, sourceVar);
		}
	}
	return 0;
}
