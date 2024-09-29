#include "MergeSheets.h"
#include "ui_plugin.h"
#include "util.h"

const QString DEST_SHEET_NAME = "myMergeSheet";

MergeSheets::MergeSheets()
{
	
}

MergeSheets::~MergeSheets() {
	closeQAxObj();
	closeDestQAxObj();
}

int MergeSheets::initSourceQAxObj() {
	sourceExcel = new QAxObject("Excel.Application");
	sourceWorkbooks = sourceExcel->querySubObject("WorkBooks");
	sourceWorkbook = sourceWorkbooks->querySubObject("Open(QString&)", ui->selectOpenFileEdit->text());
	return 0;
}

int MergeSheets::initDestQAxObj() {
	destExcel = new QAxObject("Excel.Application");
	destWorkbooks = destExcel->querySubObject("WorkBooks");
	destWorkbook = destWorkbooks->querySubObject("Open(QString&)", destFilePath);

	destWorksheets = destWorkbook->querySubObject("WorkSheets");
	if (destWorksheets == nullptr) {
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid worksheets!")));
		return -1;
	}
	int iWorkSheet = destWorksheets->property("Count").toInt();
	if (iWorkSheet < 1) {
		return -1;
		QMessageBox::warning(nullptr, QString(tr("Error")), QString(tr("iWorkSheet is wrong!")));
	}
	destWorksheet = destWorksheets->querySubObject("Item(int)", iWorkSheet);
	return 0;
}

int MergeSheets::closeQAxObj() {
	if (sourceWorkbook != nullptr) {
		sourceWorkbook->dynamicCall("Close()");
	}

	if (sourceExcel != nullptr)
	{
		sourceExcel->dynamicCall("Quit()");
		delete sourceExcel;
		sourceExcel = nullptr;

		sourceWorkbooks = nullptr;
		sourceWorkbook = nullptr;
	}
	return 0;
}

int MergeSheets::closeDestQAxObj() {
	if (destWorkbook != nullptr) {
		destWorkbook->dynamicCall("Save()");
		destWorkbook->dynamicCall("Close()");
	}

	if (destExcel != nullptr)
	{
		destExcel->dynamicCall("Quit()");
		destExcel->setParent(nullptr);
		delete destExcel;
		destExcel = nullptr;

		destWorkbooks = nullptr;
		destWorkbook = nullptr;
		destWorksheets = nullptr;
		destWorksheet = nullptr;
	}
	return 0;
}

QString MergeSheets::title() {
	return this->metaObject()->className();
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

	connect(ui->selectOpenFileEdit, SIGNAL(clicked()), this, SLOT(selectFileEditClicked()));
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
	ui->selectOpenFileEdit->setText(fileName);
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

int MergeSheets::openFile() {
	if (ui->selectOpenFileEdit->text().right(4) != "xlsx") {
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid excel file!")));
		return -1;
	}
	if (sourceExcel == nullptr) {
		initSourceQAxObj();
	}
	sourceExcel->dynamicCall("SetVisible(bool)", true);
	if (sourceWorkbook == nullptr) {
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid excel file!")));
		return -1;
	}
	QAxObject* worksheets = sourceWorkbook->querySubObject("WorkSheets");
	if (worksheets == nullptr) {
		closeQAxObj();
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid worksheets!")));
		return -1;
	}

	int iWorkSheet = worksheets->property("Count").toInt();
	QString countInfo = QString("Total number of sheets:") + QString::number(iWorkSheet);
	excelInfo += countInfo;
	int i = 1;
	while (i <= iWorkSheet) {
		QAxObject* worksheet = worksheets->querySubObject("Item(int)", i);
		readSheetInfo(worksheet, i);
		i++;
	}
	return 0;
}

int MergeSheets::readSheetInfo(QAxObject* worksheet, int serial) {
	if (worksheet == nullptr) {
		QMessageBox::warning(nullptr, QString(tr("Error")), QString(tr("worksheet is nullptr!")));
		return -1;
	}
	QString sheetInfo = "\n" + QString("sheet") + QString::number(serial) + ", ";
	QAxObject* usedrange = worksheet->querySubObject("UsedRange");

	QAxObject* rows = usedrange->querySubObject("Rows");
	int iRows = rows->property("Count").toInt();
	QString rowsInfo = QString("rows:") + QString::number(iRows) + ", ";

	QAxObject* columns = usedrange->querySubObject("Columns");
	int iColumns = columns->property("Count").toInt();
	QString columnsInfo = QString("columns:") + QString::number(iColumns);

	excelInfo += (sheetInfo + rowsInfo + columnsInfo);

	return 0;
}

void MergeSheets::onViewInfoButtonPressed() {
	ui->viewInfoButton->setEnabled(false);
	openFile();
	ui->viewInfoButton->setEnabled(true);

	emit updateNeeded(excelInfo);
	excelInfo.clear();
}

void MergeSheets::onMergeSheetsButtonPressed() {
	ui->mergeSheetsButton->setEnabled(false);
	createExcelFile();
	processSheets();
	ui->mergeSheetsButton->setEnabled(true);

	mergeInfo += "\nDest:" + destFilePath;
	emit updateNeeded(mergeInfo);
	mergeInfo.clear();
}

int MergeSheets::createExcelFile() {
	QString dir = ui->selectSaveDirEdit->text();
	time_t rawtime = QDateTime::currentDateTime().toTime_t();

	destFilePath = dir + "/" +  QString::number(rawtime) + ".xlsx";
	QAxObject* excel = new QAxObject("Excel.Application");
	excel->dynamicCall("SetVisible (bool Visible)", "false");

	QAxObject* workbooks = excel->querySubObject("WorkBooks");
	workbooks->dynamicCall("Add");
	QAxObject* workbook = excel->querySubObject("ActiveWorkBook");

	QAxObject* worksheets = workbook->querySubObject("WorkSheets");
	int count = worksheets->property("Count").toInt();
	QAxObject* worksheet = worksheets->querySubObject("Item(int)", count);
	worksheet->setProperty("Name", DEST_SHEET_NAME);

	workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(destFilePath));
	workbook->dynamicCall("Close()");
	excel->dynamicCall("Quit()");
	delete excel;
	excel = NULL;
	return 0;
}

int MergeSheets::getRowAndColumn(QAxObject* worksheet, int& pRows, int& pColumns) {
	QAxObject* usedrange = worksheet->querySubObject("UsedRange");

	QAxObject* rows = usedrange->querySubObject("Rows");
	pRows = rows->property("Count").toInt();
	mergeInfo += QString("rows:") + QString::number(pRows) + ", ";
	QAxObject* columns = usedrange->querySubObject("Columns");
	pColumns = columns->property("Count").toInt();
	mergeInfo += QString("columns:") + QString::number(pColumns);

	return 0;
}

int MergeSheets::processSheets() {
	SheetData sd;
	if (sourceExcel == nullptr) {
		initSourceQAxObj();
	}
	sourceExcel->dynamicCall("SetVisible(bool)", false);
	QAxObject* worksheets = sourceWorkbook->querySubObject("WorkSheets");
	if (worksheets == nullptr) {
		closeQAxObj();
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid worksheets!")));
		return -1;
	}

	int iWorkSheet = worksheets->property("Count").toInt();
	QString countInfo = QString("Total number of sheets:") + QString::number(iWorkSheet);
	mergeInfo += countInfo;
	int i = 1;
	while (i <= iWorkSheet) {
		QAxObject* worksheet = worksheets->querySubObject("Item(int)", i);
		mergeInfo += "\n" + QString("sheet") + QString::number(i) + ", ";

		QAxObject* usedrange = worksheet->querySubObject("UsedRange");

		QAxObject* rows = usedrange->querySubObject("Rows");
		int iRows = rows->property("Count").toInt();
		mergeInfo += QString("rows:") + QString::number(iRows) + ", ";
		QAxObject* columns = usedrange->querySubObject("Columns");
		int iColumns = columns->property("Count").toInt();
		mergeInfo += QString("columns:") + QString::number(iColumns);

		sd.rows = iRows;
		sd.columns = iColumns;
		rectWriteToDestExcel(worksheet, sd);
		i++;
	}
	closeQAxObj();
	closeDestQAxObj();
	return 0;
}

bool MergeSheets::isDestPointerValid() {
	if (destExcel == nullptr) {
		if (initDestQAxObj() != 0) {
			QMessageBox::warning(nullptr, QString(tr("Error")), QString(tr("initDestQAxObj error!")));
			return false;
		}
		destExcel->dynamicCall("SetVisible(bool)", false);
	}
	if (destWorkbook == nullptr) {
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid excel file!")));
		return false;
	}
	if (destWorksheets == nullptr) {
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid destWorksheets!")));
		return false;
	}

	if (destWorksheet == nullptr) {
		QMessageBox::warning(nullptr, QString(tr("Warning")), QString(tr("Invalid destWorksheet!")));
		return false;
	}
	return true;
}

int MergeSheets::writeToDestExcel(QAxObject* worksheet, SheetData& sd) {
	if (!isDestPointerValid()) {
		return -1;
	}
	int destRows = 0;
	QAxObject* cells = worksheet->querySubObject("Cells");
	for (int row = 1; row <= sd.rows; ++row) {
		for (int col = 1; col <= sd.columns; ++col) {
			// read one cell a time from source
			QAxObject* sourceCell = cells->querySubObject("Item(int, int)", row, col);
			QVariant value = sourceCell->dynamicCall("Value2");
			delete sourceCell;

			// write one cell a time from source
			QAxObject* destCell = destWorksheet->querySubObject("Cells(int, int)", sd.destRows, col);
			destCell->dynamicCall("SetValue(const QString &)", value);
			delete destCell;
		}
		sd.destRows++;
	}

	return 0;
}

int MergeSheets::lineWriteToDestExcel(QAxObject* worksheet, SheetData& sd) {
	if (!isDestPointerValid()) {
		return -1;
	}

	for (int row = 1; row <= sd.rows; ++row) {
		// read one line at a time from source
		QAxObject* sourceLine = worksheet->querySubObject("Range(QString)",
			dataRange(row, row, 1, sd.columns));
		QList<QVariant> lsitVar = sourceLine->property("Value").toList();

		// write one line at a time from source
		QAxObject* destLine = destWorksheet->querySubObject("Range(QString)",
			dataRange(sd.destRows, sd.destRows, 1, sd.columns));
		destLine->setProperty("Value", lsitVar);
		sd.destRows++;
	}
	return 0;
}

int MergeSheets::rectWriteToDestExcel(QAxObject* worksheet, SheetData& sd) {
	if (!isDestPointerValid()) {
		return -1;
	}
	// read  one rect at a time from source 
	QAxObject* sourceRect = worksheet->querySubObject("Range(QString)",
		dataRange(1, sd.rows, 1, sd.columns));
	QVariant var = sourceRect->property("Value");

	// write one rect at a time from source
	QAxObject* destRect = destWorksheet->querySubObject("Range(QString)",
		dataRange(sd.destRows, sd.destRows + sd.rows - 1, 1, sd.columns));
	destRect->setProperty("Value", var);

	// use copy function
	// destRect->dynamicCall("Copy(QVariant)", sourceRect->asVariant());

	sd.destRows += sd.rows;
	return 0;
}
