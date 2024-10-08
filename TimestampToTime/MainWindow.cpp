#include <ctime>
#include <iostream>

#include <QDateTime>
#include <QMessagebox>
#include <QString>
#include <QTimer>

#include "MainWindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent) {
    ui.setupUi(this);

    timer = new QTimer(this);
    connect(timer, &QTimer::timeout, this, &MainWindow::updateCaption);
    connect(ui.startStopButton, SIGNAL(pressed()), this,
        SLOT(on_startStopButton_pressked()));
    timer->start(1000);

    connect(ui.tsToTimeConvertButton, SIGNAL(pressed()),this,
        SLOT(on_tsToTimeConvertButton_pressked()));
    connect(ui.fromBjConvertButton, SIGNAL(pressed()), this,
        SLOT(on_fromBjConvertButton__pressked()));
}

MainWindow::~MainWindow() {
}

int MainWindow::updateCaption() {
    std::time_t now = std::time(nullptr);

    ui.curTimeStampEdit->setText(QString::number(now));
    return 0;
}

int MainWindow::on_startStopButton_pressed() {
    if (!isCurStop) {
        isCurStop = true;
        timer->stop();
        ui.startStopButton->setText("start");
    } else {
        isCurStop = false;
        timer->start(1000);
        ui.startStopButton->setText("stop");
    }
    return 0;
}

int MainWindow::on_tsToTimeConvertButton_pressed() {
    QString strTimestamp = ui.fromTsEdit->text();
    if (!checkTimestamp(strTimestamp)) {
        return -1;
    }
    time_t timestamp = strTimestamp.toInt();
    QDateTime dateTime = QDateTime::fromMSecsSinceEpoch(timestamp * 1000);
    QString strTime = dateTime.toString("yyyy-MM-dd hh:mm:ss");
    ui.bjTimeLineEdit->setText(strTime);
    return 0;
}

bool MainWindow::checkTimestamp(QString timestamp) {
    if (timestamp.size() != 10) {
        QMessageBox::warning(this, tr("Warn"), tr("Timestamp length error"));
        return false;
    }
    bool isInt = false;
    time_t intTimestamp = timestamp.toInt(&isInt);
    if (!isInt) {
        QMessageBox::warning(this, tr("Warn"), tr("Timestamp is not int"));
        return false;
    }
    return true;
}

int MainWindow::on_fromBjConvertButton_pressed() {
    QString strTime = ui.fromBjTimeEdit->text();
    if (!checkStrTime(strTime)) {
        return -1;
    }
    QDateTime dateTime = QDateTime::fromString(strTime, "yyyy-MM-dd hh:mm:ss");
    if (!dateTime.isValid()) {
        QMessageBox::warning(this, tr("Warn"), tr("Input is not a time!"));
        return -1;
    }
    time_t timestamp = dateTime.toMSecsSinceEpoch();
    ui.toTimeStampEdit->setText(QString::number(timestamp / 1000));
}

bool MainWindow::checkStrTime(QString strTime) {
    if (strTime.length() != 19) {
        QMessageBox::warning(this, tr("Warn"), tr("Timestamp length error"));
        return false;
    }

    return true;
}


