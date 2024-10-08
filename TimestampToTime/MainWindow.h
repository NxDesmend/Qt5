#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_MainWindow.h"

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private:
    bool checkTimestamp(QString timestamp);
    bool checkStrTime(QString strTime);

public slots:
    int updateCaption();
    int on_startStopButton_pressed();
    int on_tsToTimeConvertButton_pressed();
    int on_fromBjConvertButton_pressed();

private:
    Ui::MainWindow ui;

    QTimer* timer{ nullptr };
    bool isCurStop{ false };
};
