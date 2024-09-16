#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_primaryMathTest.h"

class primaryMathTest : public QMainWindow
{
    Q_OBJECT

public:
    primaryMathTest(QWidget *parent = nullptr);
    ~primaryMathTest();

private slots:
    void on_scorePushButton_pressed();
    void on_answerPushButton_pressed();

private:
    Ui::primaryMathTestClass ui;
};
