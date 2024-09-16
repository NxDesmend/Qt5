#include "primaryMathTest.h"

primaryMathTest::primaryMathTest(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);
}

primaryMathTest::~primaryMathTest()
{}

void primaryMathTest::on_scorePushButton_pressed() {
    int score = 0;
    QString  wrongAnwsers{};
    if (ui.lineEdit125_5->text() == "25") {
        score += 10;
    } else {
        wrongAnwsers += "1";
    }

    if (ui.lineEdit702_9->text() == "78") {
        score += 10;
    } else {
        wrongAnwsers += ",2";
    }

    if (ui.lineEdit663_13->text() == "51") {
        score += 10;
    }
    else {
        wrongAnwsers += ",3";
    }

    if (ui.lineEdit234_26->text() == "9") {
        score += 10;
    }
    else {
        wrongAnwsers += ",4";
    }

    if (ui.lineEdit375_25->text() == "15") {
        score += 10;
    }
    else {
        wrongAnwsers += ",5";
    }

    if (ui.lineEdit481_13->text() == "37") {
        score += 10;
    }
    else {
        wrongAnwsers += ",6";
    }

    if (ui.lineEdit986_34->text() == "29") {
        score += 10;
    }
    else {
        wrongAnwsers += ",7";
    }

    if (ui.lineEdit525_21->text() == "25") {
        score += 10;
    }
    else {
        wrongAnwsers += ",8";
    }
    if (ui.lineEdit897_23->text() == "39") {
        score += 10;
    }
    else {
        wrongAnwsers += ",9";
    }
    if (ui.lineEdit726_33->text() == "22") {
        score += 10;
    }
    else {
        wrongAnwsers += ",10";
    }
    if (wrongAnwsers[0] == ',') {
        QString  tmpWrongAnwsers;
        int i = 0;
        for (i = 1; i < wrongAnwsers.size(); i++) {
            tmpWrongAnwsers.push_back(wrongAnwsers[i]);
        }
        wrongAnwsers = tmpWrongAnwsers;
    }
    ui.scoreLineEdit->setText(QString::fromStdString(std::to_string(score)));
    ui.wrongLineEdit->setText(wrongAnwsers);
}

void primaryMathTest::on_answerPushButton_pressed() {
    QString anwsers = "[1]25,[2]78,[3]51,[4]9,[5]15,[6]37,[7]29,[8]25,[9]39,[10]22";
    if (ui.inputLineEdit->text() == "123456") {
        ui.answerTextEdit->setText(anwsers);
    } else {
        ui.answerTextEdit->setText("password error");
    }
}
