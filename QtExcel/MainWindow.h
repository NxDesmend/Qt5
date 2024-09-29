#pragma once

#include <QAxObject>
#include <QCloseEvent>
#include <QDir>
#include <QLibrary>
#include <QMessageBox>
#include <QPluginLoader>
#include <QSettings>
#include <QThread>
#include <QtWidgets/QMainWindow>

#include "ui_MainWindow.h"
#include "explugininterface.h"

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

protected:
    void closeEvent(QCloseEvent* event);

private slots:
    void onPluginActionTriggered(bool);
    void onCurrentPluginUpdateNeeded(QString info);

private:
    void loadSettings();
    void saveSettings();
    void populatePluginsMenu();

private:
    Ui::MainWindow ui;

    QString currentPluginFile{};
    QPointer<QPluginLoader> currentPlugin{nullptr};
    QPointer<QWidget> currentPluginGui{};
};
