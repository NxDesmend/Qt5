#include "MainWindow.h"

const QString PLUGINS_SUBFOLDER = "/explugins/";
const QString FILE_ON_DISK_DYNAMIC_PROPERTY = "absolute_file_path";

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent) {
    ui.setupUi(this);

    loadSettings();

    populatePluginsMenu();
}

MainWindow::~MainWindow() {
}

void MainWindow::loadSettings() {
    QSettings settings("Desmend", "QtExcel", this);
    currentPluginFile = settings.value("currentPluginFile", "").toString();
}

void MainWindow::saveSettings() {
    QSettings settings("Desmend", "QtExcel", this);
    settings.setValue("currentPluginFile", currentPluginFile);
}

void MainWindow::closeEvent(QCloseEvent* event)
{
    int result = QMessageBox::warning(this, tr("Exit"),
        tr("Are you sure you want to exit?"), QMessageBox::Yes, QMessageBox::No);
    if (result == QMessageBox::Yes)
    {
        saveSettings();
        event->accept();
    }
    else
    {
        event->ignore();
    }
}

void MainWindow::populatePluginsMenu() {
    // Load all plugins and populate the menus
    QDir pluginsDir(qApp->applicationDirPath() + PLUGINS_SUBFOLDER);
    QFileInfoList pluginFiles = pluginsDir.entryInfoList(QDir::NoDotAndDotDot | QDir::Files, QDir::Name);
    foreach(QFileInfo pluginFile, pluginFiles)
    {
        if (QLibrary::isLibrary(pluginFile.absoluteFilePath())) {
            QPluginLoader pluginLoader(pluginFile.absoluteFilePath(), this);
            if (EXPluginInterface* plugin = dynamic_cast<EXPluginInterface*>(pluginLoader.instance())) {
                QAction* pluginAction = ui.menuPlugins->addAction(plugin->title());
                pluginAction->setProperty(FILE_ON_DISK_DYNAMIC_PROPERTY.toStdString().c_str(), pluginFile.absoluteFilePath());
                connect(pluginAction, SIGNAL(triggered(bool)), this, SLOT(onPluginActionTriggered(bool)));
                if (currentPluginFile == pluginFile.absoluteFilePath())
                {
                    pluginAction->trigger();
                }
            } else {
                QMessageBox::warning(this, tr("Warning"),
                    QString(tr("Make sure %1 is a correct plugin for this application<br>"
                        "and it's not in use by some other application!")).arg(pluginFile.fileName()));
            }
        }
        else {
            /* QMessageBox::warning(this, tr("Warning"),
                QString(tr("Make sure only plugins exist in %1 folder.<br>"
                    "%2 is not a plugin."))
                .arg(PLUGINS_SUBFOLDER)
                .arg(pluginFile.fileName())); */
            // to do
        }
    }
    if (ui.menuPlugins->actions().count() <= 0)
    {
        QMessageBox::critical(this, tr("No Plugins"), QString(tr("This application cannot work without plugins!"
            "<br>Make sure that %1 folder exists "
            "in the same folder as the application<br>and that "
            "there are some filter plugins inside it")).arg(PLUGINS_SUBFOLDER));
        this->setEnabled(false);
    }
}

void MainWindow::onPluginActionTriggered(bool) {
    if (!currentPlugin.isNull())
    {
        delete currentPlugin;
        delete currentPluginGui;
    }

    currentPluginFile = QObject::sender()->property(FILE_ON_DISK_DYNAMIC_PROPERTY.toStdString().c_str()).toString();
    currentPlugin = new QPluginLoader(currentPluginFile, this);
    currentPluginGui = new QWidget(this);
    ui.pluginLayout->addWidget(currentPluginGui);

    EXPluginInterface* currentPluginInstance = dynamic_cast<EXPluginInterface*>(currentPlugin->instance());
    if (currentPluginInstance)
    {
        currentPluginInstance->init(currentPluginGui);
        connect(currentPlugin->instance(), SIGNAL(updateNeeded(QString)), this, SLOT(onCurrentPluginUpdateNeeded(QString)));
    }
}

void MainWindow::onCurrentPluginUpdateNeeded(QString info) {
    ui.infoTextEdit->setText(info);
}
