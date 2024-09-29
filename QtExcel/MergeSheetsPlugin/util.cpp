#include "util.h"

QString convertToExcelColumnLetter(int number) {
    if (number < 1 || number > 26 * 26 * 26) {
        return QString();
    }

    QString letter;
    if (number > 26) {
        letter = convertToExcelColumnLetter((number - 1) / 26);
    }

    return letter + QChar('A' + (number - 1) % 26);
}

QString dataRange(int rowStart, int rowEnd, int columnStart, int columnEnd) {
    QString ret = convertToExcelColumnLetter(columnStart) + QString::number(rowStart) + ":" +
        convertToExcelColumnLetter(columnEnd) + QString::number(rowEnd);
    return ret;
}
