#ifndef EXQLINEEDIT_H
#define EXQLINEEDIT_H

#include <QLineEdit>
#include<QMouseEvent>

class ExQLineEdit :public QLineEdit
{
	Q_OBJECT
public:
	explicit ExQLineEdit(QWidget* parent = 0) {
	}
	~ExQLineEdit() {
	}
protected:
	virtual void mousePressEvent(QMouseEvent* e) {
		if (e->button() == Qt::LeftButton) {
			emit clicked();
		}
	}
signals:
	void clicked();
};

#endif
