# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'PreferenceUI.ui'
##
## Created by: Qt User Interface Compiler version 6.5.3
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QFrame, QSizePolicy, QWidget)

class Ui_Form(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.resize(400, 300)
        self.frame = QFrame(Form)
        self.frame.setObjectName(u"frame")
        self.frame.setGeometry(QRect(10, 10, 261, 101))
        self.frame.setFrameShape(QFrame.Box)
        self.frame.setFrameShadow(QFrame.Raised)
        self.IncludeAngleUnit = Gui_PrefCheckBox(self.frame)
        self.IncludeAngleUnit.setObjectName(u"IncludeAngleUnit")
        self.IncludeAngleUnit.setGeometry(QRect(10, 60, 121, 17))
        self.IncludeAngleUnit.setProperty("prefEntry", u"IncludeAngle")
        self.IncludeAngleUnit.setProperty("prefPath", u"Mod\\TechDrawTitleBlockUtility\\UI")
        self.IncludeLengthUnit = Gui_PrefCheckBox(self.frame)
        self.IncludeLengthUnit.setObjectName(u"IncludeLengthUnit")
        self.IncludeLengthUnit.setGeometry(QRect(10, 40, 121, 17))
        self.IncludeLengthUnit.setProperty("prefEntry", u"IncludeLength")
        self.IncludeLengthUnit.setProperty("prefPath", u"Mod\\TechDrawTitleBlockUtility\\UI")
        self.urlLabel = Gui_UrlLabel(self.frame)
        self.urlLabel.setObjectName(u"urlLabel")
        self.urlLabel.setGeometry(QRect(10, 0, 261, 41))
        font = QFont()
        font.setBold(True)
        self.urlLabel.setFont(font)
        self.urlLabel.setTextFormat(Qt.RichText)
        self.urlLabel.setWordWrap(True)
        self.IncludeMassUnit = Gui_PrefCheckBox(self.frame)
        self.IncludeMassUnit.setObjectName(u"IncludeMassUnit")
        self.IncludeMassUnit.setGeometry(QRect(10, 80, 121, 17))
        self.IncludeMassUnit.setProperty("prefEntry", u"IncludeMass")
        self.IncludeMassUnit.setProperty("prefPath", u"Mod\\TechDrawTitleBlockUtility\\UI")
        self.frame_2 = QFrame(Form)
        self.frame_2.setObjectName(u"frame_2")
        self.frame_2.setGeometry(QRect(10, 120, 261, 61))
        self.frame_2.setFrameShape(QFrame.Box)
        self.frame_2.setFrameShadow(QFrame.Raised)
        self.UseExternalSource = Gui_PrefCheckBox(self.frame_2)
        self.UseExternalSource.setObjectName(u"UseExternalSource")
        self.UseExternalSource.setGeometry(QRect(10, 10, 231, 17))
        self.UseExternalSource.setProperty("prefEntry", u"UseExternalSource")
        self.UseExternalSource.setProperty("prefPath", u"Mod\\\\TechDrawTitleBlockUtility\\\\UI")
        self.ExternalFileChooser = Gui_PrefFileChooser(self.frame_2)
        self.ExternalFileChooser.setObjectName(u"ExternalFileChooser")
        self.ExternalFileChooser.setGeometry(QRect(10, 30, 241, 23))
        self.ExternalFileChooser.setProperty("prefEntry", u"ExternalFile")
        self.ExternalFileChooser.setProperty("prefPath", u"Mod\\TechDrawTitleBlockUtility\\UI")
        self.frame_3 = QFrame(Form)
        self.frame_3.setObjectName(u"frame_3")
        self.frame_3.setGeometry(QRect(10, 190, 261, 91))
        self.frame_3.setFrameShape(QFrame.Box)
        self.frame_3.setFrameShadow(QFrame.Raised)
        self.UseFileName = Gui_PrefCheckBox(self.frame_3)
        self.UseFileName.setObjectName(u"UseFileName")
        self.UseFileName.setGeometry(QRect(10, 10, 241, 17))
        self.UseFileName.setProperty("prefEntry", u"UseFileName")
        self.UseFileName.setProperty("prefPath", u"Mod\\TechDrawTitleBlockUtility\\UI")
        self.urlLabel_2 = Gui_UrlLabel(self.frame_3)
        self.urlLabel_2.setObjectName(u"urlLabel_2")
        self.urlLabel_2.setGeometry(QRect(10, 40, 131, 16))
        self.DrawingNumber = Gui_AccelLineEdit(self.frame_3)
        self.DrawingNumber.setObjectName(u"DrawingNumber")
        self.DrawingNumber.setGeometry(QRect(10, 60, 241, 20))
        self.DrawingNumber.setInputMethodHints(Qt.ImhNone)
        self.DrawingNumber.setClearButtonEnabled(True)

        self.retranslateUi(Form)

        QMetaObject.connectSlotsByName(Form)
    # setupUi

    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Form", None))
        self.IncludeAngleUnit.setText(QCoreApplication.translate("Form", u"Include angles", None))
        self.IncludeLengthUnit.setText(QCoreApplication.translate("Form", u"Include lengths", None))
        self.urlLabel.setText(QCoreApplication.translate("Form", u"The following items can be included even when the titleblock does not have them:", None))
        self.IncludeMassUnit.setText(QCoreApplication.translate("Form", u"Include mass", None))
        self.UseExternalSource.setText(QCoreApplication.translate("Form", u"Use external source", None))
        self.ExternalFileChooser.setFilter("")
        self.UseFileName.setText(QCoreApplication.translate("Form", u"Use filename as drawingnumber", None))
        self.urlLabel_2.setText(QCoreApplication.translate("Form", u"Field for drawing number", None))
        self.DrawingNumber.setInputMask("")
        self.DrawingNumber.setPlaceholderText(QCoreApplication.translate("Form", u"Enter field name", None))
    # retranslateUi

