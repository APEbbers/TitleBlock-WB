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
from PySide6.QtWidgets import (QApplication, QFrame, QGridLayout, QGroupBox,
    QLabel, QSizePolicy, QSplitter, QVBoxLayout,
    QWidget)

class Ui_Form(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.setWindowModality(Qt.ApplicationModal)
        Form.resize(985, 928)
        Form.setContextMenuPolicy(Qt.NoContextMenu)
        Form.setWindowTitle(u"Settings")
        Form.setInputMethodHints(Qt.ImhNone)
        self.groupBox_DrawingNumber = QGroupBox(Form)
        self.groupBox_DrawingNumber.setObjectName(u"groupBox_DrawingNumber")
        self.groupBox_DrawingNumber.setEnabled(True)
        self.groupBox_DrawingNumber.setGeometry(QRect(5, 370, 421, 101))
        font = QFont()
        font.setBold(True)
        self.groupBox_DrawingNumber.setFont(font)
        self.groupBox_DrawingNumber.setFlat(False)
        self.UseFileName = Gui_PrefCheckBox(self.groupBox_DrawingNumber)
        self.UseFileName.setObjectName(u"UseFileName")
        self.UseFileName.setGeometry(QRect(10, 20, 241, 17))
        font1 = QFont()
        font1.setBold(False)
        self.UseFileName.setFont(font1)
        self.UseFileName.setProperty("prefEntry", u"UseFileName")
        self.UseFileName.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.DrawingNumber = Gui_PrefLineEdit(self.groupBox_DrawingNumber)
        self.DrawingNumber.setObjectName(u"DrawingNumber")
        self.DrawingNumber.setEnabled(False)
        self.DrawingNumber.setGeometry(QRect(10, 70, 241, 20))
        self.DrawingNumber.setFont(font1)
        self.DrawingNumber.setProperty("prefEntry", u"DrwNrFieldName")
        self.DrawingNumber.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.label_10 = QLabel(self.groupBox_DrawingNumber)
        self.label_10.setObjectName(u"label_10")
        self.label_10.setEnabled(False)
        self.label_10.setGeometry(QRect(10, 50, 256, 16))
        self.label_10.setFont(font1)
        self.groupBox_External_Source = QGroupBox(Form)
        self.groupBox_External_Source.setObjectName(u"groupBox_External_Source")
        self.groupBox_External_Source.setGeometry(QRect(5, 5, 421, 361))
        self.groupBox_External_Source.setFont(font)
        self.groupBox_External_Source.setFlat(False)
        self.groupBox_External_Source.setCheckable(False)
        self.ExternalFileChooser = Gui_PrefFileChooser(self.groupBox_External_Source)
        self.ExternalFileChooser.setObjectName(u"ExternalFileChooser")
        self.ExternalFileChooser.setEnabled(False)
        self.ExternalFileChooser.setGeometry(QRect(10, 40, 401, 23))
        self.ExternalFileChooser.setFont(font1)
        self.ExternalFileChooser.setProperty("prefEntry", u"ExternalFile")
        self.ExternalFileChooser.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.UseExternalSource = Gui_PrefCheckBox(self.groupBox_External_Source)
        self.UseExternalSource.setObjectName(u"UseExternalSource")
        self.UseExternalSource.setGeometry(QRect(10, 20, 231, 17))
        self.UseExternalSource.setFont(font1)
        self.UseExternalSource.setProperty("prefEntry", u"UseExternalSource")
        self.UseExternalSource.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.StartCell = Gui_PrefLineEdit(self.groupBox_External_Source)
        self.StartCell.setObjectName(u"StartCell")
        self.StartCell.setEnabled(False)
        self.StartCell.setGeometry(QRect(10, 145, 117, 26))
        self.StartCell.setFont(font1)
        self.StartCell.setProperty("prefEntry", u"StartCell")
        self.StartCell.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.label_2 = QLabel(self.groupBox_External_Source)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setEnabled(False)
        self.label_2.setGeometry(QRect(130, 145, 216, 31))
        font2 = QFont()
        font2.setBold(False)
        font2.setItalic(True)
        self.label_2.setFont(font2)
        self.label_2.setWordWrap(True)
        self.AutoFillTitleBlock = Gui_PrefCheckBox(self.groupBox_External_Source)
        self.AutoFillTitleBlock.setObjectName(u"AutoFillTitleBlock")
        self.AutoFillTitleBlock.setEnabled(False)
        self.AutoFillTitleBlock.setGeometry(QRect(10, 190, 221, 17))
        self.AutoFillTitleBlock.setFont(font1)
        self.AutoFillTitleBlock.setChecked(True)
        self.AutoFillTitleBlock.setProperty("prefEntry", u"AutoFillTitleBlock")
        self.AutoFillTitleBlock.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.ImportSettingsXL = Gui_PrefCheckBox(self.groupBox_External_Source)
        self.ImportSettingsXL.setObjectName(u"ImportSettingsXL")
        self.ImportSettingsXL.setEnabled(False)
        self.ImportSettingsXL.setGeometry(QRect(10, 230, 381, 17))
        self.ImportSettingsXL.setFont(font1)
        self.ImportSettingsXL.setProperty("prefEntry", u"ImportSettingsXL")
        self.ImportSettingsXL.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.line = QFrame(self.groupBox_External_Source)
        self.line.setObjectName(u"line")
        self.line.setGeometry(QRect(10, 170, 401, 16))
        self.line.setMaximumSize(QSize(16777215, 16777215))
        self.line.setContextMenuPolicy(Qt.NoContextMenu)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line_2 = QFrame(self.groupBox_External_Source)
        self.line_2.setObjectName(u"line_2")
        self.line_2.setGeometry(QRect(10, 210, 401, 16))
        self.line_2.setFont(font1)
        self.line_2.setFrameShape(QFrame.HLine)
        self.line_2.setFrameShadow(QFrame.Sunken)
        self.splitter = QSplitter(self.groupBox_External_Source)
        self.splitter.setObjectName(u"splitter")
        self.splitter.setGeometry(QRect(10, 270, 396, 20))
        self.splitter.setOrientation(Qt.Horizontal)
        self.SheetName_Settings = Gui_PrefLineEdit(self.splitter)
        self.SheetName_Settings.setObjectName(u"SheetName_Settings")
        self.SheetName_Settings.setEnabled(False)
        self.SheetName_Settings.setFont(font1)
        self.SheetName_Settings.setProperty("prefEntry", u"SheetName_Settings")
        self.SheetName_Settings.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.splitter.addWidget(self.SheetName_Settings)
        self.splitter_2 = QSplitter(self.groupBox_External_Source)
        self.splitter_2.setObjectName(u"splitter_2")
        self.splitter_2.setGeometry(QRect(10, 90, 386, 20))
        self.splitter_2.setOrientation(Qt.Horizontal)
        self.SheetName = Gui_PrefLineEdit(self.splitter_2)
        self.SheetName.setObjectName(u"SheetName")
        self.SheetName.setEnabled(False)
        self.SheetName.setFont(font1)
        self.SheetName.setProperty("prefEntry", u"SheetName")
        self.SheetName.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.splitter_2.addWidget(self.SheetName)
        self.label_13 = QLabel(self.groupBox_External_Source)
        self.label_13.setObjectName(u"label_13")
        self.label_13.setEnabled(False)
        self.label_13.setGeometry(QRect(10, 285, 191, 51))
        self.label_13.setFont(font1)
        self.label_13.setFrameShadow(QFrame.Plain)
        self.label_13.setWordWrap(True)
        self.label_14 = QLabel(self.groupBox_External_Source)
        self.label_14.setObjectName(u"label_14")
        self.label_14.setEnabled(False)
        self.label_14.setGeometry(QRect(10, 250, 371, 16))
        self.label_14.setFont(font1)
        self.label_14.setFrameShadow(QFrame.Plain)
        self.label_14.setWordWrap(True)
        self.label_15 = QLabel(self.groupBox_External_Source)
        self.label_15.setObjectName(u"label_15")
        self.label_15.setEnabled(False)
        self.label_15.setGeometry(QRect(10, 105, 341, 51))
        self.label_15.setFont(font1)
        self.label_15.setFrameShadow(QFrame.Plain)
        self.label_15.setWordWrap(True)
        self.label_16 = QLabel(self.groupBox_External_Source)
        self.label_16.setObjectName(u"label_16")
        self.label_16.setEnabled(False)
        self.label_16.setGeometry(QRect(10, 70, 371, 16))
        self.label_16.setFont(font1)
        self.label_16.setFrameShadow(QFrame.Plain)
        self.label_16.setWordWrap(True)
        self.StartCell_Settings = Gui_PrefLineEdit(self.groupBox_External_Source)
        self.StartCell_Settings.setObjectName(u"StartCell_Settings")
        self.StartCell_Settings.setEnabled(False)
        self.StartCell_Settings.setGeometry(QRect(10, 330, 131, 20))
        self.StartCell_Settings.setFont(font1)
        self.StartCell_Settings.setProperty("prefEntry", u"StartCell_Settings")
        self.StartCell_Settings.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.label_9 = QLabel(self.groupBox_External_Source)
        self.label_9.setObjectName(u"label_9")
        self.label_9.setEnabled(False)
        self.label_9.setGeometry(QRect(145, 325, 216, 31))
        self.label_9.setFont(font2)
        self.label_9.setWordWrap(True)
        self.groupBox_Included_Items = QGroupBox(Form)
        self.groupBox_Included_Items.setObjectName(u"groupBox_Included_Items")
        self.groupBox_Included_Items.setGeometry(QRect(5, 655, 421, 181))
        self.groupBox_Included_Items.setFont(font)
        self.label_4 = QLabel(self.groupBox_Included_Items)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(10, 20, 396, 61))
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_4.sizePolicy().hasHeightForWidth())
        self.label_4.setSizePolicy(sizePolicy)
        self.label_4.setFont(font1)
        self.label_4.setTextFormat(Qt.RichText)
        self.label_4.setScaledContents(False)
        self.label_4.setAlignment(Qt.AlignLeading|Qt.AlignLeft|Qt.AlignTop)
        self.label_4.setWordWrap(True)
        self.layoutWidget = QWidget(self.groupBox_Included_Items)
        self.layoutWidget.setObjectName(u"layoutWidget")
        self.layoutWidget.setGeometry(QRect(10, 80, 144, 88))
        self.verticalLayout = QVBoxLayout(self.layoutWidget)
        self.verticalLayout.setSpacing(6)
        self.verticalLayout.setContentsMargins(5, 5, 5, 5)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.IncludeLengthUnit = Gui_PrefCheckBox(self.layoutWidget)
        self.IncludeLengthUnit.setObjectName(u"IncludeLengthUnit")
        self.IncludeLengthUnit.setFont(font1)
        self.IncludeLengthUnit.setProperty("prefEntry", u"IncludeLength")
        self.IncludeLengthUnit.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.verticalLayout.addWidget(self.IncludeLengthUnit)

        self.IncludeAngleUnit = Gui_PrefCheckBox(self.layoutWidget)
        self.IncludeAngleUnit.setObjectName(u"IncludeAngleUnit")
        self.IncludeAngleUnit.setFont(font1)
        self.IncludeAngleUnit.setProperty("prefEntry", u"IncludeAngle")
        self.IncludeAngleUnit.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.verticalLayout.addWidget(self.IncludeAngleUnit)

        self.IncludeMassUnit = Gui_PrefCheckBox(self.layoutWidget)
        self.IncludeMassUnit.setObjectName(u"IncludeMassUnit")
        self.IncludeMassUnit.setFont(font1)
        self.IncludeMassUnit.setProperty("prefEntry", u"IncludeMass")
        self.IncludeMassUnit.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.verticalLayout.addWidget(self.IncludeMassUnit)

        self.IncludeNoOfSheets = Gui_PrefCheckBox(self.layoutWidget)
        self.IncludeNoOfSheets.setObjectName(u"IncludeNoOfSheets")
        self.IncludeNoOfSheets.setFont(font1)
        self.IncludeNoOfSheets.setProperty("prefEntry", u"IncludeNoOfSheets")
        self.IncludeNoOfSheets.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.verticalLayout.addWidget(self.IncludeNoOfSheets)

        self.label = QLabel(Form)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(15, 840, 401, 16))
        font3 = QFont()
        font3.setItalic(True)
        self.label.setFont(font3)
        self.groupBox_Map_Items = QGroupBox(Form)
        self.groupBox_Map_Items.setObjectName(u"groupBox_Map_Items")
        self.groupBox_Map_Items.setGeometry(QRect(5, 475, 421, 176))
        self.groupBox_Map_Items.setFont(font)
        self.label_3 = QLabel(self.groupBox_Map_Items)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(10, 20, 321, 36))
        self.label_3.setFont(font1)
        self.label_3.setTextFormat(Qt.AutoText)
        self.label_3.setWordWrap(True)
        self.layoutWidget1 = QWidget(self.groupBox_Map_Items)
        self.layoutWidget1.setObjectName(u"layoutWidget1")
        self.layoutWidget1.setGeometry(QRect(10, 65, 401, 100))
        self.gridLayout = QGridLayout(self.layoutWidget1)
        self.gridLayout.setSpacing(6)
        self.gridLayout.setContentsMargins(5, 5, 5, 5)
        self.gridLayout.setObjectName(u"gridLayout")
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.label_5 = QLabel(self.layoutWidget1)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setFont(font1)

        self.gridLayout.addWidget(self.label_5, 0, 0, 1, 1)

        self.MapLength = Gui_PrefLineEdit(self.layoutWidget1)
        self.MapLength.setObjectName(u"MapLength")
        self.MapLength.setFont(font1)
        self.MapLength.setProperty("prefEntry", u"MapLength")
        self.MapLength.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.gridLayout.addWidget(self.MapLength, 0, 1, 1, 1)

        self.label_6 = QLabel(self.layoutWidget1)
        self.label_6.setObjectName(u"label_6")
        self.label_6.setFont(font1)

        self.gridLayout.addWidget(self.label_6, 1, 0, 1, 1)

        self.MapAngle = Gui_PrefLineEdit(self.layoutWidget1)
        self.MapAngle.setObjectName(u"MapAngle")
        self.MapAngle.setFont(font1)
        self.MapAngle.setProperty("prefEntry", u"MapAngle")
        self.MapAngle.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.gridLayout.addWidget(self.MapAngle, 1, 1, 1, 1)

        self.label_7 = QLabel(self.layoutWidget1)
        self.label_7.setObjectName(u"label_7")
        self.label_7.setFont(font1)

        self.gridLayout.addWidget(self.label_7, 2, 0, 1, 1)

        self.MapMass = Gui_PrefLineEdit(self.layoutWidget1)
        self.MapMass.setObjectName(u"MapMass")
        self.MapMass.setFont(font1)
        self.MapMass.setProperty("prefEntry", u"MapMass")
        self.MapMass.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.gridLayout.addWidget(self.MapMass, 2, 1, 1, 1)

        self.MapNoSheets = Gui_PrefLineEdit(self.layoutWidget1)
        self.MapNoSheets.setObjectName(u"MapNoSheets")
        self.MapNoSheets.setFont(font1)
        self.MapNoSheets.setInputMethodHints(Qt.ImhNone)
        self.MapNoSheets.setProperty("prefEntry", u"MapNoSheets")
        self.MapNoSheets.setProperty("prefPath", u"TechDrawTitleBlockUtility")

        self.gridLayout.addWidget(self.MapNoSheets, 3, 1, 1, 1)

        self.label_8 = QLabel(self.layoutWidget1)
        self.label_8.setObjectName(u"label_8")
        self.label_8.setFont(font1)

        self.gridLayout.addWidget(self.label_8, 3, 0, 1, 1)

        self.EnableDebug = Gui_PrefCheckBox(Form)
        self.EnableDebug.setObjectName(u"EnableDebug")
        self.EnableDebug.setGeometry(QRect(15, 865, 91, 17))
        self.EnableDebug.setChecked(True)
        self.EnableDebug.setProperty("prefEntry", u"EnableDebug")
        self.EnableDebug.setProperty("prefPath", u"TechDrawTitleBlockUtility")
        self.label_11 = QLabel(Form)
        self.label_11.setObjectName(u"label_11")
        self.label_11.setGeometry(QRect(15, 890, 326, 16))
        QWidget.setTabOrder(self.UseExternalSource, self.SheetName)
        QWidget.setTabOrder(self.SheetName, self.StartCell)
        QWidget.setTabOrder(self.StartCell, self.AutoFillTitleBlock)
        QWidget.setTabOrder(self.AutoFillTitleBlock, self.ImportSettingsXL)
        QWidget.setTabOrder(self.ImportSettingsXL, self.SheetName_Settings)
        QWidget.setTabOrder(self.SheetName_Settings, self.StartCell_Settings)
        QWidget.setTabOrder(self.StartCell_Settings, self.UseFileName)
        QWidget.setTabOrder(self.UseFileName, self.DrawingNumber)
        QWidget.setTabOrder(self.DrawingNumber, self.MapLength)
        QWidget.setTabOrder(self.MapLength, self.MapAngle)
        QWidget.setTabOrder(self.MapAngle, self.MapMass)
        QWidget.setTabOrder(self.MapMass, self.MapNoSheets)
        QWidget.setTabOrder(self.MapNoSheets, self.IncludeLengthUnit)
        QWidget.setTabOrder(self.IncludeLengthUnit, self.IncludeAngleUnit)
        QWidget.setTabOrder(self.IncludeAngleUnit, self.IncludeMassUnit)
        QWidget.setTabOrder(self.IncludeMassUnit, self.IncludeNoOfSheets)
        QWidget.setTabOrder(self.IncludeNoOfSheets, self.EnableDebug)

        self.retranslateUi(Form)
        self.UseExternalSource.toggled.connect(self.ExternalFileChooser.setEnabled)
        self.UseFileName.toggled.connect(self.label_10.setEnabled)
        self.UseFileName.toggled.connect(self.DrawingNumber.setEnabled)
        self.UseExternalSource.toggled.connect(self.SheetName.setEnabled)
        self.UseExternalSource.toggled.connect(self.label_2.setVisible)
        self.UseExternalSource.toggled.connect(self.StartCell.setEnabled)
        self.UseExternalSource.toggled.connect(self.AutoFillTitleBlock.setEnabled)
        self.UseExternalSource.toggled.connect(self.ImportSettingsXL.setEnabled)
        self.ImportSettingsXL.toggled.connect(self.groupBox_DrawingNumber.setDisabled)
        self.ImportSettingsXL.toggled.connect(self.groupBox_Included_Items.setDisabled)
        self.ImportSettingsXL.toggled.connect(self.groupBox_Map_Items.setDisabled)
        self.ImportSettingsXL.toggled.connect(self.StartCell_Settings.setEnabled)
        self.ImportSettingsXL.toggled.connect(self.label_9.setEnabled)
        self.ImportSettingsXL.toggled.connect(self.SheetName_Settings.setEnabled)
        self.ImportSettingsXL.toggled.connect(self.label_13.setEnabled)
        self.ImportSettingsXL.toggled.connect(self.label_14.setEnabled)
        self.UseExternalSource.toggled.connect(self.label_15.setEnabled)
        self.UseExternalSource.toggled.connect(self.label_16.setEnabled)

        QMetaObject.connectSlotsByName(Form)
    # setupUi

    def retranslateUi(self, Form):
        self.groupBox_DrawingNumber.setTitle(QCoreApplication.translate("Form", u"Drawing number", None))
        self.UseFileName.setText(QCoreApplication.translate("Form", u"Use filename as drawing number", None))
        self.DrawingNumber.setInputMask("")
        self.DrawingNumber.setText("")
        self.DrawingNumber.setPlaceholderText(QCoreApplication.translate("Form", u"Enter the property name...", None))
        self.label_10.setText(QCoreApplication.translate("Form", u"The property name for drawing number:", None))
        self.groupBox_External_Source.setTitle(QCoreApplication.translate("Form", u"External source", None))
        self.ExternalFileChooser.setFileName("")
        self.ExternalFileChooser.setFilter(QCoreApplication.translate("Form", u"*.xlsx", None))
        self.UseExternalSource.setText(QCoreApplication.translate("Form", u"Use external source", None))
        self.StartCell.setInputMask("")
        self.StartCell.setText("")
        self.StartCell.setPlaceholderText(QCoreApplication.translate("Form", u"Enter the start cell...", None))
        self.label_2.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>Cell adress must be like &quot;B1&quot; or like &quot;R1C1&quot;. If empty, A1 will be used.</p></body></html>", None))
        self.AutoFillTitleBlock.setText(QCoreApplication.translate("Form", u"Populate titleblock on import", None))
        self.ImportSettingsXL.setText(QCoreApplication.translate("Form", u"Import workbench settings from external source", None))
#if QT_CONFIG(tooltip)
        self.SheetName_Settings.setToolTip(QCoreApplication.translate("Form", u"Enter the name of sheet to import the settings from.", None))
#endif // QT_CONFIG(tooltip)
#if QT_CONFIG(statustip)
        self.SheetName_Settings.setStatusTip(QCoreApplication.translate("Form", u"Enter the name of sheet to import the settings from.", None))
#endif // QT_CONFIG(statustip)
#if QT_CONFIG(whatsthis)
        self.SheetName_Settings.setWhatsThis(QCoreApplication.translate("Form", u"Enter the name of sheet...", None))
#endif // QT_CONFIG(whatsthis)
        self.SheetName_Settings.setInputMask("")
        self.SheetName_Settings.setText("")
        self.SheetName_Settings.setPlaceholderText(QCoreApplication.translate("Form", u"Enter the name of sheet...", None))
#if QT_CONFIG(tooltip)
        self.SheetName.setToolTip(QCoreApplication.translate("Form", u"Enter the name of the sheet for the titleblock data", None))
#endif // QT_CONFIG(tooltip)
#if QT_CONFIG(whatsthis)
        self.SheetName.setWhatsThis(QCoreApplication.translate("Form", u"Enter the name of the sheet for the titleblock data", None))
#endif // QT_CONFIG(whatsthis)
        self.SheetName.setInputMask("")
        self.SheetName.setText("")
        self.SheetName.setPlaceholderText(QCoreApplication.translate("Form", u"Enter the name of sheet...", None))
        self.label_13.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>The startcell of the table with settings: <span style=\" font-style:italic;\">(This must be the top left cell.)</span></p></body></html>", None))
        self.label_14.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>The name of the worksheet that containts the table with settings:</p></body></html>", None))
        self.label_15.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>The startcell of the table which contains to the data for the titleblock: <span style=\" font-style:italic;\">(This must be the top left cell.)</span></p></body></html>", None))
        self.label_16.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>The name of the worksheet that contains the date for the titleblock:</p></body></html>", None))
#if QT_CONFIG(tooltip)
        self.StartCell_Settings.setToolTip(QCoreApplication.translate("Form", u"Enter the cell to import the settings from.", None))
#endif // QT_CONFIG(tooltip)
#if QT_CONFIG(whatsthis)
        self.StartCell_Settings.setWhatsThis(QCoreApplication.translate("Form", u"Enter the cell to import the settings from.", None))
#endif // QT_CONFIG(whatsthis)
        self.StartCell_Settings.setInputMask("")
        self.StartCell_Settings.setText("")
        self.StartCell_Settings.setPlaceholderText(QCoreApplication.translate("Form", u"Enter the start cell...", None))
        self.label_9.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>Cell adress must be like &quot;B1&quot; or like &quot;R1C1&quot;. If empty, A1 will be used.</p></body></html>", None))
        self.groupBox_Included_Items.setTitle(QCoreApplication.translate("Form", u"Included items", None))
        self.label_4.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>The following properties can be included from the system:</p><p><span style=\" font-style:italic;\">(When the titleblock does have one or more properties already, you will end up with duplicates! It is advised to map the properties instead)</span></p></body></html>", None))
        self.IncludeLengthUnit.setText(QCoreApplication.translate("Form", u"Include lengths", None))
        self.IncludeAngleUnit.setText(QCoreApplication.translate("Form", u"Include angles", None))
        self.IncludeMassUnit.setText(QCoreApplication.translate("Form", u"Include mass", None))
        self.IncludeNoOfSheets.setText(QCoreApplication.translate("Form", u"Include number of pages", None))
        self.label.setText(QCoreApplication.translate("Form", u"FreeCAD needs to be restarted before changes become active.", None))
        self.groupBox_Map_Items.setTitle(QCoreApplication.translate("Form", u"Map items", None))
        self.label_3.setText(QCoreApplication.translate("Form", u"<html><head/><body><p>The following system properties can be mapped to the titleblock: <span style=\" font-style:italic;\">(If not applicalbe, leave empty.)</span></p></body></html>", None))
        self.label_5.setText(QCoreApplication.translate("Form", u"Length unit:", None))
        self.MapLength.setPlaceholderText(QCoreApplication.translate("Form", u"Enter property name here...", None))
        self.label_6.setText(QCoreApplication.translate("Form", u"Angle unit:", None))
        self.MapAngle.setPlaceholderText(QCoreApplication.translate("Form", u"Enter property name here...", None))
        self.label_7.setText(QCoreApplication.translate("Form", u"Mass unit:", None))
        self.MapMass.setPlaceholderText(QCoreApplication.translate("Form", u"Enter property name here...", None))
        self.MapNoSheets.setText("")
        self.MapNoSheets.setPlaceholderText(QCoreApplication.translate("Form", u"Enter property name here...", None))
        self.label_8.setText(QCoreApplication.translate("Form", u"Number of pages:", None))
        self.EnableDebug.setText(QCoreApplication.translate("Form", u"Debug mode*", None))
        self.label_11.setText(QCoreApplication.translate("Form", u"<html><head/><body><p><span style=\" font-style:italic;\">* If enabled, extra information will be shown in the report view.</span></p></body></html>", None))
        pass
    # retranslateUi

