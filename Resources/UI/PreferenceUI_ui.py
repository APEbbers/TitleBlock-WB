# -*- coding: utf-8 -*-

################################################################################
# Form generated from reading UI file 'PreferenceUI.ui'
##
# Created by: Qt User Interface Compiler version 6.6.1
##
# WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide.QtCore import (
    QCoreApplication,
    QDate,
    QDateTime,
    QLocale,
    QMetaObject,
    QObject,
    QPoint,
    QRect,
    QSize,
    QTime,
    QUrl,
    Qt,
)
from PySide.QtGui import (
    QBrush,
    QColor,
    QConicalGradient,
    QCursor,
    QFont,
    QFontDatabase,
    QGradient,
    QIcon,
    QImage,
    QKeySequence,
    QLinearGradient,
    QPainter,
    QPalette,
    QPixmap,
    QRadialGradient,
    QTransform,
)
from PySide.QtWidgets import (
    QAbstractScrollArea,
    QApplication,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLayout,
    QScrollArea,
    QSizePolicy,
    QSpacerItem,
    QTabWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


class Ui_Form(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName("Form")
        Form.setWindowModality(Qt.ApplicationModal)
        Form.setEnabled(True)
        Form.resize(1089, 644)
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form.sizePolicy().hasHeightForWidth())
        Form.setSizePolicy(sizePolicy)
        palette = QPalette()
        brush = QBrush(QColor(255, 255, 255, 255))
        brush.setStyle(Qt.SolidPattern)
        palette.setBrush(QPalette.Active, QPalette.Base, brush)
        brush1 = QBrush(QColor(240, 240, 240, 255))
        brush1.setStyle(Qt.SolidPattern)
        palette.setBrush(QPalette.Active, QPalette.Window, brush1)
        palette.setBrush(QPalette.Inactive, QPalette.Base, brush)
        palette.setBrush(QPalette.Inactive, QPalette.Window, brush1)
        palette.setBrush(QPalette.Disabled, QPalette.Base, brush)
        palette.setBrush(QPalette.Disabled, QPalette.Window, brush1)
        Form.setPalette(palette)
        Form.setFocusPolicy(Qt.StrongFocus)
        Form.setContextMenuPolicy(Qt.NoContextMenu)
        Form.setWindowTitle("Preferences")
        Form.setWindowFilePath("")
        Form.setInputMethodHints(Qt.ImhNone)
        self.widget = QWidget(Form)
        self.widget.setObjectName("widget")
        self.widget.setGeometry(QRect(10, 525, 421, 56))
        self.label = QLabel(self.widget)
        self.label.setObjectName("label")
        self.label.setGeometry(QRect(10, 5, 401, 16))
        font = QFont()
        font.setItalic(True)
        self.label.setFont(font)
        self.layoutWidget_2 = QWidget(self.widget)
        self.layoutWidget_2.setObjectName("layoutWidget_2")
        self.layoutWidget_2.setGeometry(QRect(10, 30, 399, 21))
        self.horizontalLayout = QHBoxLayout(self.layoutWidget_2)
        self.horizontalLayout.setSpacing(3)
        self.horizontalLayout.setContentsMargins(9, 9, 9, 9)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.EnableDebug = Gui_PrefCheckBox(self.layoutWidget_2)
        self.EnableDebug.setObjectName("EnableDebug")
        self.EnableDebug.setChecked(False)
        self.EnableDebug.setProperty("prefEntry", "EnableDebug")
        self.EnableDebug.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.horizontalLayout.addWidget(self.EnableDebug)

        self.label_11 = QLabel(self.layoutWidget_2)
        self.label_11.setObjectName("label_11")

        self.horizontalLayout.addWidget(self.label_11)

        self.tabWidget = QTabWidget(Form)
        self.tabWidget.setObjectName("tabWidget")
        self.tabWidget.setEnabled(True)
        self.tabWidget.setGeometry(QRect(20, 10, 546, 516))
        self.tabWidget.setAutoFillBackground(True)
        self.TitleBlock = QWidget()
        self.TitleBlock.setObjectName("TitleBlock")
        self.scrollArea = QScrollArea(self.TitleBlock)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollArea.setGeometry(QRect(0, 0, 536, 491))
        sizePolicy1 = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        sizePolicy1.setHorizontalStretch(0)
        sizePolicy1.setVerticalStretch(0)
        sizePolicy1.setHeightForWidth(self.scrollArea.sizePolicy().hasHeightForWidth())
        self.scrollArea.setSizePolicy(sizePolicy1)
        self.scrollArea.setFocusPolicy(Qt.NoFocus)
        self.scrollArea.setAutoFillBackground(True)
        self.scrollArea.setFrameShape(QFrame.NoFrame)
        self.scrollArea.setFrameShadow(QFrame.Plain)
        self.scrollArea.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setAlignment(Qt.AlignLeading | Qt.AlignLeft | Qt.AlignTop)
        self.scrollAreaWidgetContents = QWidget()
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.scrollAreaWidgetContents.setGeometry(QRect(0, -491, 519, 1092))
        sizePolicy2 = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
        sizePolicy2.setHorizontalStretch(0)
        sizePolicy2.setVerticalStretch(0)
        sizePolicy2.setHeightForWidth(
            self.scrollAreaWidgetContents.sizePolicy().hasHeightForWidth()
        )
        self.scrollAreaWidgetContents.setSizePolicy(sizePolicy2)
        self.verticalLayout_3 = QVBoxLayout(self.scrollAreaWidgetContents)
        self.verticalLayout_3.setSpacing(3)
        self.verticalLayout_3.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.verticalLayout_3.setSizeConstraint(QLayout.SetNoConstraint)
        self.toolButton = QToolButton(self.scrollAreaWidgetContents)
        self.toolButton.setObjectName("toolButton")
        self.toolButton.setCheckable(True)
        self.toolButton.setPopupMode(QToolButton.InstantPopup)
        self.toolButton.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toolButton.setAutoRaise(False)
        self.toolButton.setArrowType(Qt.DownArrow)

        self.verticalLayout_3.addWidget(self.toolButton)

        self.frame_DrawingNumber_2 = QFrame(self.scrollAreaWidgetContents)
        self.frame_DrawingNumber_2.setObjectName("frame_DrawingNumber_2")
        self.frame_DrawingNumber_2.setEnabled(True)
        self.frame_DrawingNumber_2.setMinimumSize(QSize(0, 101))
        font1 = QFont()
        font1.setBold(True)
        self.frame_DrawingNumber_2.setFont(font1)
        self.frame_DrawingNumber_2.setFrameShape(QFrame.StyledPanel)
        self.frame_DrawingNumber_2.setFrameShadow(QFrame.Sunken)
        self.frame_DrawingNumber_2.setLineWidth(1)
        self.layoutWidget = QWidget(self.frame_DrawingNumber_2)
        self.layoutWidget.setObjectName("layoutWidget")
        self.layoutWidget.setGeometry(QRect(5, 15, 401, 74))
        self.formLayout_17 = QFormLayout(self.layoutWidget)
        self.formLayout_17.setSpacing(3)
        self.formLayout_17.setContentsMargins(9, 9, 9, 9)
        self.formLayout_17.setObjectName("formLayout_17")
        self.formLayout_17.setVerticalSpacing(10)
        self.formLayout_17.setContentsMargins(0, 0, 0, 0)
        self.UsePageName = Gui_PrefCheckBox(self.layoutWidget)
        self.UsePageName.setObjectName("UsePageName")
        font2 = QFont()
        font2.setBold(False)
        self.UsePageName.setFont(font2)
        self.UsePageName.setProperty("prefEntry", "UsePageName")
        self.UsePageName.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_17.setWidget(0, QFormLayout.SpanningRole, self.UsePageName)

        self.formLayout_16 = QFormLayout()
        self.formLayout_16.setSpacing(3)
        self.formLayout_16.setObjectName("formLayout_16")
        self.label_36 = QLabel(self.layoutWidget)
        self.label_36.setObjectName("label_36")
        self.label_36.setEnabled(False)
        self.label_36.setFont(font2)

        self.formLayout_16.setWidget(0, QFormLayout.SpanningRole, self.label_36)

        self.DrawingNumber_Page = Gui_PrefLineEdit(self.layoutWidget)
        self.DrawingNumber_Page.setObjectName("DrawingNumber_Page")
        self.DrawingNumber_Page.setEnabled(False)
        self.DrawingNumber_Page.setFont(font2)
        self.DrawingNumber_Page.setProperty("prefEntry", "DrwNrFieldName_Page")
        self.DrawingNumber_Page.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_16.setWidget(
            1, QFormLayout.SpanningRole, self.DrawingNumber_Page
        )

        self.formLayout_17.setLayout(1, QFormLayout.SpanningRole, self.formLayout_16)

        self.verticalLayout_3.addWidget(self.frame_DrawingNumber_2)

        self.toolButton_2 = QToolButton(self.scrollAreaWidgetContents)
        self.toolButton_2.setObjectName("toolButton_2")
        self.toolButton_2.setCheckable(True)
        self.toolButton_2.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toolButton_2.setArrowType(Qt.DownArrow)

        self.verticalLayout_3.addWidget(self.toolButton_2)

        self.frame_DrawingNumber = QFrame(self.scrollAreaWidgetContents)
        self.frame_DrawingNumber.setObjectName("frame_DrawingNumber")
        self.frame_DrawingNumber.setEnabled(True)
        self.frame_DrawingNumber.setMinimumSize(QSize(0, 101))
        self.frame_DrawingNumber.setFont(font1)
        self.frame_DrawingNumber.setFrameShape(QFrame.StyledPanel)
        self.layoutWidget1 = QWidget(self.frame_DrawingNumber)
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.layoutWidget1.setGeometry(QRect(10, 15, 396, 74))
        self.formLayout_19 = QFormLayout(self.layoutWidget1)
        self.formLayout_19.setSpacing(3)
        self.formLayout_19.setContentsMargins(9, 9, 9, 9)
        self.formLayout_19.setObjectName("formLayout_19")
        self.formLayout_19.setVerticalSpacing(10)
        self.formLayout_19.setContentsMargins(0, 0, 0, 0)
        self.UseFileName = Gui_PrefCheckBox(self.layoutWidget1)
        self.UseFileName.setObjectName("UseFileName")
        self.UseFileName.setFont(font2)
        self.UseFileName.setChecked(False)
        self.UseFileName.setProperty("prefEntry", "UseFileName")
        self.UseFileName.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_19.setWidget(0, QFormLayout.SpanningRole, self.UseFileName)

        self.formLayout_18 = QFormLayout()
        self.formLayout_18.setSpacing(3)
        self.formLayout_18.setObjectName("formLayout_18")
        self.label_10 = QLabel(self.layoutWidget1)
        self.label_10.setObjectName("label_10")
        self.label_10.setEnabled(False)
        self.label_10.setFont(font2)

        self.formLayout_18.setWidget(0, QFormLayout.SpanningRole, self.label_10)

        self.DrawingNumber = Gui_PrefLineEdit(self.layoutWidget1)
        self.DrawingNumber.setObjectName("DrawingNumber")
        self.DrawingNumber.setEnabled(False)
        self.DrawingNumber.setFont(font2)
        self.DrawingNumber.setProperty("prefEntry", "DrwNrFieldName")
        self.DrawingNumber.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_18.setWidget(1, QFormLayout.SpanningRole, self.DrawingNumber)

        self.formLayout_19.setLayout(1, QFormLayout.SpanningRole, self.formLayout_18)

        self.verticalLayout_3.addWidget(self.frame_DrawingNumber)

        self.toolButton_3 = QToolButton(self.scrollAreaWidgetContents)
        self.toolButton_3.setObjectName("toolButton_3")
        self.toolButton_3.setCheckable(True)
        self.toolButton_3.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toolButton_3.setArrowType(Qt.DownArrow)

        self.verticalLayout_3.addWidget(self.toolButton_3)

        self.frame_Map_Items = QFrame(self.scrollAreaWidgetContents)
        self.frame_Map_Items.setObjectName("frame_Map_Items")
        self.frame_Map_Items.setMinimumSize(QSize(0, 186))
        self.frame_Map_Items.setFont(font1)
        self.frame_Map_Items.setFrameShape(QFrame.StyledPanel)
        self.label_3 = QLabel(self.frame_Map_Items)
        self.label_3.setObjectName("label_3")
        self.label_3.setGeometry(QRect(10, 20, 321, 36))
        self.label_3.setFont(font2)
        self.label_3.setTextFormat(Qt.AutoText)
        self.label_3.setWordWrap(True)
        self.layoutWidget2 = QWidget(self.frame_Map_Items)
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.layoutWidget2.setGeometry(QRect(10, 65, 371, 108))
        self.gridLayout = QGridLayout(self.layoutWidget2)
        self.gridLayout.setSpacing(3)
        self.gridLayout.setContentsMargins(9, 9, 9, 9)
        self.gridLayout.setObjectName("gridLayout")
        self.gridLayout.setHorizontalSpacing(10)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.MapNoSheets = Gui_PrefLineEdit(self.layoutWidget2)
        self.MapNoSheets.setObjectName("MapNoSheets")
        self.MapNoSheets.setFont(font2)
        self.MapNoSheets.setInputMethodHints(Qt.ImhNone)
        self.MapNoSheets.setProperty("prefEntry", "MapNoSheets")
        self.MapNoSheets.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout.addWidget(self.MapNoSheets, 3, 1, 1, 1)

        self.MapLength = Gui_PrefLineEdit(self.layoutWidget2)
        self.MapLength.setObjectName("MapLength")
        self.MapLength.setFont(font2)
        self.MapLength.setProperty("prefEntry", "MapLength")
        self.MapLength.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout.addWidget(self.MapLength, 0, 1, 1, 1)

        self.label_5 = QLabel(self.layoutWidget2)
        self.label_5.setObjectName("label_5")
        self.label_5.setFont(font2)

        self.gridLayout.addWidget(self.label_5, 0, 0, 1, 1)

        self.MapAngle = Gui_PrefLineEdit(self.layoutWidget2)
        self.MapAngle.setObjectName("MapAngle")
        self.MapAngle.setFont(font2)
        self.MapAngle.setProperty("prefEntry", "MapAngle")
        self.MapAngle.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout.addWidget(self.MapAngle, 1, 1, 1, 1)

        self.label_8 = QLabel(self.layoutWidget2)
        self.label_8.setObjectName("label_8")
        self.label_8.setFont(font2)

        self.gridLayout.addWidget(self.label_8, 3, 0, 1, 1)

        self.label_6 = QLabel(self.layoutWidget2)
        self.label_6.setObjectName("label_6")
        self.label_6.setFont(font2)

        self.gridLayout.addWidget(self.label_6, 1, 0, 1, 1)

        self.label_7 = QLabel(self.layoutWidget2)
        self.label_7.setObjectName("label_7")
        self.label_7.setFont(font2)

        self.gridLayout.addWidget(self.label_7, 2, 0, 1, 1)

        self.MapMass = Gui_PrefLineEdit(self.layoutWidget2)
        self.MapMass.setObjectName("MapMass")
        self.MapMass.setFont(font2)
        self.MapMass.setProperty("prefEntry", "MapMass")
        self.MapMass.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout.addWidget(self.MapMass, 2, 1, 1, 1)

        self.verticalLayout_3.addWidget(self.frame_Map_Items)

        self.toolButton_4 = QToolButton(self.scrollAreaWidgetContents)
        self.toolButton_4.setObjectName("toolButton_4")
        self.toolButton_4.setCheckable(True)
        self.toolButton_4.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toolButton_4.setArrowType(Qt.DownArrow)

        self.verticalLayout_3.addWidget(self.toolButton_4)

        self.frame_Map_DocInfo = QFrame(self.scrollAreaWidgetContents)
        self.frame_Map_DocInfo.setObjectName("frame_Map_DocInfo")
        self.frame_Map_DocInfo.setMinimumSize(QSize(0, 325))
        self.frame_Map_DocInfo.setFont(font1)
        self.frame_Map_DocInfo.setFrameShape(QFrame.StyledPanel)
        self.label_17 = QLabel(self.frame_Map_DocInfo)
        self.label_17.setObjectName("label_17")
        self.label_17.setGeometry(QRect(10, 20, 321, 36))
        self.label_17.setFont(font2)
        self.label_17.setTextFormat(Qt.AutoText)
        self.label_17.setWordWrap(True)
        self.layoutWidget3 = QWidget(self.frame_Map_DocInfo)
        self.layoutWidget3.setObjectName("layoutWidget3")
        self.layoutWidget3.setGeometry(QRect(11, 66, 371, 248))
        self.gridLayout_2 = QGridLayout(self.layoutWidget3)
        self.gridLayout_2.setSpacing(3)
        self.gridLayout_2.setContentsMargins(9, 9, 9, 9)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.label_18 = QLabel(self.layoutWidget3)
        self.label_18.setObjectName("label_18")
        self.label_18.setFont(font2)

        self.gridLayout_2.addWidget(self.label_18, 3, 0, 1, 1)

        self.label_24 = QLabel(self.layoutWidget3)
        self.label_24.setObjectName("label_24")
        self.label_24.setFont(font2)

        self.gridLayout_2.addWidget(self.label_24, 6, 0, 1, 1)

        self.label_25 = QLabel(self.layoutWidget3)
        self.label_25.setObjectName("label_25")
        self.label_25.setFont(font2)

        self.gridLayout_2.addWidget(self.label_25, 7, 0, 1, 1)

        self.label_26 = QLabel(self.layoutWidget3)
        self.label_26.setObjectName("label_26")
        self.label_26.setFont(font2)

        self.gridLayout_2.addWidget(self.label_26, 8, 0, 1, 1)

        self.DocInfo_Company = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_Company.setObjectName("DocInfo_Company")
        self.DocInfo_Company.setFont(font2)
        self.DocInfo_Company.setInputMethodHints(Qt.ImhNone)
        self.DocInfo_Company.setProperty("prefEntry", "DocInfo_Company")
        self.DocInfo_Company.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_Company, 5, 1, 1, 2)

        self.DocInfo_CreatedBy = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_CreatedBy.setObjectName("DocInfo_CreatedBy")
        self.DocInfo_CreatedBy.setFont(font2)
        self.DocInfo_CreatedBy.setProperty("prefEntry", "DocInfo_CreatedBy")
        self.DocInfo_CreatedBy.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_CreatedBy, 1, 1, 1, 2)

        self.label_23 = QLabel(self.layoutWidget3)
        self.label_23.setObjectName("label_23")
        self.label_23.setFont(font2)

        self.gridLayout_2.addWidget(self.label_23, 5, 0, 1, 1)

        self.DocInfo_Name = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_Name.setObjectName("DocInfo_Name")
        self.DocInfo_Name.setFont(font2)
        self.DocInfo_Name.setProperty("prefEntry", "DocInfo_Name")
        self.DocInfo_Name.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_Name, 0, 1, 1, 2)

        self.label_21 = QLabel(self.layoutWidget3)
        self.label_21.setObjectName("label_21")
        self.label_21.setFont(font2)

        self.gridLayout_2.addWidget(self.label_21, 1, 0, 1, 1)

        self.DocInfo_CreatedDate = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_CreatedDate.setObjectName("DocInfo_CreatedDate")
        self.DocInfo_CreatedDate.setFont(font2)
        self.DocInfo_CreatedDate.setProperty("prefEntry", "DocInfo_CreatedDate")
        self.DocInfo_CreatedDate.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_CreatedDate, 2, 1, 1, 2)

        self.label_22 = QLabel(self.layoutWidget3)
        self.label_22.setObjectName("label_22")
        self.label_22.setFont(font2)

        self.gridLayout_2.addWidget(self.label_22, 4, 0, 1, 1)

        self.label_20 = QLabel(self.layoutWidget3)
        self.label_20.setObjectName("label_20")
        self.label_20.setFont(font2)

        self.gridLayout_2.addWidget(self.label_20, 0, 0, 1, 1)

        self.DocInfo_LastModifiedBy = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_LastModifiedBy.setObjectName("DocInfo_LastModifiedBy")
        self.DocInfo_LastModifiedBy.setFont(font2)
        self.DocInfo_LastModifiedBy.setInputMethodHints(Qt.ImhNone)
        self.DocInfo_LastModifiedBy.setProperty("prefEntry", "DocInfo_LastModifiedBy")
        self.DocInfo_LastModifiedBy.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_LastModifiedBy, 3, 1, 1, 2)

        self.label_19 = QLabel(self.layoutWidget3)
        self.label_19.setObjectName("label_19")
        self.label_19.setFont(font2)

        self.gridLayout_2.addWidget(self.label_19, 2, 0, 1, 1)

        self.DocInfo_LastModifiedDate = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_LastModifiedDate.setObjectName("DocInfo_LastModifiedDate")
        self.DocInfo_LastModifiedDate.setFont(font2)
        self.DocInfo_LastModifiedDate.setInputMethodHints(Qt.ImhNone)
        self.DocInfo_LastModifiedDate.setProperty(
            "prefEntry", "DocInfo_LastModifiedDate"
        )
        self.DocInfo_LastModifiedDate.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_2.addWidget(self.DocInfo_LastModifiedDate, 4, 1, 1, 2)

        self.DocInfo_License = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_License.setObjectName("DocInfo_License")
        self.DocInfo_License.setFont(font2)
        self.DocInfo_License.setInputMethodHints(Qt.ImhNone)
        self.DocInfo_License.setProperty("prefEntry", "DocInfo_License")
        self.DocInfo_License.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_License, 6, 1, 1, 2)

        self.DocInfo_LicenseURL = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_LicenseURL.setObjectName("DocInfo_LicenseURL")
        self.DocInfo_LicenseURL.setFont(font2)
        self.DocInfo_LicenseURL.setInputMethodHints(Qt.ImhNone)
        self.DocInfo_LicenseURL.setProperty("prefEntry", "DocInfo_LicenseURL")
        self.DocInfo_LicenseURL.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_LicenseURL, 7, 1, 1, 2)

        self.DocInfo_Comment = Gui_PrefLineEdit(self.layoutWidget3)
        self.DocInfo_Comment.setObjectName("DocInfo_Comment")
        self.DocInfo_Comment.setFont(font2)
        self.DocInfo_Comment.setInputMethodHints(Qt.ImhNone)
        self.DocInfo_Comment.setProperty("prefEntry", "DocInfo_Comment")
        self.DocInfo_Comment.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_2.addWidget(self.DocInfo_Comment, 8, 1, 1, 2)

        self.verticalLayout_3.addWidget(self.frame_Map_DocInfo)

        self.toolButton_5 = QToolButton(self.scrollAreaWidgetContents)
        self.toolButton_5.setObjectName("toolButton_5")
        self.toolButton_5.setCheckable(True)
        self.toolButton_5.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toolButton_5.setArrowType(Qt.DownArrow)

        self.verticalLayout_3.addWidget(self.toolButton_5)

        self.frame_Included_Items = QFrame(self.scrollAreaWidgetContents)
        self.frame_Included_Items.setObjectName("frame_Included_Items")
        self.frame_Included_Items.setMinimumSize(QSize(0, 181))
        self.frame_Included_Items.setFont(font1)
        self.frame_Included_Items.setFrameShape(QFrame.StyledPanel)
        self.label_4 = QLabel(self.frame_Included_Items)
        self.label_4.setObjectName("label_4")
        self.label_4.setGeometry(QRect(10, 20, 396, 61))
        sizePolicy3 = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        sizePolicy3.setHorizontalStretch(0)
        sizePolicy3.setVerticalStretch(0)
        sizePolicy3.setHeightForWidth(self.label_4.sizePolicy().hasHeightForWidth())
        self.label_4.setSizePolicy(sizePolicy3)
        self.label_4.setFont(font2)
        self.label_4.setTextFormat(Qt.RichText)
        self.label_4.setScaledContents(False)
        self.label_4.setAlignment(Qt.AlignLeading | Qt.AlignLeft | Qt.AlignTop)
        self.label_4.setWordWrap(True)
        self.layoutWidget4 = QWidget(self.frame_Included_Items)
        self.layoutWidget4.setObjectName("layoutWidget4")
        self.layoutWidget4.setGeometry(QRect(10, 80, 146, 96))
        self.verticalLayout = QVBoxLayout(self.layoutWidget4)
        self.verticalLayout.setSpacing(3)
        self.verticalLayout.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout.setObjectName("verticalLayout")
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.IncludeLengthUnit = Gui_PrefCheckBox(self.layoutWidget4)
        self.IncludeLengthUnit.setObjectName("IncludeLengthUnit")
        self.IncludeLengthUnit.setFont(font2)
        self.IncludeLengthUnit.setProperty("prefEntry", "IncludeLength")
        self.IncludeLengthUnit.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.verticalLayout.addWidget(self.IncludeLengthUnit)

        self.IncludeAngleUnit = Gui_PrefCheckBox(self.layoutWidget4)
        self.IncludeAngleUnit.setObjectName("IncludeAngleUnit")
        self.IncludeAngleUnit.setFont(font2)
        self.IncludeAngleUnit.setProperty("prefEntry", "IncludeAngle")
        self.IncludeAngleUnit.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.verticalLayout.addWidget(self.IncludeAngleUnit)

        self.IncludeMassUnit = Gui_PrefCheckBox(self.layoutWidget4)
        self.IncludeMassUnit.setObjectName("IncludeMassUnit")
        self.IncludeMassUnit.setFont(font2)
        self.IncludeMassUnit.setProperty("prefEntry", "IncludeMass")
        self.IncludeMassUnit.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.verticalLayout.addWidget(self.IncludeMassUnit)

        self.IncludeNoOfSheets = Gui_PrefCheckBox(self.layoutWidget4)
        self.IncludeNoOfSheets.setObjectName("IncludeNoOfSheets")
        self.IncludeNoOfSheets.setFont(font2)
        self.IncludeNoOfSheets.setProperty("prefEntry", "IncludeNoOfSheets")
        self.IncludeNoOfSheets.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.verticalLayout.addWidget(self.IncludeNoOfSheets)

        self.verticalLayout_3.addWidget(self.frame_Included_Items)

        self.frame = QFrame(self.scrollAreaWidgetContents)
        self.frame.setObjectName("frame")
        self.frame.setFrameShape(QFrame.NoFrame)
        self.frame.setFrameShadow(QFrame.Raised)

        self.verticalLayout_3.addWidget(self.frame)

        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.tabWidget.addTab(self.TitleBlock, "")
        self.External_Source = QWidget()
        self.External_Source.setObjectName("External_Source")
        self.frame_2 = QFrame(self.External_Source)
        self.frame_2.setObjectName("frame_2")
        self.frame_2.setGeometry(QRect(5, 10, 456, 241))
        sizePolicy4 = QSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        sizePolicy4.setHorizontalStretch(0)
        sizePolicy4.setVerticalStretch(0)
        sizePolicy4.setHeightForWidth(self.frame_2.sizePolicy().hasHeightForWidth())
        self.frame_2.setSizePolicy(sizePolicy4)
        self.frame_2.setFrameShape(QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QFrame.Sunken)
        self.layoutWidget5 = QWidget(self.frame_2)
        self.layoutWidget5.setObjectName("layoutWidget5")
        self.layoutWidget5.setGeometry(QRect(10, 15, 340, 215))
        self.formLayout_13 = QFormLayout(self.layoutWidget5)
        self.formLayout_13.setSpacing(3)
        self.formLayout_13.setContentsMargins(9, 9, 9, 9)
        self.formLayout_13.setObjectName("formLayout_13")
        self.formLayout_13.setVerticalSpacing(10)
        self.formLayout_13.setContentsMargins(0, 0, 0, 0)
        self.formLayout_11 = QFormLayout()
        self.formLayout_11.setSpacing(3)
        self.formLayout_11.setObjectName("formLayout_11")
        self.ExternalFileChooser = Gui_PrefFileChooser(self.layoutWidget5)
        self.ExternalFileChooser.setObjectName("ExternalFileChooser")
        self.ExternalFileChooser.setEnabled(True)
        self.ExternalFileChooser.setFont(font2)
        self.ExternalFileChooser.setProperty("prefEntry", "ExternalFile")
        self.ExternalFileChooser.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_11.setWidget(
            1, QFormLayout.SpanningRole, self.ExternalFileChooser
        )

        self.UseExternalSource = Gui_PrefCheckBox(self.layoutWidget5)
        self.UseExternalSource.setObjectName("UseExternalSource")
        self.UseExternalSource.setFont(font2)
        self.UseExternalSource.setProperty("prefEntry", "UseExternalSource")
        self.UseExternalSource.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_11.setWidget(
            0, QFormLayout.SpanningRole, self.UseExternalSource
        )

        self.formLayout_13.setLayout(0, QFormLayout.SpanningRole, self.formLayout_11)

        self.formLayout_12 = QFormLayout()
        self.formLayout_12.setSpacing(3)
        self.formLayout_12.setObjectName("formLayout_12")
        self.label_16 = QLabel(self.layoutWidget5)
        self.label_16.setObjectName("label_16")
        self.label_16.setEnabled(True)
        self.label_16.setFont(font2)
        self.label_16.setFrameShadow(QFrame.Plain)
        self.label_16.setWordWrap(True)

        self.formLayout_12.setWidget(0, QFormLayout.SpanningRole, self.label_16)

        self.SheetName = Gui_PrefLineEdit(self.layoutWidget5)
        self.SheetName.setObjectName("SheetName")
        self.SheetName.setEnabled(True)
        self.SheetName.setFont(font2)
        self.SheetName.setProperty("prefEntry", "SheetName")
        self.SheetName.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_12.setWidget(1, QFormLayout.SpanningRole, self.SheetName)

        self.formLayout_13.setLayout(1, QFormLayout.SpanningRole, self.formLayout_12)

        self.gridLayout_13 = QGridLayout()
        self.gridLayout_13.setSpacing(3)
        self.gridLayout_13.setObjectName("gridLayout_13")
        self.gridLayout_13.setHorizontalSpacing(6)
        self.gridLayout_13.setVerticalSpacing(0)
        self.label_2 = QLabel(self.layoutWidget5)
        self.label_2.setObjectName("label_2")
        self.label_2.setEnabled(True)
        sizePolicy4.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy4)
        self.label_2.setMinimumSize(QSize(210, 30))
        font3 = QFont()
        font3.setBold(False)
        font3.setItalic(True)
        self.label_2.setFont(font3)
        self.label_2.setWordWrap(True)

        self.gridLayout_13.addWidget(self.label_2, 2, 1, 1, 1)

        self.StartCell = Gui_PrefLineEdit(self.layoutWidget5)
        self.StartCell.setObjectName("StartCell")
        self.StartCell.setEnabled(True)
        self.StartCell.setFont(font2)
        self.StartCell.setProperty("prefEntry", "StartCell")
        self.StartCell.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_13.addWidget(self.StartCell, 2, 0, 1, 1)

        self.label_15 = QLabel(self.layoutWidget5)
        self.label_15.setObjectName("label_15")
        self.label_15.setEnabled(True)
        sizePolicy5 = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        sizePolicy5.setHorizontalStretch(0)
        sizePolicy5.setVerticalStretch(0)
        sizePolicy5.setHeightForWidth(self.label_15.sizePolicy().hasHeightForWidth())
        self.label_15.setSizePolicy(sizePolicy5)
        self.label_15.setMinimumSize(QSize(0, 30))
        self.label_15.setFont(font2)
        self.label_15.setFrameShadow(QFrame.Plain)
        self.label_15.setWordWrap(True)

        self.gridLayout_13.addWidget(self.label_15, 0, 0, 2, 3)

        self.formLayout_13.setLayout(2, QFormLayout.SpanningRole, self.gridLayout_13)

        self.AutoFillTitleBlock = Gui_PrefCheckBox(self.layoutWidget5)
        self.AutoFillTitleBlock.setObjectName("AutoFillTitleBlock")
        self.AutoFillTitleBlock.setEnabled(False)
        self.AutoFillTitleBlock.setFont(font2)
        self.AutoFillTitleBlock.setChecked(False)
        self.AutoFillTitleBlock.setProperty("prefEntry", "AutoFillTitleBlock")
        self.AutoFillTitleBlock.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_13.setWidget(
            3, QFormLayout.SpanningRole, self.AutoFillTitleBlock
        )

        self.frame_3 = QFrame(self.External_Source)
        self.frame_3.setObjectName("frame_3")
        self.frame_3.setGeometry(QRect(5, 265, 456, 176))
        self.frame_3.setFrameShape(QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QFrame.Sunken)
        self.layoutWidget6 = QWidget(self.frame_3)
        self.layoutWidget6.setObjectName("layoutWidget6")
        self.layoutWidget6.setGeometry(QRect(10, 15, 334, 152))
        self.formLayout_15 = QFormLayout(self.layoutWidget6)
        self.formLayout_15.setSpacing(3)
        self.formLayout_15.setContentsMargins(9, 9, 9, 9)
        self.formLayout_15.setObjectName("formLayout_15")
        self.formLayout_15.setVerticalSpacing(10)
        self.formLayout_15.setContentsMargins(0, 0, 0, 0)
        self.ImportSettingsXL = Gui_PrefCheckBox(self.layoutWidget6)
        self.ImportSettingsXL.setObjectName("ImportSettingsXL")
        self.ImportSettingsXL.setEnabled(False)
        self.ImportSettingsXL.setFont(font2)
        self.ImportSettingsXL.setProperty("prefEntry", "ImportSettingsXL")
        self.ImportSettingsXL.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_15.setWidget(0, QFormLayout.SpanningRole, self.ImportSettingsXL)

        self.formLayout_14 = QFormLayout()
        self.formLayout_14.setSpacing(3)
        self.formLayout_14.setObjectName("formLayout_14")
        self.label_14 = QLabel(self.layoutWidget6)
        self.label_14.setObjectName("label_14")
        self.label_14.setEnabled(True)
        self.label_14.setFont(font2)
        self.label_14.setFrameShadow(QFrame.Plain)
        self.label_14.setWordWrap(True)

        self.formLayout_14.setWidget(0, QFormLayout.SpanningRole, self.label_14)

        self.SheetName_Settings = Gui_PrefLineEdit(self.layoutWidget6)
        self.SheetName_Settings.setObjectName("SheetName_Settings")
        self.SheetName_Settings.setEnabled(True)
        self.SheetName_Settings.setFont(font2)
        self.SheetName_Settings.setProperty("prefEntry", "SheetName_Settings")
        self.SheetName_Settings.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_14.setWidget(
            1, QFormLayout.SpanningRole, self.SheetName_Settings
        )

        self.formLayout_15.setLayout(1, QFormLayout.SpanningRole, self.formLayout_14)

        self.gridLayout_14 = QGridLayout()
        self.gridLayout_14.setSpacing(3)
        self.gridLayout_14.setObjectName("gridLayout_14")
        self.StartCell_Settings = Gui_PrefLineEdit(self.layoutWidget6)
        self.StartCell_Settings.setObjectName("StartCell_Settings")
        self.StartCell_Settings.setEnabled(True)
        self.StartCell_Settings.setFont(font2)
        self.StartCell_Settings.setProperty("prefEntry", "StartCell_Settings")
        self.StartCell_Settings.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_14.addWidget(self.StartCell_Settings, 1, 0, 1, 1)

        self.label_9 = QLabel(self.layoutWidget6)
        self.label_9.setObjectName("label_9")
        self.label_9.setEnabled(True)
        sizePolicy4.setHeightForWidth(self.label_9.sizePolicy().hasHeightForWidth())
        self.label_9.setSizePolicy(sizePolicy4)
        self.label_9.setMinimumSize(QSize(210, 30))
        self.label_9.setFont(font3)
        self.label_9.setWordWrap(True)

        self.gridLayout_14.addWidget(self.label_9, 1, 1, 1, 1)

        self.label_13 = QLabel(self.layoutWidget6)
        self.label_13.setObjectName("label_13")
        self.label_13.setEnabled(True)
        sizePolicy5.setHeightForWidth(self.label_13.sizePolicy().hasHeightForWidth())
        self.label_13.setSizePolicy(sizePolicy5)
        self.label_13.setMinimumSize(QSize(0, 30))
        self.label_13.setFont(font2)
        self.label_13.setFrameShadow(QFrame.Plain)
        self.label_13.setAlignment(Qt.AlignLeading | Qt.AlignLeft | Qt.AlignTop)
        self.label_13.setWordWrap(True)

        self.gridLayout_14.addWidget(self.label_13, 0, 0, 1, 2)

        self.formLayout_15.setLayout(2, QFormLayout.SpanningRole, self.gridLayout_14)

        self.tabWidget.addTab(self.External_Source, "")
        self.TechDraw_Workbench = QWidget()
        self.TechDraw_Workbench.setObjectName("TechDraw_Workbench")
        self.frame1 = QFrame(self.TechDraw_Workbench)
        self.frame1.setObjectName("frame1")
        self.frame1.setGeometry(QRect(5, 10, 456, 121))
        sizePolicy4.setHeightForWidth(self.frame1.sizePolicy().hasHeightForWidth())
        self.frame1.setSizePolicy(sizePolicy4)
        self.frame1.setFrameShape(QFrame.StyledPanel)
        self.frame1.setFrameShadow(QFrame.Sunken)
        self.layoutWidget7 = QWidget(self.frame1)
        self.layoutWidget7.setObjectName("layoutWidget7")
        self.layoutWidget7.setGeometry(QRect(10, 15, 229, 95))
        self.formLayout_10 = QFormLayout(self.layoutWidget7)
        self.formLayout_10.setSpacing(3)
        self.formLayout_10.setContentsMargins(9, 9, 9, 9)
        self.formLayout_10.setObjectName("formLayout_10")
        self.formLayout_10.setContentsMargins(0, 0, 0, 0)
        self.AddToolBarTechDraw = Gui_PrefCheckBox(self.layoutWidget7)
        self.AddToolBarTechDraw.setObjectName("AddToolBarTechDraw")
        self.AddToolBarTechDraw.setEnabled(True)
        self.AddToolBarTechDraw.setFont(font2)
        self.AddToolBarTechDraw.setChecked(True)
        self.AddToolBarTechDraw.setProperty("prefEntry", "AddToolBarTechDraw")
        self.AddToolBarTechDraw.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_10.setWidget(0, QFormLayout.LabelRole, self.AddToolBarTechDraw)

        self.Import_templates = Gui_PrefCheckBox(self.layoutWidget7)
        self.Import_templates.setObjectName("Import_templates")
        self.Import_templates.setProperty("prefEntry", "Import_templates")
        self.Import_templates.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_10.setWidget(1, QFormLayout.LabelRole, self.Import_templates)

        self.formLayout_9 = QFormLayout()
        self.formLayout_9.setSpacing(3)
        self.formLayout_9.setObjectName("formLayout_9")
        self.Default_Template = Gui_PrefComboBox(self.layoutWidget7)
        self.Default_Template.addItem("")
        self.Default_Template.addItem("")
        self.Default_Template.addItem("")
        self.Default_Template.addItem("")
        self.Default_Template.addItem("")
        self.Default_Template.addItem("")
        self.Default_Template.setObjectName("Default_Template")
        self.Default_Template.setEnabled(False)
        sizePolicy4.setHeightForWidth(
            self.Default_Template.sizePolicy().hasHeightForWidth()
        )
        self.Default_Template.setSizePolicy(sizePolicy4)
        self.Default_Template.setMinimumSize(QSize(150, 0))
        self.Default_Template.setProperty("prefEntry", "Default_Template")
        self.Default_Template.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_9.setWidget(1, QFormLayout.SpanningRole, self.Default_Template)

        self.label_12 = QLabel(self.layoutWidget7)
        self.label_12.setObjectName("label_12")
        self.label_12.setEnabled(False)

        self.formLayout_9.setWidget(0, QFormLayout.SpanningRole, self.label_12)

        self.formLayout_10.setLayout(2, QFormLayout.LabelRole, self.formLayout_9)

        self.tabWidget.addTab(self.TechDraw_Workbench, "")
        self.UIsettings = QWidget()
        self.UIsettings.setObjectName("UIsettings")
        self.Spreadsheet_Layout = QFrame(self.UIsettings)
        self.Spreadsheet_Layout.setObjectName("Spreadsheet_Layout")
        self.Spreadsheet_Layout.setGeometry(QRect(5, 10, 456, 291))
        self.Spreadsheet_Layout.setFrameShape(QFrame.StyledPanel)
        self.Spreadsheet_Layout.setFrameShadow(QFrame.Sunken)
        self.layoutWidget8 = QWidget(self.Spreadsheet_Layout)
        self.layoutWidget8.setObjectName("layoutWidget8")
        self.layoutWidget8.setGeometry(QRect(17, 27, 301, 248))
        self.formLayout = QFormLayout(self.layoutWidget8)
        self.formLayout.setSpacing(3)
        self.formLayout.setContentsMargins(9, 9, 9, 9)
        self.formLayout.setObjectName("formLayout")
        self.formLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_3 = QGridLayout()
        self.gridLayout_3.setSpacing(3)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.label_27 = QLabel(self.layoutWidget8)
        self.label_27.setObjectName("label_27")

        self.gridLayout_3.addWidget(self.label_27, 0, 0, 1, 1)

        self.SprHeaderBackGround = Gui_PrefColorButton(self.layoutWidget8)
        self.SprHeaderBackGround.setObjectName("SprHeaderBackGround")
        self.SprHeaderBackGround.setAutoDefault(False)
        self.SprHeaderBackGround.setColor(QColor(243, 202, 98))
        self.SprHeaderBackGround.setAllowTransparency(False)
        self.SprHeaderBackGround.setProperty("prefEntry", "SpreadSheetHeaderBackGround")
        self.SprHeaderBackGround.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_3.addWidget(self.SprHeaderBackGround, 0, 1, 1, 1)

        self.label_28 = QLabel(self.layoutWidget8)
        self.label_28.setObjectName("label_28")

        self.gridLayout_3.addWidget(self.label_28, 1, 0, 1, 1)

        self.SprHeaderForeGround = Gui_PrefColorButton(self.layoutWidget8)
        self.SprHeaderForeGround.setObjectName("SprHeaderForeGround")
        self.SprHeaderForeGround.setCheckable(False)
        self.SprHeaderForeGround.setChecked(False)
        self.SprHeaderForeGround.setAutoDefault(False)
        self.SprHeaderForeGround.setFlat(False)
        self.SprHeaderForeGround.setColor(QColor(0, 0, 0))
        self.SprHeaderForeGround.setAllowTransparency(False)
        self.SprHeaderForeGround.setProperty("prefEntry", "SpreadSheetHeaderForeGround")
        self.SprHeaderForeGround.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_3.addWidget(self.SprHeaderForeGround, 1, 1, 1, 1)

        self.formLayout.setLayout(0, QFormLayout.LabelRole, self.gridLayout_3)

        self.gridLayout_4 = QGridLayout()
        self.gridLayout_4.setSpacing(3)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.label_29 = QLabel(self.layoutWidget8)
        self.label_29.setObjectName("label_29")
        sizePolicy6 = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        sizePolicy6.setHorizontalStretch(40)
        sizePolicy6.setVerticalStretch(0)
        sizePolicy6.setHeightForWidth(self.label_29.sizePolicy().hasHeightForWidth())
        self.label_29.setSizePolicy(sizePolicy6)

        self.gridLayout_4.addWidget(self.label_29, 0, 0, 1, 1)

        self.SprHeaderFontStyle_Bold = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprHeaderFontStyle_Bold.setObjectName("SprHeaderFontStyle_Bold")
        self.SprHeaderFontStyle_Bold.setFont(font1)
        self.SprHeaderFontStyle_Bold.setChecked(True)
        self.SprHeaderFontStyle_Bold.setProperty(
            "prefEntry", "SpreadsheetHeaderFontStyle_Bold"
        )
        self.SprHeaderFontStyle_Bold.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_4.addWidget(self.SprHeaderFontStyle_Bold, 0, 1, 1, 1)

        self.SprHeaderFontStyle_Italic = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprHeaderFontStyle_Italic.setObjectName("SprHeaderFontStyle_Italic")
        self.SprHeaderFontStyle_Italic.setFont(font)
        self.SprHeaderFontStyle_Italic.setProperty(
            "prefEntry", "SpreadsheetHeaderFontStyle_Italic"
        )
        self.SprHeaderFontStyle_Italic.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_4.addWidget(self.SprHeaderFontStyle_Italic, 0, 2, 1, 1)

        self.SprHeaderFontStyle_Underline = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprHeaderFontStyle_Underline.setObjectName("SprHeaderFontStyle_Underline")
        font4 = QFont()
        font4.setUnderline(True)
        self.SprHeaderFontStyle_Underline.setFont(font4)
        self.SprHeaderFontStyle_Underline.setChecked(True)
        self.SprHeaderFontStyle_Underline.setProperty(
            "prefEntry", "SpreadsheetHeaderFontStyle_Underline"
        )
        self.SprHeaderFontStyle_Underline.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_4.addWidget(self.SprHeaderFontStyle_Underline, 0, 3, 1, 1)

        self.formLayout.setLayout(1, QFormLayout.SpanningRole, self.gridLayout_4)

        self.gridLayout_5 = QGridLayout()
        self.gridLayout_5.setSpacing(3)
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.label_31 = QLabel(self.layoutWidget8)
        self.label_31.setObjectName("label_31")

        self.gridLayout_5.addWidget(self.label_31, 2, 0, 1, 1)

        self.SprTableBackGround_1 = Gui_PrefColorButton(self.layoutWidget8)
        self.SprTableBackGround_1.setObjectName("SprTableBackGround_1")
        self.SprTableBackGround_1.setColor(QColor(169, 169, 169))
        self.SprTableBackGround_1.setProperty(
            "prefEntry", "SpreadSheetTableBackGround_1"
        )
        self.SprTableBackGround_1.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_5.addWidget(self.SprTableBackGround_1, 0, 1, 1, 1)

        self.label_30 = QLabel(self.layoutWidget8)
        self.label_30.setObjectName("label_30")

        self.gridLayout_5.addWidget(self.label_30, 0, 0, 1, 1)

        self.SprTableForeGround = Gui_PrefColorButton(self.layoutWidget8)
        self.SprTableForeGround.setObjectName("SprTableForeGround")
        self.SprTableForeGround.setColor(QColor(0, 0, 0))
        self.SprTableForeGround.setProperty("prefEntry", "SpreadSheetTableForeGround")
        self.SprTableForeGround.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_5.addWidget(self.SprTableForeGround, 2, 1, 1, 1)

        self.label_33 = QLabel(self.layoutWidget8)
        self.label_33.setObjectName("label_33")

        self.gridLayout_5.addWidget(self.label_33, 1, 0, 1, 1)

        self.SprTableBackGround_2 = Gui_PrefColorButton(self.layoutWidget8)
        self.SprTableBackGround_2.setObjectName("SprTableBackGround_2")
        self.SprTableBackGround_2.setColor(QColor(128, 128, 128))
        self.SprTableBackGround_2.setProperty(
            "prefEntry", "SpreadSheetTableBackGround_2"
        )
        self.SprTableBackGround_2.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_5.addWidget(self.SprTableBackGround_2, 1, 1, 1, 1)

        self.formLayout.setLayout(2, QFormLayout.LabelRole, self.gridLayout_5)

        self.gridLayout_6 = QGridLayout()
        self.gridLayout_6.setSpacing(3)
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.label_32 = QLabel(self.layoutWidget8)
        self.label_32.setObjectName("label_32")
        sizePolicy6.setHeightForWidth(self.label_32.sizePolicy().hasHeightForWidth())
        self.label_32.setSizePolicy(sizePolicy6)

        self.gridLayout_6.addWidget(self.label_32, 0, 0, 1, 1)

        self.SprTableFontStyle_Bold = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprTableFontStyle_Bold.setObjectName("SprTableFontStyle_Bold")
        self.SprTableFontStyle_Bold.setFont(font1)
        self.SprTableFontStyle_Bold.setProperty(
            "prefEntry", "SpreadsheetTableFontStyle_Bold"
        )
        self.SprTableFontStyle_Bold.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_6.addWidget(self.SprTableFontStyle_Bold, 0, 1, 1, 1)

        self.SprTableFontStyle_Italic = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprTableFontStyle_Italic.setObjectName("SprTableFontStyle_Italic")
        self.SprTableFontStyle_Italic.setFont(font)
        self.SprTableFontStyle_Italic.setProperty(
            "prefEntry", "SpreadsheetTableFontStyle_Italic"
        )
        self.SprTableFontStyle_Italic.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_6.addWidget(self.SprTableFontStyle_Italic, 0, 2, 1, 1)

        self.SprTableFontStyle_Underline = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprTableFontStyle_Underline.setObjectName("SprTableFontStyle_Underline")
        self.SprTableFontStyle_Underline.setFont(font4)
        self.SprTableFontStyle_Underline.setProperty(
            "prefEntry", "SpreadsheetTableFontStyle_Underline"
        )
        self.SprTableFontStyle_Underline.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_6.addWidget(self.SprTableFontStyle_Underline, 0, 3, 1, 1)

        self.formLayout.setLayout(3, QFormLayout.SpanningRole, self.gridLayout_6)

        self.gridLayout_7 = QGridLayout()
        self.gridLayout_7.setSpacing(3)
        self.gridLayout_7.setObjectName("gridLayout_7")
        self.gridLayout_7.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.gridLayout_7.setHorizontalSpacing(6)
        self.label_34 = QLabel(self.layoutWidget8)
        self.label_34.setObjectName("label_34")
        sizePolicy6.setHeightForWidth(self.label_34.sizePolicy().hasHeightForWidth())
        self.label_34.setSizePolicy(sizePolicy6)

        self.gridLayout_7.addWidget(self.label_34, 0, 0, 1, 1)

        self.SprColumnFontStyle_Bold = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprColumnFontStyle_Bold.setObjectName("SprColumnFontStyle_Bold")
        self.SprColumnFontStyle_Bold.setFont(font1)
        self.SprColumnFontStyle_Bold.setChecked(True)
        self.SprColumnFontStyle_Bold.setProperty(
            "prefEntry", "SpreadsheetColumnFontStyle_Bold"
        )
        self.SprColumnFontStyle_Bold.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_7.addWidget(self.SprColumnFontStyle_Bold, 0, 1, 1, 1)

        self.SprColumnFontStyle_Italic = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprColumnFontStyle_Italic.setObjectName("SprColumnFontStyle_Italic")
        self.SprColumnFontStyle_Italic.setFont(font)
        self.SprColumnFontStyle_Italic.setProperty(
            "prefEntry", "SpreadsheetColumnFontStyle_Italic"
        )
        self.SprColumnFontStyle_Italic.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_7.addWidget(self.SprColumnFontStyle_Italic, 0, 2, 1, 1)

        self.SprColumnFontStyle_Underline = Gui_PrefCheckBox(self.layoutWidget8)
        self.SprColumnFontStyle_Underline.setObjectName("SprColumnFontStyle_Underline")
        self.SprColumnFontStyle_Underline.setFont(font4)
        self.SprColumnFontStyle_Underline.setProperty(
            "prefEntry", "SpreadsheetColumnFontStyle_Underline"
        )
        self.SprColumnFontStyle_Underline.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_7.addWidget(self.SprColumnFontStyle_Underline, 0, 3, 1, 1)

        self.formLayout.setLayout(4, QFormLayout.SpanningRole, self.gridLayout_7)

        self.formLayout_2 = QFormLayout()
        self.formLayout_2.setSpacing(3)
        self.formLayout_2.setObjectName("formLayout_2")
        self.label_35 = QLabel(self.layoutWidget8)
        self.label_35.setObjectName("label_35")

        self.formLayout_2.setWidget(0, QFormLayout.LabelRole, self.label_35)

        self.AutoFitFactor = Gui_PrefDoubleSpinBox(self.layoutWidget8)
        self.AutoFitFactor.setObjectName("AutoFitFactor")
        self.AutoFitFactor.setSingleStep(0.500000000000000)
        self.AutoFitFactor.setValue(7.500000000000000)
        self.AutoFitFactor.setProperty("prefEntry", "AutoFitFactor")
        self.AutoFitFactor.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_2.setWidget(0, QFormLayout.FieldRole, self.AutoFitFactor)

        self.formLayout.setLayout(5, QFormLayout.LabelRole, self.formLayout_2)

        self.tabWidget.addTab(self.UIsettings, "")
        self.tab = QWidget()
        self.tab.setObjectName("tab")
        sizePolicy7 = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.MinimumExpanding)
        sizePolicy7.setHorizontalStretch(0)
        sizePolicy7.setVerticalStretch(0)
        sizePolicy7.setHeightForWidth(self.tab.sizePolicy().hasHeightForWidth())
        self.tab.setSizePolicy(sizePolicy7)
        self.scrollArea_2 = QScrollArea(self.tab)
        self.scrollArea_2.setObjectName("scrollArea_2")
        self.scrollArea_2.setGeometry(QRect(5, 0, 531, 486))
        sizePolicy.setHeightForWidth(self.scrollArea_2.sizePolicy().hasHeightForWidth())
        self.scrollArea_2.setSizePolicy(sizePolicy)
        self.scrollArea_2.setAutoFillBackground(False)
        self.scrollArea_2.setFrameShape(QFrame.NoFrame)
        self.scrollArea_2.setFrameShadow(QFrame.Plain)
        self.scrollArea_2.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scrollArea_2.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scrollArea_2.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        self.scrollArea_2.setWidgetResizable(False)
        self.scrollArea_2.setAlignment(Qt.AlignLeading | Qt.AlignLeft | Qt.AlignTop)
        self.scrollAreaWidgetContents_2 = QWidget()
        self.scrollAreaWidgetContents_2.setObjectName("scrollAreaWidgetContents_2")
        self.scrollAreaWidgetContents_2.setGeometry(QRect(0, -218, 445, 1000))
        sizePolicy2.setHeightForWidth(
            self.scrollAreaWidgetContents_2.sizePolicy().hasHeightForWidth()
        )
        self.scrollAreaWidgetContents_2.setSizePolicy(sizePolicy2)
        self.verticalLayout_2 = QVBoxLayout(self.scrollAreaWidgetContents_2)
        self.verticalLayout_2.setSpacing(9)
        self.verticalLayout_2.setContentsMargins(9, 9, 9, 9)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout_2.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.verticalLayout_2.setContentsMargins(14, 10, 9, 9)
        self.toolButton_6 = QToolButton(self.scrollAreaWidgetContents_2)
        self.toolButton_6.setObjectName("toolButton_6")
        self.toolButton_6.setCheckable(True)
        self.toolButton_6.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toolButton_6.setArrowType(Qt.DownArrow)

        self.verticalLayout_2.addWidget(self.toolButton_6)

        self.frame_31 = QFrame(self.scrollAreaWidgetContents_2)
        self.frame_31.setObjectName("frame_31")
        sizePolicy4.setHeightForWidth(self.frame_31.sizePolicy().hasHeightForWidth())
        self.frame_31.setSizePolicy(sizePolicy4)
        self.frame_31.setMinimumSize(QSize(419, 300))
        self.frame_31.setMaximumSize(QSize(419, 375))
        self.frame_31.setAutoFillBackground(True)
        self.frame_31.setFrameShape(QFrame.StyledPanel)
        self.frame_31.setFrameShadow(QFrame.Sunken)
        self.frame_31.setLineWidth(0)
        self.UseSimpleList = Gui_PrefCheckBox(self.frame_31)
        self.UseSimpleList.setObjectName("UseSimpleList")
        self.UseSimpleList.setGeometry(QRect(10, 10, 399, 19))
        self.UseSimpleList.setFont(font2)
        self.UseSimpleList.setProperty("prefEntry", "UseSimpleList")
        self.UseSimpleList.setProperty("prefPath", "Mod/TitleBlock Workbench")
        self.frame_5 = QFrame(self.frame_31)
        self.frame_5.setObjectName("frame_5")
        self.frame_5.setEnabled(False)
        self.frame_5.setGeometry(QRect(-1, 35, 421, 266))
        self.frame_5.setFrameShape(QFrame.NoFrame)
        self.frame_5.setFrameShadow(QFrame.Raised)
        self.layoutWidget9 = QWidget(self.frame_5)
        self.layoutWidget9.setObjectName("layoutWidget9")
        self.layoutWidget9.setGeometry(QRect(10, 10, 401, 251))
        self.formLayout_5 = QFormLayout(self.layoutWidget9)
        self.formLayout_5.setSpacing(3)
        self.formLayout_5.setContentsMargins(9, 9, 9, 9)
        self.formLayout_5.setObjectName("formLayout_5")
        self.formLayout_5.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.formLayout_5.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        self.formLayout_5.setRowWrapPolicy(QFormLayout.DontWrapRows)
        self.formLayout_5.setLabelAlignment(
            Qt.AlignLeading | Qt.AlignLeft | Qt.AlignTop
        )
        self.formLayout_5.setFormAlignment(Qt.AlignLeading | Qt.AlignLeft | Qt.AlignTop)
        self.formLayout_5.setVerticalSpacing(10)
        self.formLayout_5.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_8 = QGridLayout()
        self.gridLayout_8.setSpacing(3)
        self.gridLayout_8.setObjectName("gridLayout_8")
        self.ExternalFileChooser_SimpleList = Gui_PrefFileChooser(self.layoutWidget9)
        self.ExternalFileChooser_SimpleList.setObjectName(
            "ExternalFileChooser_SimpleList"
        )
        self.ExternalFileChooser_SimpleList.setEnabled(False)
        self.ExternalFileChooser_SimpleList.setMinimumSize(QSize(0, 20))
        self.ExternalFileChooser_SimpleList.setFont(font2)
        self.ExternalFileChooser_SimpleList.setProperty(
            "prefEntry", "ExternalFile_SimpleList"
        )
        self.ExternalFileChooser_SimpleList.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_8.addWidget(self.ExternalFileChooser_SimpleList, 3, 0, 1, 1)

        self.UseExternalSource_SimpleList = Gui_PrefCheckBox(self.layoutWidget9)
        self.UseExternalSource_SimpleList.setObjectName("UseExternalSource_SimpleList")
        self.UseExternalSource_SimpleList.setProperty(
            "prefEntry", "UseExternalSource_SimpleList"
        )
        self.UseExternalSource_SimpleList.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.gridLayout_8.addWidget(self.UseExternalSource_SimpleList, 2, 0, 1, 1)

        self.formLayout_5.setLayout(0, QFormLayout.SpanningRole, self.gridLayout_8)

        self.gridLayout_9 = QGridLayout()
        self.gridLayout_9.setSpacing(3)
        self.gridLayout_9.setObjectName("gridLayout_9")
        self.label_37 = QLabel(self.layoutWidget9)
        self.label_37.setObjectName("label_37")
        self.label_37.setEnabled(False)
        self.label_37.setFont(font2)
        self.label_37.setFrameShadow(QFrame.Plain)
        self.label_37.setWordWrap(True)

        self.gridLayout_9.addWidget(self.label_37, 0, 0, 1, 1)

        self.SheetName_SimpleList = Gui_PrefLineEdit(self.layoutWidget9)
        self.SheetName_SimpleList.setObjectName("SheetName_SimpleList")
        self.SheetName_SimpleList.setEnabled(False)
        self.SheetName_SimpleList.setMinimumSize(QSize(0, 20))
        self.SheetName_SimpleList.setFont(font2)
        self.SheetName_SimpleList.setProperty("prefEntry", "SheetName_SimpleList")
        self.SheetName_SimpleList.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_9.addWidget(self.SheetName_SimpleList, 1, 0, 1, 1)

        self.formLayout_5.setLayout(1, QFormLayout.SpanningRole, self.gridLayout_9)

        self.gridLayout_10 = QGridLayout()
        self.gridLayout_10.setSpacing(3)
        self.gridLayout_10.setObjectName("gridLayout_10")
        self.label_39 = QLabel(self.layoutWidget9)
        self.label_39.setObjectName("label_39")
        self.label_39.setEnabled(False)
        sizePolicy4.setHeightForWidth(self.label_39.sizePolicy().hasHeightForWidth())
        self.label_39.setSizePolicy(sizePolicy4)
        self.label_39.setMinimumSize(QSize(215, 0))
        self.label_39.setFont(font3)
        self.label_39.setWordWrap(True)

        self.gridLayout_10.addWidget(self.label_39, 2, 1, 1, 1)

        self.StartCell_SimpleList = Gui_PrefLineEdit(self.layoutWidget9)
        self.StartCell_SimpleList.setObjectName("StartCell_SimpleList")
        self.StartCell_SimpleList.setEnabled(False)
        self.StartCell_SimpleList.setMinimumSize(QSize(0, 20))
        self.StartCell_SimpleList.setFont(font2)
        self.StartCell_SimpleList.setProperty("prefEntry", "StartCell_SimpleList")
        self.StartCell_SimpleList.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_10.addWidget(self.StartCell_SimpleList, 2, 0, 1, 1)

        self.label_38 = QLabel(self.layoutWidget9)
        self.label_38.setObjectName("label_38")
        self.label_38.setEnabled(False)
        sizePolicy5.setHeightForWidth(self.label_38.sizePolicy().hasHeightForWidth())
        self.label_38.setSizePolicy(sizePolicy5)
        self.label_38.setMinimumSize(QSize(0, 5))
        self.label_38.setMaximumSize(QSize(16777215, 30))
        self.label_38.setFont(font2)
        self.label_38.setFrameShadow(QFrame.Plain)
        self.label_38.setWordWrap(True)

        self.gridLayout_10.addWidget(self.label_38, 1, 0, 1, 2)

        self.formLayout_5.setLayout(2, QFormLayout.SpanningRole, self.gridLayout_10)

        self.gridLayout_11 = QGridLayout()
        self.gridLayout_11.setSpacing(3)
        self.gridLayout_11.setObjectName("gridLayout_11")
        self.PropertyName_SimpleList = Gui_PrefLineEdit(self.layoutWidget9)
        self.PropertyName_SimpleList.setObjectName("PropertyName_SimpleList")
        self.PropertyName_SimpleList.setMinimumSize(QSize(0, 20))
        self.PropertyName_SimpleList.setProperty("prefEntry", "PropertyName_SimpleList")
        self.PropertyName_SimpleList.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_11.addWidget(self.PropertyName_SimpleList, 2, 0, 1, 1)

        self.label_40 = QLabel(self.layoutWidget9)
        self.label_40.setObjectName("label_40")

        self.gridLayout_11.addWidget(self.label_40, 1, 0, 1, 1)

        self.formLayout_5.setLayout(3, QFormLayout.SpanningRole, self.gridLayout_11)

        self.UsePageNames_SimpleList = Gui_PrefCheckBox(self.layoutWidget9)
        self.UsePageNames_SimpleList.setObjectName("UsePageNames_SimpleList")
        self.UsePageNames_SimpleList.setProperty("prefEntry", "UsePageNames_SimpleList")
        self.UsePageNames_SimpleList.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_5.setWidget(
            4, QFormLayout.SpanningRole, self.UsePageNames_SimpleList
        )

        self.verticalLayout_2.addWidget(self.frame_31)

        self.toolButton_7 = QToolButton(self.scrollAreaWidgetContents_2)
        self.toolButton_7.setObjectName("toolButton_7")
        self.toolButton_7.setCheckable(True)
        self.toolButton_7.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toolButton_7.setArrowType(Qt.DownArrow)

        self.verticalLayout_2.addWidget(self.toolButton_7)

        self.frame_4 = QFrame(self.scrollAreaWidgetContents_2)
        self.frame_4.setObjectName("frame_4")
        sizePolicy.setHeightForWidth(self.frame_4.sizePolicy().hasHeightForWidth())
        self.frame_4.setSizePolicy(sizePolicy)
        self.frame_4.setMinimumSize(QSize(280, 360))
        self.frame_4.setMaximumSize(QSize(419, 455))
        self.frame_4.setAutoFillBackground(True)
        self.frame_4.setFrameShape(QFrame.StyledPanel)
        self.frame_4.setFrameShadow(QFrame.Sunken)
        self.frame_4.setLineWidth(1)
        self.frame_21 = QFrame(self.frame_4)
        self.frame_21.setObjectName("frame_21")
        self.frame_21.setEnabled(False)
        self.frame_21.setGeometry(QRect(0, 29, 421, 331))
        self.frame_21.setFrameShape(QFrame.NoFrame)
        self.frame_21.setFrameShadow(QFrame.Plain)
        self.layoutWidget10 = QWidget(self.frame_21)
        self.layoutWidget10.setObjectName("layoutWidget10")
        self.layoutWidget10.setGeometry(QRect(10, 10, 396, 316))
        self.formLayout_8 = QFormLayout(self.layoutWidget10)
        self.formLayout_8.setSpacing(3)
        self.formLayout_8.setContentsMargins(9, 9, 9, 9)
        self.formLayout_8.setObjectName("formLayout_8")
        self.formLayout_8.setVerticalSpacing(10)
        self.formLayout_8.setContentsMargins(0, 0, 0, 0)
        self.formLayout_7 = QFormLayout()
        self.formLayout_7.setSpacing(3)
        self.formLayout_7.setObjectName("formLayout_7")
        self.formLayout_7.setSizeConstraint(QLayout.SetDefaultConstraint)
        self.formLayout_7.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        self.UseExternalSource_AdvancedList = Gui_PrefCheckBox(self.layoutWidget10)
        self.UseExternalSource_AdvancedList.setObjectName(
            "UseExternalSource_AdvancedList"
        )
        self.UseExternalSource_AdvancedList.setProperty(
            "prefEntry", "UseExternalSource_AdvancedList"
        )
        self.UseExternalSource_AdvancedList.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.formLayout_7.setWidget(
            0, QFormLayout.SpanningRole, self.UseExternalSource_AdvancedList
        )

        self.ExternalFileChooser_AdvancedList = Gui_PrefFileChooser(self.layoutWidget10)
        self.ExternalFileChooser_AdvancedList.setObjectName(
            "ExternalFileChooser_AdvancedList"
        )
        self.ExternalFileChooser_AdvancedList.setEnabled(False)
        self.ExternalFileChooser_AdvancedList.setFont(font2)
        self.ExternalFileChooser_AdvancedList.setProperty(
            "prefEntry", "ExternalFile_AdvancedList"
        )
        self.ExternalFileChooser_AdvancedList.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.formLayout_7.setWidget(
            1, QFormLayout.SpanningRole, self.ExternalFileChooser_AdvancedList
        )

        self.formLayout_8.setLayout(0, QFormLayout.SpanningRole, self.formLayout_7)

        self.formLayout_6 = QFormLayout()
        self.formLayout_6.setSpacing(3)
        self.formLayout_6.setObjectName("formLayout_6")
        self.formLayout_6.setSizeConstraint(QLayout.SetNoConstraint)
        self.formLayout_6.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self.label_41 = QLabel(self.layoutWidget10)
        self.label_41.setObjectName("label_41")
        self.label_41.setEnabled(False)
        self.label_41.setFont(font2)
        self.label_41.setFrameShadow(QFrame.Plain)
        self.label_41.setWordWrap(True)

        self.formLayout_6.setWidget(0, QFormLayout.SpanningRole, self.label_41)

        self.SheetName_AdvancedList = Gui_PrefLineEdit(self.layoutWidget10)
        self.SheetName_AdvancedList.setObjectName("SheetName_AdvancedList")
        self.SheetName_AdvancedList.setEnabled(False)
        self.SheetName_AdvancedList.setMinimumSize(QSize(0, 20))
        self.SheetName_AdvancedList.setFont(font2)
        self.SheetName_AdvancedList.setProperty("prefEntry", "SheetName_AdvancedList")
        self.SheetName_AdvancedList.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.formLayout_6.setWidget(
            1, QFormLayout.SpanningRole, self.SheetName_AdvancedList
        )

        self.formLayout_8.setLayout(1, QFormLayout.SpanningRole, self.formLayout_6)

        self.gridLayout_12 = QGridLayout()
        self.gridLayout_12.setSpacing(3)
        self.gridLayout_12.setObjectName("gridLayout_12")
        self.StartCell_AdvancedList = Gui_PrefLineEdit(self.layoutWidget10)
        self.StartCell_AdvancedList.setObjectName("StartCell_AdvancedList")
        self.StartCell_AdvancedList.setEnabled(False)
        self.StartCell_AdvancedList.setFont(font2)
        self.StartCell_AdvancedList.setProperty("prefEntry", "StartCell_AdvancedList")
        self.StartCell_AdvancedList.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.gridLayout_12.addWidget(self.StartCell_AdvancedList, 1, 0, 1, 1)

        self.label_42 = QLabel(self.layoutWidget10)
        self.label_42.setObjectName("label_42")
        self.label_42.setEnabled(False)
        sizePolicy4.setHeightForWidth(self.label_42.sizePolicy().hasHeightForWidth())
        self.label_42.setSizePolicy(sizePolicy4)
        self.label_42.setMinimumSize(QSize(215, 0))
        self.label_42.setFont(font3)
        self.label_42.setWordWrap(True)

        self.gridLayout_12.addWidget(self.label_42, 1, 1, 1, 1)

        self.label_43 = QLabel(self.layoutWidget10)
        self.label_43.setObjectName("label_43")
        self.label_43.setEnabled(False)
        sizePolicy5.setHeightForWidth(self.label_43.sizePolicy().hasHeightForWidth())
        self.label_43.setSizePolicy(sizePolicy5)
        self.label_43.setMinimumSize(QSize(0, 5))
        self.label_43.setMaximumSize(QSize(16777215, 30))
        self.label_43.setFont(font2)
        self.label_43.setFrameShadow(QFrame.Plain)
        self.label_43.setWordWrap(True)

        self.gridLayout_12.addWidget(self.label_43, 0, 0, 1, 2)

        self.formLayout_8.setLayout(2, QFormLayout.SpanningRole, self.gridLayout_12)

        self.formLayout_3 = QFormLayout()
        self.formLayout_3.setSpacing(3)
        self.formLayout_3.setObjectName("formLayout_3")
        self.label_45 = QLabel(self.layoutWidget10)
        self.label_45.setObjectName("label_45")
        self.label_45.setWordWrap(True)

        self.formLayout_3.setWidget(0, QFormLayout.SpanningRole, self.label_45)

        self.SortingPrefix_AdvancedList = Gui_PrefLineEdit(self.layoutWidget10)
        self.SortingPrefix_AdvancedList.setObjectName("SortingPrefix_AdvancedList")
        self.SortingPrefix_AdvancedList.setMinimumSize(QSize(0, 20))
        self.SortingPrefix_AdvancedList.setBaseSize(QSize(0, 20))
        self.SortingPrefix_AdvancedList.setProperty(
            "prefEntry", "SortingPrefix_AdvancedList"
        )
        self.SortingPrefix_AdvancedList.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.formLayout_3.setWidget(
            1, QFormLayout.SpanningRole, self.SortingPrefix_AdvancedList
        )

        self.formLayout_8.setLayout(4, QFormLayout.SpanningRole, self.formLayout_3)

        self.UsePageNames_AdvancesList = Gui_PrefCheckBox(self.layoutWidget10)
        self.UsePageNames_AdvancesList.setObjectName("UsePageNames_AdvancesList")
        self.UsePageNames_AdvancesList.setProperty(
            "prefEntry", "UsePageNames_AdvancesList"
        )
        self.UsePageNames_AdvancesList.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.formLayout_8.setWidget(
            5, QFormLayout.SpanningRole, self.UsePageNames_AdvancesList
        )

        self.formLayout_4 = QFormLayout()
        self.formLayout_4.setSpacing(3)
        self.formLayout_4.setObjectName("formLayout_4")
        self.label_44 = QLabel(self.layoutWidget10)
        self.label_44.setObjectName("label_44")

        self.formLayout_4.setWidget(0, QFormLayout.SpanningRole, self.label_44)

        self.PropertyName_AdvancedList = Gui_PrefLineEdit(self.layoutWidget10)
        self.PropertyName_AdvancedList.setObjectName("PropertyName_AdvancedList")
        self.PropertyName_AdvancedList.setMinimumSize(QSize(0, 20))
        self.PropertyName_AdvancedList.setProperty(
            "prefEntry", "PropertyName_AdvancedList"
        )
        self.PropertyName_AdvancedList.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )

        self.formLayout_4.setWidget(
            1, QFormLayout.SpanningRole, self.PropertyName_AdvancedList
        )

        self.formLayout_8.setLayout(3, QFormLayout.SpanningRole, self.formLayout_4)

        self.UseAdvancedList = Gui_PrefCheckBox(self.frame_4)
        self.UseAdvancedList.setObjectName("UseAdvancedList")
        self.UseAdvancedList.setGeometry(QRect(10, 5, 394, 19))
        self.UseAdvancedList.setFont(font2)
        self.UseAdvancedList.setChecked(False)
        self.UseAdvancedList.setProperty("prefEntry", "UseAdvancedList")
        self.UseAdvancedList.setProperty("prefPath", "Mod/TitleBlock Workbench")

        self.verticalLayout_2.addWidget(self.frame_4)

        self.verticalSpacer = QSpacerItem(
            20, 10, QSizePolicy.Minimum, QSizePolicy.Expanding
        )

        self.verticalLayout_2.addItem(self.verticalSpacer)

        self.scrollArea_2.setWidget(self.scrollAreaWidgetContents_2)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QWidget()
        self.tab_2.setObjectName("tab_2")
        self.frame_6 = QFrame(self.tab_2)
        self.frame_6.setObjectName("frame_6")
        self.frame_6.setGeometry(QRect(10, 170, 456, 241))
        sizePolicy4.setHeightForWidth(self.frame_6.sizePolicy().hasHeightForWidth())
        self.frame_6.setSizePolicy(sizePolicy4)
        self.frame_6.setFrameShape(QFrame.StyledPanel)
        self.frame_6.setFrameShadow(QFrame.Sunken)
        self.EnableRecompute_FillTitleBlock = Gui_PrefCheckBox(self.frame_6)
        self.EnableRecompute_FillTitleBlock.setObjectName(
            "EnableRecompute_FillTitleBlock"
        )
        self.EnableRecompute_FillTitleBlock.setEnabled(False)
        self.EnableRecompute_FillTitleBlock.setGeometry(QRect(5, 5, 436, 17))
        self.EnableRecompute_FillTitleBlock.setProperty(
            "prefEntry", "EnableRecompute_FillSpreadsheet"
        )
        self.EnableRecompute_FillTitleBlock.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )
        self.EnableRecompute_FillSpreadsheet = Gui_PrefCheckBox(self.frame_6)
        self.EnableRecompute_FillSpreadsheet.setObjectName(
            "EnableRecompute_FillSpreadsheet"
        )
        self.EnableRecompute_FillSpreadsheet.setGeometry(QRect(5, 30, 436, 17))
        self.EnableRecompute_FillSpreadsheet.setProperty(
            "prefEntry", "EnableRecompute_FillTitleBlock"
        )
        self.EnableRecompute_FillSpreadsheet.setProperty(
            "prefPath", "Mod/TitleBlock Workbench"
        )
        self.label_46 = QLabel(self.tab_2)
        self.label_46.setObjectName("label_46")
        self.label_46.setGeometry(QRect(10, 5, 491, 141))
        self.tabWidget.addTab(self.tab_2, "")
        self.urlLabel = Gui_UrlLabel(Form)
        self.urlLabel.setObjectName("urlLabel")
        self.urlLabel.setGeometry(QRect(20, 585, 321, 26))
        self.urlLabel.setWordWrap(True)
        self.urlLabel.setTextInteractionFlags(
            Qt.LinksAccessibleByMouse | Qt.TextSelectableByMouse
        )
        QWidget.setTabOrder(self.tabWidget, self.UseFileName)
        QWidget.setTabOrder(self.UseFileName, self.DrawingNumber)
        QWidget.setTabOrder(self.DrawingNumber, self.MapLength)
        QWidget.setTabOrder(self.MapLength, self.MapAngle)
        QWidget.setTabOrder(self.MapAngle, self.MapMass)
        QWidget.setTabOrder(self.MapMass, self.MapNoSheets)
        QWidget.setTabOrder(self.MapNoSheets, self.DocInfo_Name)
        QWidget.setTabOrder(self.DocInfo_Name, self.DocInfo_CreatedBy)
        QWidget.setTabOrder(self.DocInfo_CreatedBy, self.DocInfo_CreatedDate)
        QWidget.setTabOrder(self.DocInfo_CreatedDate, self.DocInfo_LastModifiedBy)
        QWidget.setTabOrder(self.DocInfo_LastModifiedBy, self.DocInfo_LastModifiedDate)
        QWidget.setTabOrder(self.DocInfo_LastModifiedDate, self.DocInfo_Company)
        QWidget.setTabOrder(self.DocInfo_Company, self.IncludeLengthUnit)
        QWidget.setTabOrder(self.IncludeLengthUnit, self.IncludeAngleUnit)
        QWidget.setTabOrder(self.IncludeAngleUnit, self.IncludeMassUnit)
        QWidget.setTabOrder(self.IncludeMassUnit, self.IncludeNoOfSheets)
        QWidget.setTabOrder(self.IncludeNoOfSheets, self.UseExternalSource)
        QWidget.setTabOrder(self.UseExternalSource, self.SheetName)
        QWidget.setTabOrder(self.SheetName, self.StartCell)
        QWidget.setTabOrder(self.StartCell, self.AutoFillTitleBlock)
        QWidget.setTabOrder(self.AutoFillTitleBlock, self.ImportSettingsXL)
        QWidget.setTabOrder(self.ImportSettingsXL, self.SheetName_Settings)
        QWidget.setTabOrder(self.SheetName_Settings, self.StartCell_Settings)
        QWidget.setTabOrder(self.StartCell_Settings, self.AddToolBarTechDraw)
        QWidget.setTabOrder(self.AddToolBarTechDraw, self.Import_templates)
        QWidget.setTabOrder(self.Import_templates, self.Default_Template)

        self.retranslateUi(Form)
        self.toolButton_7.clicked["bool"].connect(self.frame_4.setHidden)
        self.toolButton_5.clicked["bool"].connect(self.frame_Included_Items.setHidden)
        self.toolButton_4.clicked["bool"].connect(self.frame_Map_DocInfo.setHidden)
        self.toolButton_3.clicked["bool"].connect(self.frame_Map_Items.setHidden)
        self.toolButton_2.clicked["bool"].connect(self.frame_DrawingNumber.setHidden)
        self.toolButton.clicked["bool"].connect(self.frame_DrawingNumber_2.setHidden)
        self.toolButton_6.clicked["bool"].connect(self.frame_3.setHidden)
        self.UseAdvancedList.toggled.connect(self.frame_2.setEnabled)
        self.UseSimpleList.toggled.connect(self.frame_5.setEnabled)
        self.UsePageName.toggled.connect(self.DrawingNumber_Page.setEnabled)
        self.UseFileName.toggled.connect(self.label_10.setEnabled)
        self.UseFileName.toggled.connect(self.DrawingNumber.setEnabled)
        self.UseSimpleList.toggled.connect(self.label_38.setEnabled)
        self.UseSimpleList.toggled.connect(self.StartCell_SimpleList.setEnabled)
        self.UseSimpleList.toggled.connect(self.label_39.setEnabled)
        self.UseExternalSource_SimpleList.toggled.connect(
            self.ExternalFileChooser_SimpleList.setEnabled
        )
        self.UseSimpleList.toggled.connect(self.label_37.setEnabled)
        self.UseSimpleList.toggled.connect(self.SheetName_SimpleList.setEnabled)
        self.toolButton_6.toggled.connect(self.label_37.setEnabled)
        self.UsePageNames_SimpleList.toggled.connect(self.label_40.setDisabled)
        self.UsePageNames_SimpleList.toggled.connect(
            self.PropertyName_SimpleList.setDisabled
        )
        self.UseAdvancedList.toggled.connect(self.label_41.setEnabled)
        self.UseAdvancedList.toggled.connect(self.SheetName_AdvancedList.setEnabled)
        self.UseAdvancedList.toggled.connect(self.label_43.setEnabled)
        self.UseAdvancedList.toggled.connect(self.StartCell_AdvancedList.setEnabled)
        self.UseAdvancedList.toggled.connect(self.label_42.setEnabled)
        self.UseExternalSource_AdvancedList.toggled.connect(
            self.ExternalFileChooser_AdvancedList.setEnabled
        )
        self.UsePageNames_AdvancesList.toggled.connect(self.label_44.setDisabled)
        self.UsePageNames_AdvancesList.toggled.connect(
            self.PropertyName_AdvancedList.setDisabled
        )

        self.tabWidget.setCurrentIndex(5)
        self.SprHeaderBackGround.setDefault(True)
        self.SprHeaderForeGround.setDefault(False)

    # setupUi

    def retranslateUi(self, Form):
        self.label.setText(
            QCoreApplication.translate(
                "Form",
                "FreeCAD needs to be restarted before changes become active.",
                None,
            )
        )
        self.EnableDebug.setText(QCoreApplication.translate("Form", "Debug mode", None))
        self.label_11.setText(
            QCoreApplication.translate(
                "Form",
                '<html><head/><body><p><span style=" font-style:italic;">(If enabled, extra information will be shown in the report view.)</span></p></body></html>',
                None,
            )
        )
        self.toolButton.setText(QCoreApplication.translate("Form", "Page name", None))
        self.UsePageName.setText(
            QCoreApplication.translate("Form", "Use page name as drawing number", None)
        )
        self.label_36.setText(
            QCoreApplication.translate(
                "Form", "The property name for drawing number:", None
            )
        )
        self.DrawingNumber_Page.setInputMask("")
        self.DrawingNumber_Page.setText("")
        self.DrawingNumber_Page.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the property name...", None)
        )
        self.toolButton_2.setText(QCoreApplication.translate("Form", "Filename", None))
        self.UseFileName.setText(
            QCoreApplication.translate("Form", "Use filename as drawing number", None)
        )
        self.label_10.setText(
            QCoreApplication.translate(
                "Form", "The property name for drawing number:", None
            )
        )
        self.DrawingNumber.setInputMask("")
        self.DrawingNumber.setText("")
        self.DrawingNumber.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the property name...", None)
        )
        self.toolButton_3.setText(
            QCoreApplication.translate("Form", "Map system properties", None)
        )
        self.label_3.setText(
            QCoreApplication.translate(
                "Form",
                '<html><head/><body><p>The following system properties can be mapped to the titleblock: <span style=" font-style:italic;">(If not applicalbe, leave empty.)</span></p></body></html>',
                None,
            )
        )
        self.MapNoSheets.setText("")
        self.MapNoSheets.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.MapLength.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.label_5.setText(QCoreApplication.translate("Form", "Length unit:", None))
        self.MapAngle.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.label_8.setText(
            QCoreApplication.translate("Form", "Number of pages:            ", None)
        )
        self.label_6.setText(QCoreApplication.translate("Form", "Angle unit:", None))
        self.label_7.setText(QCoreApplication.translate("Form", "Mass unit:", None))
        self.MapMass.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.toolButton_4.setText(
            QCoreApplication.translate("Form", "Map document information", None)
        )
        self.label_17.setText(
            QCoreApplication.translate(
                "Form",
                '<html><head/><body><p>The following document information values can be mapped to the titleblock: <span style=" font-style:italic;">(If not applicalbe, leave empty.)</span></p></body></html>',
                None,
            )
        )
        self.label_18.setText(
            QCoreApplication.translate("Form", "Last modified by:", None)
        )
        self.label_24.setText(
            QCoreApplication.translate("Form", "License information:", None)
        )
        self.label_25.setText(QCoreApplication.translate("Form", "License URL:", None))
        self.label_26.setText(QCoreApplication.translate("Form", "Comment:", None))
        self.DocInfo_Company.setText("")
        self.DocInfo_Company.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.DocInfo_CreatedBy.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.label_23.setText(QCoreApplication.translate("Form", "Company:", None))
        self.DocInfo_Name.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.label_21.setText(QCoreApplication.translate("Form", "Created by:", None))
        self.DocInfo_CreatedDate.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.label_22.setText(
            QCoreApplication.translate("Form", "Last modification date:", None)
        )
        self.label_20.setText(QCoreApplication.translate("Form", "Name:", None))
        self.DocInfo_LastModifiedBy.setText("")
        self.DocInfo_LastModifiedBy.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.label_19.setText(
            QCoreApplication.translate("Form", "Creation date:", None)
        )
        self.DocInfo_LastModifiedDate.setText("")
        self.DocInfo_LastModifiedDate.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.DocInfo_License.setText("")
        self.DocInfo_License.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.DocInfo_LicenseURL.setText("")
        self.DocInfo_LicenseURL.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.DocInfo_Comment.setText("")
        self.DocInfo_Comment.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter property name here...", None)
        )
        self.toolButton_5.setText(
            QCoreApplication.translate("Form", "Include system properties", None)
        )
        self.label_4.setText(
            QCoreApplication.translate(
                "Form",
                '<html><head/><body><p>The following properties can be included from the system:</p><p><span style=" font-style:italic;">(When the titleblock does have one or more properties already, you will end up with duplicates! It is advised to map the properties instead)</span></p></body></html>',
                None,
            )
        )
        self.IncludeLengthUnit.setText(
            QCoreApplication.translate("Form", "Include lengths", None)
        )
        self.IncludeAngleUnit.setText(
            QCoreApplication.translate("Form", "Include angles", None)
        )
        self.IncludeMassUnit.setText(
            QCoreApplication.translate("Form", "Include mass", None)
        )
        self.IncludeNoOfSheets.setText(
            QCoreApplication.translate("Form", "Include number of pages", None)
        )
        self.tabWidget.setTabText(
            self.tabWidget.indexOf(self.TitleBlock),
            QCoreApplication.translate("Form", "TitleBlock", None),
        )
        self.ExternalFileChooser.setFileName("")
        self.ExternalFileChooser.setFilter(
            QCoreApplication.translate("Form", "*.xlsx; *.FCStd", None)
        )
        self.UseExternalSource.setText(
            QCoreApplication.translate("Form", "Use external source", None)
        )
        self.label_16.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>The name of the worksheet that contains the date for the titleblock:</p></body></html>",
                None,
            )
        )
        # if QT_CONFIG(tooltip)
        self.SheetName.setToolTip(
            QCoreApplication.translate(
                "Form", "Enter the name of the sheet for the titleblock data", None
            )
        )
        # endif // QT_CONFIG(tooltip)
        # if QT_CONFIG(whatsthis)
        self.SheetName.setWhatsThis(
            QCoreApplication.translate(
                "Form", "Enter the name of the sheet for the titleblock data", None
            )
        )
        # endif // QT_CONFIG(whatsthis)
        self.SheetName.setInputMask("")
        self.SheetName.setText("")
        self.SheetName.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the name of sheet...", None)
        )
        self.label_2.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>Cell adress must be like &quot;B1&quot; or like &quot;R1C1&quot;. If empty, A1 will be used.</p></body></html>",
                None,
            )
        )
        self.StartCell.setInputMask("")
        self.StartCell.setText("")
        self.StartCell.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the start cell...", None)
        )
        self.label_15.setText(
            QCoreApplication.translate(
                "Form",
                '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">\n'
                '<html><head><meta name="qrichtext" content="1" /><style type="text/css">\n'
                "p, li { white-space: pre-wrap; }\n"
                "</style></head><body style=\" font-family:'MS Shell Dlg 2'; font-size:8pt; font-weight:400; font-style:normal;\">\n"
                '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">The startcell of the table which contains to the data for the titleblock: 	<span style=" font-style:italic;">(This must be the top left cell.)</span></p></body></html>',
                None,
            )
        )
        self.AutoFillTitleBlock.setText(
            QCoreApplication.translate("Form", "Populate titleblock on import", None)
        )
        self.ImportSettingsXL.setText(
            QCoreApplication.translate(
                "Form", "Import workbench settings from external source", None
            )
        )
        self.label_14.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>The name of the worksheet that containts the table with settings:</p></body></html>",
                None,
            )
        )
        # if QT_CONFIG(tooltip)
        self.SheetName_Settings.setToolTip(
            QCoreApplication.translate(
                "Form", "Enter the name of sheet to import the settings from.", None
            )
        )
        # endif // QT_CONFIG(tooltip)
        # if QT_CONFIG(statustip)
        self.SheetName_Settings.setStatusTip(
            QCoreApplication.translate(
                "Form", "Enter the name of sheet to import the settings from.", None
            )
        )
        # endif // QT_CONFIG(statustip)
        # if QT_CONFIG(whatsthis)
        self.SheetName_Settings.setWhatsThis(
            QCoreApplication.translate("Form", "Enter the name of sheet...", None)
        )
        # endif // QT_CONFIG(whatsthis)
        self.SheetName_Settings.setInputMask("")
        self.SheetName_Settings.setText("")
        self.SheetName_Settings.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the name of sheet...", None)
        )
        # if QT_CONFIG(tooltip)
        self.StartCell_Settings.setToolTip(
            QCoreApplication.translate(
                "Form", "Enter the cell to import the settings from.", None
            )
        )
        # endif // QT_CONFIG(tooltip)
        # if QT_CONFIG(whatsthis)
        self.StartCell_Settings.setWhatsThis(
            QCoreApplication.translate(
                "Form", "Enter the cell to import the settings from.", None
            )
        )
        # endif // QT_CONFIG(whatsthis)
        self.StartCell_Settings.setInputMask("")
        self.StartCell_Settings.setText("")
        self.StartCell_Settings.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the start cell...", None)
        )
        self.label_9.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>Cell adress must be like &quot;B1&quot; or like &quot;R1C1&quot;. If empty, A1 will be used.</p></body></html>",
                None,
            )
        )
        self.label_13.setText(
            QCoreApplication.translate(
                "Form",
                '<html><head/><body><p>The startcell of the table with settings: <span style=" font-style:italic;">(This must be the top left cell.)</span></p></body></html>',
                None,
            )
        )
        self.tabWidget.setTabText(
            self.tabWidget.indexOf(self.External_Source),
            QCoreApplication.translate("Form", "External Source", None),
        )
        self.AddToolBarTechDraw.setText(
            QCoreApplication.translate(
                "Form", "Add toolbar to the TechDraw Workbench", None
            )
        )
        self.Import_templates.setText(
            QCoreApplication.translate("Form", "Import example templates", None)
        )
        self.Default_Template.setItemText(
            0, QCoreApplication.translate("Form", "A0_Landscape", None)
        )
        self.Default_Template.setItemText(
            1, QCoreApplication.translate("Form", "A1_Landscape", None)
        )
        self.Default_Template.setItemText(
            2, QCoreApplication.translate("Form", "A2_Landscape", None)
        )
        self.Default_Template.setItemText(
            3, QCoreApplication.translate("Form", "A3_Landscape", None)
        )
        self.Default_Template.setItemText(
            4, QCoreApplication.translate("Form", "A4_Landscape", None)
        )
        self.Default_Template.setItemText(
            5, QCoreApplication.translate("Form", "A4_Portrait", None)
        )

        self.label_12.setText(
            QCoreApplication.translate("Form", "Set default template", None)
        )
        self.tabWidget.setTabText(
            self.tabWidget.indexOf(self.TechDraw_Workbench),
            QCoreApplication.translate("Form", "TechDraw Workbench", None),
        )
        self.label_27.setText(
            QCoreApplication.translate("Form", "Header background       ", None)
        )
        self.label_28.setText(
            QCoreApplication.translate("Form", "Header foreground", None)
        )
        self.label_29.setText(
            QCoreApplication.translate("Form", "Header font style", None)
        )
        self.SprHeaderFontStyle_Bold.setText(
            QCoreApplication.translate("Form", "Bold", None)
        )
        self.SprHeaderFontStyle_Italic.setText(
            QCoreApplication.translate("Form", "Italic", None)
        )
        self.SprHeaderFontStyle_Underline.setText(
            QCoreApplication.translate("Form", "Underline", None)
        )
        self.label_31.setText(
            QCoreApplication.translate("Form", "Table foreground", None)
        )
        self.label_30.setText(
            QCoreApplication.translate("Form", "Table background 1       ", None)
        )
        self.label_33.setText(
            QCoreApplication.translate("Form", "Table background 2", None)
        )
        self.label_32.setText(
            QCoreApplication.translate("Form", "Table font style", None)
        )
        self.SprTableFontStyle_Bold.setText(
            QCoreApplication.translate("Form", "Bold", None)
        )
        self.SprTableFontStyle_Italic.setText(
            QCoreApplication.translate("Form", "Italic", None)
        )
        self.SprTableFontStyle_Underline.setText(
            QCoreApplication.translate("Form", "Underline", None)
        )
        self.label_34.setText(
            QCoreApplication.translate("Form", "1st column font style", None)
        )
        self.SprColumnFontStyle_Bold.setText(
            QCoreApplication.translate("Form", "Bold", None)
        )
        self.SprColumnFontStyle_Italic.setText(
            QCoreApplication.translate("Form", "Italic", None)
        )
        self.SprColumnFontStyle_Underline.setText(
            QCoreApplication.translate("Form", "Underline", None)
        )
        self.label_35.setText(
            QCoreApplication.translate("Form", "Width factor for AutoFit", None)
        )
        self.tabWidget.setTabText(
            self.tabWidget.indexOf(self.UIsettings),
            QCoreApplication.translate("Form", "UI Settings", None),
        )
        self.toolButton_6.setText(
            QCoreApplication.translate("Form", "Simple drawing list", None)
        )
        self.UseSimpleList.setText(
            QCoreApplication.translate("Form", "Use simple drawing list", None)
        )
        self.ExternalFileChooser_SimpleList.setFileName("")
        self.ExternalFileChooser_SimpleList.setFilter(
            QCoreApplication.translate("Form", "*.xlsx; *.FCStd", None)
        )
        self.UseExternalSource_SimpleList.setText(
            QCoreApplication.translate("Form", "Use external source", None)
        )
        self.label_37.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>The name of the worksheet that contains the drawing list:</p></body></html>",
                None,
            )
        )
        # if QT_CONFIG(tooltip)
        self.SheetName_SimpleList.setToolTip(
            QCoreApplication.translate(
                "Form", "Enter the name of the sheet for the titleblock data", None
            )
        )
        # endif // QT_CONFIG(tooltip)
        # if QT_CONFIG(whatsthis)
        self.SheetName_SimpleList.setWhatsThis(
            QCoreApplication.translate(
                "Form", "Enter the name of the sheet for the titleblock data", None
            )
        )
        # endif // QT_CONFIG(whatsthis)
        self.SheetName_SimpleList.setInputMask("")
        self.SheetName_SimpleList.setText("")
        self.SheetName_SimpleList.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the name of sheet...", None)
        )
        self.label_39.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>Cell adress must be like &quot;B1&quot; or like &quot;R1C1&quot;. If empty, A1 will be used.</p></body></html>",
                None,
            )
        )
        self.StartCell_SimpleList.setInputMask("")
        self.StartCell_SimpleList.setText("")
        self.StartCell_SimpleList.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the start cell...", None)
        )
        self.label_38.setText(
            QCoreApplication.translate(
                "Form",
                '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">\n'
                '<html><head><meta name="qrichtext" content="1" /><style type="text/css">\n'
                "p, li { white-space: pre-wrap; }\n"
                "</style></head><body style=\" font-family:'MS Shell Dlg 2'; font-size:8pt; font-weight:400; font-style:normal;\">\n"
                '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">The startcell of the table which contains the drawing list:	     	<span style=" font-style:italic;">(This must be the top left cell.)</span></p></body></html>',
                None,
            )
        )
        self.PropertyName_SimpleList.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the property name...", None)
        )
        self.label_40.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>The property value to look for in the drawing list:</p></body></html>",
                None,
            )
        )
        self.UsePageNames_SimpleList.setText(
            QCoreApplication.translate(
                "Form", "Use page names instead of property names", None
            )
        )
        self.toolButton_7.setText(
            QCoreApplication.translate("Form", "Advanced drawing list", None)
        )
        self.UseExternalSource_AdvancedList.setText(
            QCoreApplication.translate("Form", "Use external source", None)
        )
        self.ExternalFileChooser_AdvancedList.setFileName("")
        self.ExternalFileChooser_AdvancedList.setFilter(
            QCoreApplication.translate("Form", "*.xlsx; *.FCStd", None)
        )
        self.label_41.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>The name of the worksheet that contains the drawing list:</p></body></html>",
                None,
            )
        )
        # if QT_CONFIG(tooltip)
        self.SheetName_AdvancedList.setToolTip(
            QCoreApplication.translate(
                "Form", "Enter the name of the sheet for the titleblock data", None
            )
        )
        # endif // QT_CONFIG(tooltip)
        # if QT_CONFIG(whatsthis)
        self.SheetName_AdvancedList.setWhatsThis(
            QCoreApplication.translate(
                "Form", "Enter the name of the sheet for the titleblock data", None
            )
        )
        # endif // QT_CONFIG(whatsthis)
        self.SheetName_AdvancedList.setInputMask("")
        self.SheetName_AdvancedList.setText("")
        self.SheetName_AdvancedList.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the name of sheet...", None)
        )
        self.StartCell_AdvancedList.setInputMask("")
        self.StartCell_AdvancedList.setText("")
        self.StartCell_AdvancedList.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the start cell...", None)
        )
        self.label_42.setText(
            QCoreApplication.translate(
                "Form",
                "<html><head/><body><p>Cell adress must be like &quot;B1&quot; or like &quot;R1C1&quot;. If empty, A1 will be used.</p></body></html>",
                None,
            )
        )
        self.label_43.setText(
            QCoreApplication.translate(
                "Form",
                '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">\n'
                '<html><head><meta name="qrichtext" content="1" /><style type="text/css">\n'
                "p, li { white-space: pre-wrap; }\n"
                "</style></head><body style=\" font-family:'MS Shell Dlg 2'; font-size:8pt; font-weight:400; font-style:normal;\">\n"
                '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">The startcell of the table which contains the drawing list:     		<span style=" font-style:italic;">(This must be the top left cell.)</span></p></body></html>',
                None,
            )
        )
        self.label_45.setText(
            QCoreApplication.translate(
                "Form",
                '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">\n'
                '<html><head><meta name="qrichtext" content="1" /><style type="text/css">\n'
                "p, li { white-space: pre-wrap; }\n"
                "</style></head><body style=\" font-family:'MS Shell Dlg 2'; font-size:8pt; font-weight:400; font-style:normal;\">\n"
                '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">If an prefix is used to sort the pages, enter it here. It will be ignored when filling in the titleblock. For example: &quot;01_&quot;</p></body></html>',
                None,
            )
        )
        self.SortingPrefix_AdvancedList.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the prefix...", None)
        )
        self.UsePageNames_AdvancesList.setText(
            QCoreApplication.translate(
                "Form", "Use page names instead of property names", None
            )
        )
        self.label_44.setText(
            QCoreApplication.translate(
                "Form", "The property name to look for in the external source:", None
            )
        )
        self.PropertyName_AdvancedList.setPlaceholderText(
            QCoreApplication.translate("Form", "Enter the property name...", None)
        )
        self.UseAdvancedList.setText(
            QCoreApplication.translate("Form", "Use advanced drawing list", None)
        )
        self.tabWidget.setTabText(
            self.tabWidget.indexOf(self.tab),
            QCoreApplication.translate("Form", "Drawing list", None),
        )
        self.EnableRecompute_FillTitleBlock.setText(
            QCoreApplication.translate(
                "Form", "update titleblock spreadsheet with recompute", None
            )
        )
        self.EnableRecompute_FillSpreadsheet.setText(
            QCoreApplication.translate(
                "Form", "Populate titleblock with recompute", None
            )
        )
        self.label_46.setText(
            QCoreApplication.translate(
                "Form",
                '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">\n'
                '<html><head><meta name="qrichtext" content="1" /><style type="text/css">\n'
                "p, li { white-space: pre-wrap; }\n"
                "</style></head><body style=\" font-family:'MS Shell Dlg 2'; font-size:8pt; font-weight:400; font-style:normal;\">\n"
                '<p style=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" font-size:10pt; font-weight:600; color:#ff0000;">Experimental features!</span></p>\n'
                '<p style=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">Here you can set if the spreadsheet and/or titleblock must be updated when events happen.</span></p>\n'
                '<p style=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">current supported events:</span></p>\n'
                "<p style"
                '=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">- Recompute documents</span></p>\n'
                '<p style=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">- Recompute objects</span></p>\n'
                '<p style="-qt-paragraph-type:empty; margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; color:#000000;"><br /></p>\n'
                '<p style=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">Currently these functions are working when the one of the following workbenches are activated:</span></p>\n'
                '<p style=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">- TechDraw workbench</span></p>\n'
                '<p style=" margin-top:1px;'
                ' margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">- Spreadsheet workbench</span></p>\n'
                '<p style=" margin-top:1px; margin-bottom:1px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;"><span style=" color:#000000;">- TitleBlock workbench</span></p></body></html>',
                None,
            )
        )
        self.tabWidget.setTabText(
            self.tabWidget.indexOf(self.tab_2),
            QCoreApplication.translate("Form", "Events", None),
        )
        self.urlLabel.setText(
            QCoreApplication.translate(
                "Form",
                '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">\n'
                '<html><head><meta name="qrichtext" content="1" /><style type="text/css">\n'
                "p, li { white-space: pre-wrap; }\n"
                "</style></head><body style=\" font-family:'MS Shell Dlg 2'; font-size:8pt; font-weight:400; font-style:normal;\">\n"
                '<p style=" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;">For information see:				<span style=" text-decoration: underline; color:#0000ff;">https://github.com/APEbbers/TitleBlock-WB/wiki</span></p></body></html>',
                None,
            )
        )
        self.urlLabel.setUrl(
            QCoreApplication.translate(
                "Form", "https://github.com/APEbbers/TitleBlock-WB/wiki", None
            )
        )
        pass

    # retranslateUi
