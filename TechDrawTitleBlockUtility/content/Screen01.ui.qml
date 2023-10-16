/*
This is a UI file (.ui.qml) that is intended to be edited in Qt Design Studio only.
It is supposed to be strictly declarative and only uses a subset of QML. If you edit
this file manually, you might introduce QML code that is not supported by Qt Design Studio.
Check out https://doc.qt.io/qtcreator/creator-quick-ui-forms.html for details on .ui.qml files.
*/

import QtQuick 6.5
import QtQuick.Controls 6.5
import TechDrawTitleBlockUtility

Rectangle {
    id: rectangle
    width: Constants.width
    height: Constants.height

    color: Constants.backgroundColor

    Frame {
        id: frame
        x: 8
        y: 8
        width: 360
        height: 208

        Text {
            id: text2
            x: 0
            y: 0
            width: 338
            height: 28
            text: qsTr("Add te the following properties can be added  to the spreadsheet even when these are not present on drawing:")
            font.pixelSize: 12
            horizontalAlignment: Text.AlignLeft
            wrapMode: Text.Wrap
        }

        CheckBox {
            id: lengthUnit
            x: 0
            y: 47
            text: qsTr("Length unit")
        }

        CheckBox {
            id: angleUnit
            x: 0
            y: 93
            text: qsTr("Angle unit")
        }

        CheckBox {
            id: massUnit
            x: 0
            y: 139
            text: qsTr("Mass unit")
        }
    }

    Frame {
        id: frame1
        x: 8
        y: 222
        width: 360
        height: 101

        CheckBox {
            id: externalSheet
            x: 0
            y: 0
            text: qsTr("Use external spreadsheet")
        }

        Button {
            id: addExcelPath
            x: 273
            y: 46
            width: 63
            height: 29
            text: qsTr("Browse")
        }

        TextField {
            id: excelPath
            x: 0
            y: 46
            width: 267
            height: 29
            text: "Some path"
            placeholderText: qsTr("Text Field")
        }
    }
    states: [
        State {
            name: "clicked"
        }
    ]
}


