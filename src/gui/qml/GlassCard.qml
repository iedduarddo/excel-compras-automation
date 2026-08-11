import QtQuick
import QtQuick.Controls

FocusScope {
    id: card

    default property alias contentData: content.data
    property color surfaceColor: Qt.rgba(1, 1, 1, 0.66)
    property color borderColor: Qt.rgba(1, 1, 1, 0.82)
    property int radius: 22
    property int padding: 20

    Rectangle {
        x: 2
        y: 6
        width: card.width
        height: card.height
        color: Qt.rgba(0, 0, 0, 0.10)
        radius: card.radius
    }

    Rectangle {
        id: background
        anchors.fill: parent
        color: card.surfaceColor
        radius: card.radius
        border.width: 1
        border.color: card.borderColor
    }

    Item {
        id: content
        anchors.fill: parent
        anchors.margins: card.padding
    }
}
