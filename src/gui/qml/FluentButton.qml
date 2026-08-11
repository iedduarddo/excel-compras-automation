import QtQuick
import QtQuick.Controls

Button {
    id: control

    property string kind: "secondary"
    property color accentColor: "#0f6cbd"
    property color textColor: kind === "primary" ? "white" : "#172033"
    property color surfaceColor: "#f7f9fc"

    implicitHeight: 44
    implicitWidth: Math.max(112, contentItem.implicitWidth + leftPadding + rightPadding)
    leftPadding: 18
    rightPadding: 18
    topPadding: 10
    bottomPadding: 10
    hoverEnabled: true

    contentItem: Text {
        text: control.text
        color: control.enabled ? control.textColor : Qt.rgba(0.25, 0.28, 0.34, 0.48)
        font.pixelSize: 14
        font.weight: control.kind === "primary" ? Font.DemiBold : Font.Medium
        horizontalAlignment: Text.AlignHCenter
        verticalAlignment: Text.AlignVCenter
        elide: Text.ElideRight
    }

    background: Rectangle {
        radius: 12
        color: {
            if (!control.enabled)
                return Qt.rgba(0.74, 0.76, 0.80, 0.34)
            if (control.kind === "primary")
                return control.down ? Qt.darker(control.accentColor, 1.18)
                                    : (control.hovered ? Qt.lighter(control.accentColor, 1.08)
                                                       : control.accentColor)
            if (control.kind === "danger")
                return control.down ? "#b42318" : (control.hovered ? "#d92d20" : "#c4322b")
            return control.down ? Qt.darker(control.surfaceColor, 1.08)
                                : (control.hovered ? Qt.lighter(control.surfaceColor, 1.12)
                                                   : control.surfaceColor)
        }
        border.width: control.visualFocus ? 2 : 1
        border.color: control.visualFocus ? control.accentColor : Qt.rgba(0.32, 0.36, 0.44, 0.24)

        Behavior on color { ColorAnimation { duration: 120 } }
    }
}
