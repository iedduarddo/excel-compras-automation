import QtQuick
import QtQuick.Controls
import QtQuick.Dialogs as PlatformDialogs
import QtQuick.Layouts

ApplicationWindow {
    id: window
    objectName: "mainWindow"
    width: 1280
    height: 820
    minimumWidth: 940
    minimumHeight: 660
    visible: false
    title: "Excel Compras Automation"
    color: nativeMaterialActive ? "transparent" : canvasColor

    property bool nativeMaterialActive: false
    property string nativeMaterialName: "solid"
    property string nativeMaterialReason: ""
    property bool darkMode: initialDarkMode
    property int currentPage: 0
    property color canvasColor: darkMode ? "#111827" : "#eef3f8"
    property color textColor: darkMode ? "#f7f9fc" : "#172033"
    property color mutedColor: darkMode ? "#b8c2d3" : "#5f6b7a"
    property color cardColor: darkMode ? Qt.rgba(0.10, 0.14, 0.21, 0.92)
                                       : Qt.rgba(1, 1, 1, 0.90)
    property color cardBorder: darkMode ? Qt.rgba(1, 1, 1, 0.14)
                                        : Qt.rgba(1, 1, 1, 0.88)
    property color accentColor: "#0f6cbd"

    function executeQuick(command) {
        commandField.text = command
        bridge.executeRequest(command)
    }

    onClosing: function(close) {
        if (!bridge.requestClose())
            close.accepted = false
    }

    Rectangle {
        anchors.fill: parent
        opacity: window.nativeMaterialActive ? 0.90 : 1.0
        gradient: Gradient {
            GradientStop { position: 0.0; color: window.darkMode ? "#17253a" : "#dceaf7" }
            GradientStop { position: 0.48; color: window.darkMode ? "#101827" : "#f3f7fb" }
            GradientStop { position: 1.0; color: window.darkMode ? "#1b2130" : "#e8eef5" }
        }
    }

    RowLayout {
        anchors.fill: parent
        anchors.margins: 14
        spacing: 14

        GlassCard {
            id: navigation
            Layout.fillHeight: true
            Layout.preferredWidth: 222
            Layout.minimumWidth: 210
            surfaceColor: window.cardColor
            borderColor: window.cardBorder
            padding: 16

            ColumnLayout {
                anchors.fill: parent
                spacing: 10

                Text {
                    text: "EXCEL COMPRAS"
                    color: window.mutedColor
                    font.pixelSize: 11
                    font.weight: Font.DemiBold
                    font.letterSpacing: 1.4
                }
                Text {
                    text: "Automation"
                    color: window.textColor
                    font.pixelSize: 23
                    font.weight: Font.DemiBold
                    Layout.bottomMargin: 18
                }

                FluentButton {
                    Layout.fillWidth: true
                    text: "Assistente"
                    kind: window.currentPage === 0 ? "primary" : "secondary"
                    accentColor: window.accentColor
                    surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                    textColor: window.currentPage === 0 ? "white" : window.textColor
                    onClicked: window.currentPage = 0
                }
                FluentButton {
                    Layout.fillWidth: true
                    text: "Configuracoes"
                    kind: window.currentPage === 1 ? "primary" : "secondary"
                    accentColor: window.accentColor
                    surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                    textColor: window.currentPage === 1 ? "white" : window.textColor
                    onClicked: window.currentPage = 1
                }

                Item { Layout.fillHeight: true }

                Rectangle {
                    Layout.fillWidth: true
                    implicitHeight: 74
                    radius: 14
                    color: window.darkMode ? Qt.rgba(0.15, 0.45, 0.30, 0.20)
                                           : Qt.rgba(0.10, 0.62, 0.38, 0.12)
                    border.color: window.darkMode ? "#4ed18a" : "#2c8c5a"
                    border.width: 1
                    Column {
                        anchors.fill: parent
                        anchors.margins: 12
                        spacing: 5
                        Text {
                            text: "Modo seguro"
                            color: window.textColor
                            font.pixelSize: 13
                            font.weight: Font.DemiBold
                        }
                        Text {
                            width: parent.width
                            text: "Previa e confirmacao obrigatorias"
                            color: window.mutedColor
                            font.pixelSize: 11
                            wrapMode: Text.WordWrap
                        }
                    }
                }

                Text {
                    text: "v" + appVersion + "  •  " + window.nativeMaterialName
                    color: window.mutedColor
                    font.pixelSize: 11
                    Layout.topMargin: 8
                }
            }
        }

        ColumnLayout {
            Layout.fillWidth: true
            Layout.fillHeight: true
            spacing: 12

            GlassCard {
                Layout.fillWidth: true
                Layout.preferredHeight: 86
                Layout.minimumHeight: 86
                surfaceColor: window.cardColor
                borderColor: window.cardBorder
                padding: 14

                RowLayout {
                    anchors.fill: parent
                    spacing: 12

                    ColumnLayout {
                        Layout.fillWidth: true
                        spacing: 2
                        Text {
                            text: window.currentPage === 0 ? "Central de automacao" : "Preferencias"
                            color: window.textColor
                            font.pixelSize: 24
                            font.weight: Font.DemiBold
                        }
                        Text {
                            text: window.currentPage === 0
                                  ? "Importe, descreva e confirme antes de gerar qualquer copia."
                                  : "Personalize a execucao sem alterar os limites de seguranca."
                            color: window.mutedColor
                            font.pixelSize: 12
                        }
                    }

                    FluentButton {
                        objectName: "themeToggleButton"
                        text: window.darkMode ? "Modo claro" : "Modo escuro"
                        textColor: window.textColor
                        surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                        onClicked: window.darkMode = !window.darkMode
                    }

                    Rectangle {
                        implicitWidth: statusText.implicitWidth + 28
                        implicitHeight: 34
                        radius: 17
                        color: bridge.statusTone === "error" ? Qt.rgba(0.85, 0.15, 0.12, 0.16)
                             : bridge.statusTone === "warning" ? Qt.rgba(0.96, 0.62, 0.08, 0.18)
                             : bridge.statusTone === "success" ? Qt.rgba(0.10, 0.62, 0.38, 0.16)
                             : Qt.rgba(0.20, 0.45, 0.75, 0.13)
                        Text {
                            id: statusText
                            anchors.centerIn: parent
                            text: bridge.busy ? "Processando..." : bridge.statusText
                            textFormat: Text.PlainText
                            color: window.textColor
                            font.pixelSize: 12
                            font.weight: Font.Medium
                        }
                    }
                }
            }

            StackLayout {
                Layout.fillWidth: true
                Layout.fillHeight: true
                currentIndex: window.currentPage

                ScrollView {
                    id: assistantScroll
                    clip: true
                    ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

                    GridLayout {
                        width: assistantScroll.availableWidth
                        columns: width >= 1030 ? 2 : 1
                        columnSpacing: 14
                        rowSpacing: 14

                        GlassCard {
                            Layout.fillWidth: true
                            Layout.fillHeight: true
                            Layout.minimumHeight: 590
                            Layout.preferredWidth: 340
                            Layout.rowSpan: parent.columns === 2 ? 2 : 1
                            surfaceColor: window.cardColor
                            borderColor: window.cardBorder

                            ColumnLayout {
                                anchors.fill: parent
                                spacing: 12
                                RowLayout {
                                    Layout.fillWidth: true
                                    ColumnLayout {
                                        Layout.fillWidth: true
                                        spacing: 2
                                        Text {
                                            text: "1  Planilhas de entrada"
                                            color: window.textColor
                                            font.pixelSize: 17
                                            font.weight: Font.DemiBold
                                        }
                                        Text {
                                            text: "As acoes rapidas usam todos os itens listados."
                                            color: window.mutedColor
                                            font.pixelSize: 11
                                        }
                                    }
                                    Rectangle {
                                        implicitWidth: 38
                                        implicitHeight: 28
                                        radius: 14
                                        color: Qt.rgba(0.06, 0.42, 0.74, 0.15)
                                        Text {
                                            anchors.centerIn: parent
                                            text: bridge.inputs.length
                                            color: window.textColor
                                            font.weight: Font.DemiBold
                                        }
                                    }
                                }

                                Rectangle {
                                    Layout.fillWidth: true
                                    Layout.fillHeight: true
                                    radius: 15
                                    color: window.darkMode ? Qt.rgba(0, 0, 0, 0.20)
                                                           : Qt.rgba(0.96, 0.98, 1, 0.74)
                                    border.color: window.darkMode ? Qt.rgba(1, 1, 1, 0.11)
                                                                  : Qt.rgba(0.25, 0.32, 0.42, 0.14)
                                    ListView {
                                        id: inputList
                                        objectName: "inputList"
                                        anchors.fill: parent
                                        anchors.margins: 8
                                        clip: true
                                        spacing: 6
                                        model: bridge.inputs
                                        delegate: Rectangle {
                                            required property var modelData
                                            width: inputList.width
                                            height: 48
                                            radius: 11
                                            color: hoverHandler.hovered
                                                   ? Qt.rgba(0.06, 0.42, 0.74, 0.13)
                                                   : "transparent"
                                            Text {
                                                anchors.fill: parent
                                                anchors.leftMargin: 12
                                                anchors.rightMargin: 12
                                                text: modelData.name
                                                textFormat: Text.PlainText
                                                color: window.textColor
                                                verticalAlignment: Text.AlignVCenter
                                                elide: Text.ElideMiddle
                                                font.pixelSize: 13
                                            }
                                            HoverHandler { id: hoverHandler }
                                        }
                                        Text {
                                            anchors.centerIn: parent
                                            visible: bridge.inputs.length === 0
                                            text: "Nenhuma planilha adicionada"
                                            color: window.mutedColor
                                        }
                                    }
                                }

                                GridLayout {
                                    Layout.fillWidth: true
                                    columns: 2
                                    columnSpacing: 8
                                    rowSpacing: 8
                                    FluentButton {
                                        objectName: "addFilesButton"
                                        Layout.fillWidth: true
                                        text: "Adicionar"
                                        kind: "primary"
                                        accentColor: window.accentColor
                                        onClicked: fileDialog.open()
                                    }
                                    FluentButton {
                                        Layout.fillWidth: true
                                        text: "Atualizar"
                                        textColor: window.textColor
                                        surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                                        onClicked: bridge.refreshInputs()
                                    }
                                    FluentButton {
                                        Layout.columnSpan: 2
                                        Layout.fillWidth: true
                                        text: "Abrir pasta de entrada"
                                        textColor: window.textColor
                                        surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                                        onClicked: bridge.openWorkspace("input")
                                    }
                                }
                            }
                        }

                        GlassCard {
                            Layout.fillWidth: true
                            Layout.minimumHeight: 264
                            surfaceColor: window.cardColor
                            borderColor: window.cardBorder

                            ColumnLayout {
                                anchors.fill: parent
                                spacing: 10
                                Text {
                                    text: "2  O que voce quer fazer?"
                                    color: window.textColor
                                    font.pixelSize: 17
                                    font.weight: Font.DemiBold
                                }
                                RowLayout {
                                    Layout.fillWidth: true
                                    TextField {
                                        id: commandField
                                        objectName: "commandField"
                                        Layout.fillWidth: true
                                        implicitHeight: 46
                                        placeholderText: "Ex.: limpar, organizar e resumir todas"
                                        color: window.textColor
                                        placeholderTextColor: window.mutedColor
                                        enabled: !bridge.busy
                                        background: Rectangle {
                                            radius: 13
                                            color: window.darkMode ? Qt.rgba(0, 0, 0, 0.22) : "#fbfcfe"
                                            border.width: commandField.activeFocus ? 2 : 1
                                            border.color: commandField.activeFocus
                                                          ? window.accentColor
                                                          : Qt.rgba(0.25, 0.32, 0.42, 0.24)
                                        }
                                        Keys.onReturnPressed: bridge.executeRequest(text)
                                    }
                                    FluentButton {
                                        objectName: "executeButton"
                                        text: "Executar"
                                        kind: "primary"
                                        accentColor: window.accentColor
                                        enabled: !bridge.busy
                                        onClicked: bridge.executeRequest(commandField.text)
                                    }
                                    FluentButton {
                                        text: "Voz"
                                        textColor: window.textColor
                                        surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                                        enabled: bridge.voiceAvailable && !bridge.busy
                                        onClicked: bridge.captureVoice()
                                        ToolTip.visible: hovered
                                        ToolTip.text: bridge.voiceAvailable
                                                      ? "Transcreve para revisao; nao executa automaticamente"
                                                      : "Use a digitacao por voz do sistema"
                                    }
                                }
                                Text {
                                    text: "A voz somente preenche o pedido. Voce sempre revisa e executa."
                                    color: window.mutedColor
                                    font.pixelSize: 11
                                }
                                GridLayout {
                                    Layout.fillWidth: true
                                    columns: 2
                                    columnSpacing: 8
                                    rowSpacing: 8
                                    Repeater {
                                        model: [
                                            ["Reconhecer", "reconhecer todas"],
                                            ["Diagnosticar", "diagnosticar todas"],
                                            ["Limpar e organizar", "limpar e organizar todas"],
                                            ["Criar relatorio", "resumir e criar relatorio de todas"]
                                        ]
                                        FluentButton {
                                            required property var modelData
                                            Layout.fillWidth: true
                                            text: modelData[0]
                                            textColor: window.textColor
                                            surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                                            enabled: !bridge.busy
                                            onClicked: window.executeQuick(modelData[1])
                                        }
                                    }
                                }
                            }
                        }

                        GlassCard {
                            Layout.fillWidth: true
                            Layout.minimumHeight: 312
                            surfaceColor: window.cardColor
                            borderColor: window.cardBorder

                            ColumnLayout {
                                anchors.fill: parent
                                spacing: 10
                                RowLayout {
                                    Layout.fillWidth: true
                                    Text {
                                        Layout.fillWidth: true
                                        text: "3  Previa e resultado"
                                        color: window.textColor
                                        font.pixelSize: 17
                                        font.weight: Font.DemiBold
                                    }
                                    Rectangle {
                                        visible: bridge.planCount > 0
                                        implicitWidth: planBadge.implicitWidth + 22
                                        implicitHeight: 28
                                        radius: 14
                                        color: Qt.rgba(0.96, 0.62, 0.08, 0.18)
                                        Text {
                                            id: planBadge
                                            anchors.centerIn: parent
                                            text: bridge.planCount + " aguardando"
                                            color: window.textColor
                                            font.pixelSize: 11
                                            font.weight: Font.DemiBold
                                        }
                                    }
                                }
                                ScrollView {
                                    Layout.fillWidth: true
                                    Layout.fillHeight: true
                                    TextArea {
                                        id: previewArea
                                        objectName: "previewArea"
                                        text: bridge.previewMarkdown
                                        textFormat: TextEdit.PlainText
                                        readOnly: true
                                        wrapMode: TextEdit.Wrap
                                        selectByMouse: true
                                        color: window.textColor
                                        font.pixelSize: 12
                                        padding: 14
                                        background: Rectangle {
                                            radius: 14
                                            color: window.darkMode ? Qt.rgba(0, 0, 0, 0.22) : "#fbfcfe"
                                            border.color: Qt.rgba(0.25, 0.32, 0.42, 0.18)
                                        }
                                    }
                                }
                                RowLayout {
                                    Layout.fillWidth: true
                                    FluentButton {
                                        objectName: "confirmPlanButton"
                                        text: "Confirmar plano"
                                        kind: "primary"
                                        accentColor: "#107c41"
                                        enabled: bridge.canConfirm
                                        onClicked: {
                                            const token = bridge.prepareConfirmation()
                                            if (token.length > 0) {
                                                confirmationDialog.token = token
                                                confirmationDialog.open()
                                            }
                                        }
                                    }
                                    FluentButton {
                                        objectName: "cancelPlanButton"
                                        text: "Cancelar plano"
                                        textColor: window.textColor
                                        surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                                        enabled: bridge.canConfirm
                                        onClicked: bridge.cancelPlans(bridge.planRevision)
                                    }
                                    Item { Layout.fillWidth: true }
                                    FluentButton {
                                        text: "Abrir saida"
                                        textColor: window.textColor
                                        surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                                        onClicked: bridge.openWorkspace("output")
                                    }
                                }
                            }
                        }
                    }
                }

                GlassCard {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    Layout.margins: 2
                    surfaceColor: window.cardColor
                    borderColor: window.cardBorder

                    ColumnLayout {
                        anchors.fill: parent
                        spacing: 16
                        Text {
                            text: "Preferencias da automacao"
                            color: window.textColor
                            font.pixelSize: 20
                            font.weight: Font.DemiBold
                        }
                        Text {
                            text: "As alteracoes sao salvas na pasta de dados do assistente."
                            color: window.mutedColor
                            font.pixelSize: 12
                        }
                        Text { text: "Nome do candidato"; color: window.textColor; font.weight: Font.Medium }
                        TextField {
                            id: candidateField
                            Layout.fillWidth: true
                            implicitHeight: 46
                            text: bridge.candidateName
                            color: window.textColor
                            background: Rectangle {
                                radius: 13
                                color: window.darkMode ? Qt.rgba(0, 0, 0, 0.22) : "#fbfcfe"
                                border.color: candidateField.activeFocus ? window.accentColor
                                                                         : Qt.rgba(0.25, 0.32, 0.42, 0.22)
                            }
                        }
                        Switch {
                            id: pivotSwitch
                            text: "Usar Tabela Dinamica nativa do Excel"
                            checked: bridge.useNativePivot
                            enabled: bridge.nativeExcelAvailable && !bridge.busy
                            palette.text: window.textColor
                            indicator: Rectangle {
                                implicitWidth: 44
                                implicitHeight: 24
                                x: pivotSwitch.leftPadding
                                y: parent.height / 2 - height / 2
                                radius: 12
                                color: pivotSwitch.checked
                                       ? window.accentColor
                                       : (window.darkMode ? "#43516a" : "#aeb8c6")
                                border.width: pivotSwitch.visualFocus ? 2 : 1
                                border.color: pivotSwitch.visualFocus
                                              ? window.accentColor
                                              : Qt.rgba(0.25, 0.32, 0.42, 0.30)
                                Rectangle {
                                    x: pivotSwitch.checked ? parent.width - width - 2 : 2
                                    y: 2
                                    width: 20
                                    height: 20
                                    radius: 10
                                    color: "#ffffff"
                                    Behavior on x { NumberAnimation { duration: 120 } }
                                }
                            }
                        }
                        Text {
                            text: bridge.nativeExcelAvailable
                                  ? "O Excel Desktop sera aberto em segundo plano quando necessario."
                                  : "No macOS, o resumo compativel e usado automaticamente."
                            color: window.mutedColor
                            font.pixelSize: 11
                        }
                        Text { text: "Intervalo do monitor (segundos)"; color: window.textColor; font.weight: Font.Medium }
                        TextField {
                            id: intervalField
                            implicitWidth: 180
                            implicitHeight: 44
                            text: Number(bridge.pollIntervalSeconds).toLocaleString(Qt.locale(), "f", 1)
                            color: window.textColor
                            validator: DoubleValidator { bottom: 0.5; top: 60.0; decimals: 1 }
                            background: Rectangle {
                                radius: 13
                                color: window.darkMode ? Qt.rgba(0, 0, 0, 0.22) : "#fbfcfe"
                                border.color: intervalField.activeFocus ? window.accentColor
                                                                        : Qt.rgba(0.25, 0.32, 0.42, 0.22)
                            }
                        }
                        RowLayout {
                            spacing: 10
                            FluentButton {
                                text: "Salvar configuracoes"
                                kind: "primary"
                                accentColor: window.accentColor
                                enabled: !bridge.busy && intervalField.acceptableInput
                                onClicked: bridge.saveSettings(
                                    candidateField.text,
                                    pivotSwitch.checked,
                                    Number.fromLocaleString(Qt.locale(), intervalField.text)
                                )
                            }
                            FluentButton {
                                text: bridge.monitoring ? "Parar monitor" : "Iniciar monitor"
                                textColor: window.textColor
                                surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                                enabled: !bridge.busy
                                onClicked: bridge.toggleMonitor()
                            }
                        }
                        Rectangle {
                            Layout.fillWidth: true
                            implicitHeight: 76
                            radius: 14
                            color: Qt.rgba(0.06, 0.42, 0.74, 0.12)
                            border.color: Qt.rgba(0.06, 0.42, 0.74, 0.36)
                            Text {
                                anchors.fill: parent
                                anchors.margins: 14
                                text: "O monitor apenas atualiza a lista de entrada. Ele nunca executa ou confirma alteracoes automaticamente."
                                color: window.textColor
                                wrapMode: Text.WordWrap
                                verticalAlignment: Text.AlignVCenter
                                font.pixelSize: 12
                            }
                        }
                        Item { Layout.fillHeight: true }
                    }
                }
            }

            ProgressBar {
                Layout.fillWidth: true
                indeterminate: true
                visible: bridge.busy
            }
        }
    }

    PlatformDialogs.FileDialog {
        id: fileDialog
        title: "Adicionar planilhas"
        fileMode: PlatformDialogs.FileDialog.OpenFiles
        nameFilters: ["Planilhas Excel (*.xlsx *.xlsm)"]
        onAccepted: bridge.importFiles(selectedFiles)
    }

    Dialog {
        id: confirmationDialog
        objectName: "confirmationDialog"
        property string token: ""
        anchors.centerIn: parent
        modal: true
        closePolicy: Popup.CloseOnEscape
        width: Math.min(520, window.width - 64)
        padding: 24
        onOpened: keepReviewingButton.forceActiveFocus()
        background: Rectangle {
            radius: 22
            color: window.darkMode ? "#202a3b" : "#ffffff"
            border.width: 1
            border.color: window.cardBorder
        }
        contentItem: ColumnLayout {
            spacing: 14
            Text {
                text: "Confirmar alteracoes?"
                color: window.textColor
                font.pixelSize: 21
                font.weight: Font.DemiBold
            }
            Text {
                Layout.fillWidth: true
                text: "Executar " + bridge.planCount + " plano(s) desta previa? Os resultados serao gravados em novas copias e os arquivos originais serao preservados."
                color: window.mutedColor
                wrapMode: Text.WordWrap
                font.pixelSize: 13
            }
            RowLayout {
                Layout.fillWidth: true
                Layout.topMargin: 8
                Item { Layout.fillWidth: true }
                FluentButton {
                    id: keepReviewingButton
                    text: "Continuar revisando"
                    textColor: window.textColor
                    surfaceColor: window.darkMode ? "#253247" : "#f7f9fc"
                    onClicked: confirmationDialog.close()
                }
                FluentButton {
                    text: "Confirmar e gerar copias"
                    kind: "primary"
                    accentColor: "#107c41"
                    onClicked: {
                        const savedToken = confirmationDialog.token
                        confirmationDialog.close()
                        bridge.confirmPlans(savedToken)
                    }
                }
            }
        }
    }

    Dialog {
        id: errorDialog
        anchors.centerIn: parent
        modal: true
        closePolicy: Popup.CloseOnEscape
        property string message: ""
        width: Math.min(500, window.width - 64)
        padding: 24
        background: Rectangle {
            radius: 22
            color: window.darkMode ? "#202a3b" : "#ffffff"
            border.width: 1
            border.color: window.cardBorder
        }
        contentItem: ColumnLayout {
            spacing: 14
            Text {
                text: "A automacao nao foi concluida"
                color: window.textColor
                font.pixelSize: 19
                font.weight: Font.DemiBold
            }
            Text {
                Layout.fillWidth: true
                text: errorDialog.message
                textFormat: Text.PlainText
                color: window.mutedColor
                wrapMode: Text.WordWrap
                font.pixelSize: 13
            }
            RowLayout {
                Layout.fillWidth: true
                Item { Layout.fillWidth: true }
                FluentButton {
                    text: "Entendi"
                    kind: "primary"
                    accentColor: window.accentColor
                    onClicked: errorDialog.close()
                }
            }
        }
    }

    Connections {
        target: bridge
        function onVoiceTranscriptReady(text) {
            commandField.text = text
            commandField.forceActiveFocus()
        }
        function onErrorRaised(message) {
            errorDialog.message = message
            errorDialog.open()
        }
        function onCloseBlocked() {
            errorDialog.message = "Aguarde a operacao atual terminar antes de fechar."
            errorDialog.open()
        }
    }

    Shortcut { sequences: [StandardKey.Open]; onActivated: fileDialog.open() }
    Shortcut { sequence: "F5"; onActivated: bridge.refreshInputs() }
    Shortcut { sequence: "Ctrl+L"; onActivated: { window.currentPage = 0; commandField.forceActiveFocus() } }
}
