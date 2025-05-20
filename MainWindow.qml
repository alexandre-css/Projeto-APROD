import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Effects
import QtCharts

ApplicationWindow {
    width: 1440
    height: 900
    visible: true
    visibility: ApplicationWindow.Maximized
    Material.theme: Material.Light
    Material.primary: "#3cb3e6"
    Material.accent: "#1976d2"

    property string currentPage: "dashboard"

    ListModel { id: pesosModel }

    Rectangle {
        anchors.fill: parent
        gradient: Gradient {
            GradientStop { position: 0.0; color: "#f9fafb" }
            GradientStop { position: 1.0; color: "#e5e7ea" }
        }

        RowLayout {
            anchors.fill: parent
            spacing: 0

            // MENU LATERAL
            Rectangle {
                id: sidebar
                width: hovered ? 220 : 110
                color: "#fff"
                border.color: "#e0e0e0"
                Layout.fillHeight: true
                z: 3
                property bool hovered: false

                Behavior on width { NumberAnimation { duration: 180; easing.type: Easing.OutQuad } }

                HoverHandler {
                    id: hoverHandler
                    acceptedDevices: PointerDevice.Mouse | PointerDevice.TouchPad
                    cursorShape: Qt.PointingHandCursor
                    onHoveredChanged: sidebar.hovered = hovered
                }

                ColumnLayout {
                    anchors.fill: parent
                    spacing: 40

                    Item { Layout.fillHeight: true }

                    Repeater {
                        model: [
                            { icon: "dashboard.png", label: "Dashboard", page: "dashboard" },
                            { icon: "compara.png", label: "Comparar", page: "comparar" },
                            { icon: "bars.png", label: "Gráficos", page: "graficos" },
                            { icon: "peso.png", label: "Peso", page: "pesos" },
                            { icon: "settings.png", label: "Configurações", page: "config" }
                        ]
                        RowLayout {
                            Layout.alignment: Qt.AlignLeft
                            spacing: 14
                            width: parent.width

                            Item {
                                width: 15
                                Layout.alignment: Qt.AlignVCenter
                            }

                            Image {
                                source: modelData.icon ? "assets/icons/" + modelData.icon : ""
                                Layout.preferredWidth: 56
                                Layout.preferredHeight: 56
                                Layout.minimumWidth: 56
                                Layout.minimumHeight: 56
                                Layout.maximumWidth: 56
                                Layout.maximumHeight: 56
                                fillMode: Image.PreserveAspectFit
                                Layout.alignment: Qt.AlignVCenter
                            }
                            Text {
                                text: modelData.label ? modelData.label : ""
                                visible: sidebar.hovered
                                opacity: sidebar.hovered ? 1 : 0
                                font.pixelSize: 17
                                color: "#232946"
                                verticalAlignment: Text.AlignVCenter
                                horizontalAlignment: Text.AlignLeft
                                elide: Text.ElideRight
                                width: 120
                                Layout.alignment: Qt.AlignVCenter
                                Behavior on opacity { NumberAnimation { duration: 120 } }
                            }
                            MouseArea {
                                anchors.fill: parent
                                hoverEnabled: true
                                cursorShape: Qt.PointingHandCursor
                                onClicked: currentPage = modelData.page
                            }
                            Item {
                                Layout.fillWidth: true
                                visible: sidebar.hovered
                            }
                        }
                    }

                    Item { Layout.fillHeight: true }
                }
            }

            // ÁREA PRINCIPAL
            Rectangle {
                color: "transparent"
                Layout.fillWidth: true
                Layout.fillHeight: true

                Loader {
                    id: pageLoader
                    anchors.fill: parent
                    sourceComponent: currentPage === "dashboard" ? dashboardPage
                                        : currentPage === "pesos" ? pesosPage
                                        : dashboardPage

                    onLoaded: {
                        if (currentPage === "dashboard" && backend && backend.atualizar_grafico) {
                            backend.atualizar_grafico()
                        }
                        if (currentPage === "pesos" && backend && backend.atualizar_tabela_pesos) {
                            backend.atualizar_tabela_pesos()
                            pesosModel.clear()
                            for (var i = 0; i < backend.tabela_pesos.length; ++i) {
                                pesosModel.append(backend.tabela_pesos[i])
                            }
                        }
                    }
                }
            }
        }
    }

    Popup {
        id: saveConfirmation
        width: 400
        height: 140
        modal: true
        focus: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent

        Rectangle {
            anchors.fill: parent
            color: "#ffffff"
            radius: 8

            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20

                Text {
                    text: "Pesos salvos com sucesso!"
                    font.pixelSize: 18
                    color: "#3cb3e6"
                    Layout.alignment: Qt.AlignHCenter
                }

                Text {
                    text: "Arquivo disponível na pasta:\n'Documentos/Analyzer-Dev-APROD/pesos salvos'"
                    font.pixelSize: 14
                    color: "#232946"
                    wrapMode: Text.WordWrap
                    Layout.fillWidth: true
                }
            }
        }
    }


////////////////////////////////////// COMEÇO DASHBOARD PAGE

    Component {
        id: dashboardPage

        // CONTAINER INSTITUCIONAL ANCORADO NO TOPO (sem height fixo)
        Rectangle {
            color: "transparent"
            Layout.fillWidth: true
            Layout.fillHeight: true
            // ColumnLayout principal do dashboardPage
            ColumnLayout {
                anchors.top: parent.top
                anchors.topMargin: 20
                anchors.horizontalCenter: parent.horizontalCenter
                width: Math.min(parent.width - 140, 1200)
                spacing: 10

                // KPIs
                RowLayout {
                    Layout.fillWidth: true
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 18
                    Layout.bottomMargin: 4

                    // KPI 1: Total de Minutas
                    Rectangle {
                        Layout.preferredWidth: 320
                        height: 110
                        radius: 18
                        border.color: "#3cb3e6"
                        border.width: 2
                        color: Qt.rgba(245/255, 250/255, 255/255, 1)
                        layer.enabled: true
                        layer.effect: MultiEffect {
                            shadowEnabled: true
                            shadowColor: "#b8d6f3"
                            shadowBlur: 1.0
                            shadowHorizontalOffset: 4
                        }
                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 14
                            spacing: 12
                            Image {
                                source: "assets/icons/sum.png"
                                Layout.preferredWidth: 38
                                Layout.preferredHeight: 38
                                Layout.minimumWidth: 38
                                Layout.minimumHeight: 38
                                Layout.maximumWidth: 38
                                Layout.maximumHeight: 38
                                fillMode: Image.PreserveAspectFit
                            }
                            Column {
                                spacing: 2
                                Text {
                                    text: "Total de Minutas"
                                    font.pixelSize: 18
                                    color: "#3cb3e6"
                                    font.bold: true
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                }
                                Text {
                                    text: kpis && kpis.minutas ? kpis.minutas : ""
                                    font.pixelSize: 34
                                    color: "#232946"
                                    font.bold: true
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                }
                            }
                        }
                    }

                    // KPI 2: Dia Mais Produtivo
                    Rectangle {
                        Layout.preferredWidth: 320
                        height: 110
                        radius: 18
                        border.color: "#3cb3e6"
                        border.width: 2
                        color: Qt.rgba(245/255, 250/255, 255/255, 1)
                        layer.enabled: true
                        layer.effect: MultiEffect {
                            shadowEnabled: true
                            shadowColor: "#b8d6f3"
                            shadowBlur: 1.0
                            shadowHorizontalOffset: 4
                        }
                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 14
                            spacing: 12
                            Image {
                                source: "assets/icons/calendar.png"
                                Layout.preferredWidth: 38
                                Layout.preferredHeight: 38
                                Layout.minimumWidth: 38
                                Layout.minimumHeight: 38
                                Layout.maximumWidth: 38
                                Layout.maximumHeight: 38
                                fillMode: Image.PreserveAspectFit
                            }
                            Column {
                                spacing: 2
                                Text {
                                    text: "Dia Mais Produtivo"
                                    font.pixelSize: 18
                                    color: "#3cb3e6"
                                    font.bold: true
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                }
                                Text {
                                    text: kpis && kpis.dia ? kpis.dia : ""
                                    font.pixelSize: 20
                                    color: "#232946"
                                    font.bold: true
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                }
                            }
                        }
                    }

                    // KPI 3: Top 3 Usuários
                    Rectangle {
                        Layout.preferredWidth: 320
                        height: 110
                        radius: 18
                        border.color: "#3cb3e6"
                        border.width: 2
                        color: Qt.rgba(245/255, 250/255, 255/255, 1)
                        layer.enabled: true
                        layer.effect: MultiEffect {
                            shadowEnabled: true
                            shadowColor: "#b8d6f3"
                            shadowBlur: 1.0
                            shadowHorizontalOffset: 4
                        }
                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 14
                            spacing: 12
                            Image {
                                source: "assets/icons/top.png"
                                Layout.preferredWidth: 45
                                Layout.preferredHeight: 45
                                Layout.minimumWidth: 45
                                Layout.minimumHeight: 45
                                Layout.maximumWidth: 45
                                Layout.maximumHeight: 45
                                fillMode: Image.PreserveAspectFit
                            }
                            Column {
                                spacing: 2
                                Text {
                                    text: "Top 3 Usuários"
                                    font.pixelSize: 18
                                    color: "#3cb3e6"
                                    font.bold: true
                                    elide: Text.ElideRight
                                    wrapMode: Text.WordWrap
                                    width: 180
                                }
                                Text {
                                    text: kpis && kpis.top3 ? kpis.top3 : ""
                                    font.pixelSize: 12
                                    color: "#232946"
                                    font.bold: true
                                    wrapMode: Text.WordWrap
                                    elide: Text.ElideRight
                                    width: 180
                                    maximumLineCount: 3
                                }
                            }
                        }
                    }
                }
                // Espaço entre KPIs e botões/campo de meses
                Item { height: 2 }

                // Botões e campo de meses centralizados
                RowLayout {
                    Layout.fillWidth: true
                    Layout.alignment: Qt.AlignHCenter
                    Layout.topMargin: 0 // Margem consistente após os KPIs
                    Layout.bottomMargin: 4  // Mínimo espaço após controles
                    spacing: 18

                    // Campo de meses e botão Importar Excel
                    RowLayout {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 60
                        Layout.bottomMargin: 8
                        Layout.topMargin: 8
                        spacing: 18
                        Layout.alignment: Qt.AlignHCenter

                        // Botão Importar Excel (mantém largura fixa)
                        Rectangle {
                            id: importButton
                            Layout.preferredWidth: 220
                            Layout.preferredHeight: 48
                            radius: 14
                            gradient: Gradient {
                                GradientStop { position: 0.0; color: "#3cb3e6" }
                                GradientStop { position: 1.0; color: "#1976d2" }
                            }
                            property bool hovered: false
                            scale: hovered ? 1.03 : 1.0
                            Behavior on scale { NumberAnimation { duration: 120; easing.type: Easing.OutQuad } }

                            layer.enabled: true
                            layer.effect: MultiEffect {
                                shadowEnabled: true
                                shadowColor: "#b8d6f3"
                                shadowBlur: 1.0
                                shadowHorizontalOffset: 2
                                shadowVerticalOffset: 2
                            }

                            MouseArea {
                                anchors.fill: parent
                                hoverEnabled: true
                                onEntered: importButton.hovered = true
                                onExited: importButton.hovered = false
                                cursorShape: Qt.PointingHandCursor
                                onClicked: backend.importar_arquivo_excel()
                            }

                            RowLayout {
                                anchors.centerIn: parent
                                spacing: 8

                                Image {
                                    source: "assets/icons/excel.png"
                                    Layout.preferredWidth: 30
                                    Layout.preferredHeight: 30
                                    Layout.minimumWidth: 30
                                    Layout.minimumHeight: 30
                                    Layout.maximumWidth: 30
                                    Layout.maximumHeight: 30
                                    fillMode: Image.PreserveAspectFit
                                    Layout.alignment: Qt.AlignVCenter
                                }

                                Text {
                                    text: "Importar Excel"
                                    color: "#fff"
                                    font.bold: true
                                    font.pixelSize: 17
                                }
                            }
                        }

                        // Campo de meses (começa após o botão, preenche o resto do espaço)
                        Rectangle {
                            id: arquivosCarregadosBox
                            Layout.fillWidth: true
                            Layout.preferredHeight: 85
                            radius: 14
                            color: "#f5faff"
                            border.width: 2
                            border.color: "#3cb3e6"

                            // Texto de instrução
                            Text {
                                id: chipLegend
                                text: "Clique nos meses para ativar/desativar filtros:"
                                font.pixelSize: 12
                                color: "#3cb3e6"
                                anchors.top: parent.top
                                anchors.left: parent.left
                                anchors.margins: 8
                                height: 18 // Altura fixa para evitar sobreposições
                            }

                            // Container com rolagem horizontal
                            Flickable {
                                id: mesesFlickable
                                anchors.fill: parent
                                anchors.topMargin: 30  // IMPORTANTE: Valor fixo para evitar sobreposições
                                anchors.leftMargin: 10
                                anchors.rightMargin: 10
                                anchors.bottomMargin: 8
                                contentWidth: mesesFlow.width
                                flickableDirection: Flickable.HorizontalFlick
                                clip: true

                                // Usar Flow em vez de Row para quebrar linhas se necessário
                                Flow {
                                    id: mesesFlow
                                    width: parent.width - 20
                                    spacing: 10

                                    Repeater {
                                        id: mesesRepeater
                                        model: backend && backend.mesesDisponiveis ? backend.mesesDisponiveis : []

                                        Rectangle {
                                            width: monthText.width + 36
                                            height: 36
                                            radius: 10
                                            property bool forceUpdate: false

                                            // Timer para forçar atualização mais frequente
                                            Timer {
                                                interval: 20
                                                running: true
                                                repeat: true
                                                onTriggered: {
                                                    parent.forceUpdate = !parent.forceUpdate
                                                    parent.isActive = backend && backend.mesesAtivos &&
                                                                     backend.mesesAtivos.indexOf(modelData) >= 0
                                                }
                                            }

                                            property bool isActive: backend && backend.mesesAtivos &&
                                                                   backend.mesesAtivos.indexOf(modelData) >= 0

                                            // Cores com MAIOR contraste
                                            color: isActive ? "#3cb3e6" : "#f7f7f7"  // Azul forte vs quase branco
                                            border.color: isActive ? "#1976d2" : "#dddddd"
                                            border.width: isActive ? 3 : 1
                                            opacity: isActive ? 1.0 : 0.6  // Diferença maior de opacidade

                                            // Efeitos visuais APENAS para itens ativos
                                            layer.enabled: isActive
                                            layer.effect: MultiEffect {
                                                shadowEnabled: isActive
                                                shadowColor: "#1976d2"
                                                shadowBlur: 1.0
                                                shadowVerticalOffset: 2
                                            }

                                            // Marcador visual claro para itens ativos
                                            Rectangle {
                                                width: 8
                                                height: 8
                                                radius: 4
                                                color: "#ffffff"
                                                visible: parent.isActive
                                                anchors {
                                                    right: parent.right
                                                    top: parent.top
                                                    margins: 6
                                                }
                                            }

                                            Text {
                                                id: monthText
                                                text: modelData
                                                anchors.centerIn: parent
                                                font.pixelSize: 15
                                                font.bold: parent.isActive
                                                color: parent.isActive ? "#ffffff" : "#232946"
                                            }

                                            MouseArea {
                                                anchors.fill: parent
                                                cursorShape: Qt.PointingHandCursor
                                                onClicked: {
                                                    console.log("Clicado: " + modelData)
                                                    backend.toggleMesAtivo(modelData)
                                                }
                                                hoverEnabled: true
                                                onEntered: parent.opacity = parent.isActive ? 1.0 : 0.85
                                                onExited: parent.opacity = parent.isActive ? 1.0 : 0.7
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Indicador de mais conteúdo
                        Rectangle {
                            visible: mesesFlickable.contentWidth > mesesFlickable.width
                            anchors.right: parent.right
                            anchors.verticalCenter: parent.verticalCenter
                            width: 20
                            height: 20
                            radius: 10
                            color: "#3cb3e6"
                            opacity: 0.7

                            Text {
                                anchors.centerIn: parent
                                text: "›"
                                color: "white"
                                font.bold: true
                                font.pixelSize: 16
                            }
                        }
                    }
                }

                // Espaço entre botões/campo de meses e gráfico
                Item { height: 2 }

                // Gráfico centralizado (altura garantida)
                Rectangle {
                    width: Math.min(parent.width - 160, 1000)
                    height: 620
                    radius: 22
                    border.color: "#3cb3e6"
                    border.width: 2
                    anchors.horizontalCenter: parent.horizontalCenter

                    ChartView {
                        anchors.fill: parent
                        antialiasing: true
                        backgroundColor: "#fff"
                        legend.visible: false
                        animationOptions: ChartView.SeriesAnimations
                        animationDuration: 900

                        HorizontalBarSeries {
                            id: serieMinutas
                            axisY: eixoUsuarios
                            axisX: eixoValores
                            barWidth: 0.2
                            labelsVisible: true
                            labelsFormat: "@value"
                            labelsPosition: AbstractBarSeries.LabelsOutsideEnd
                            BarSet {
                                label: "Minutas"
                                values: backend && backend.valores ? backend.valores : []
                                brushFilename: "assets/gradients/bargradient.png"
                                labelFont: Qt.font({ pixelSize: 16, bold: true, family: "Arial" })
                                labelColor: "#1976d2"
                            }
                        }
                        BarCategoryAxis {
                            id: eixoUsuarios
                            categories: backend && backend.nomes ? backend.nomes : []
                            labelsFont.pixelSize: 12
                        }
                        ValueAxis {
                            id: eixoValores
                            min: 0
                            max: backend && backend.valores && backend.valores.length > 0
                                ? Math.max.apply(Math, backend.valores) * 1.1
                                : 10
                            labelFormat: "%.0f"
                        }
                        title: "Minutas por Usuário"
                        titleFont: Qt.font({ pixelSize: 12, bold: true, family: "Arial" })
                        titleColor: "#1976d2"
                    }
                }
            } // Fecha ColumnLayout
        } // Fecha Rectangle
    } // Fecha Component dashboardPage

                                        // PÁGINA DE ATRIBUIÇÃO DE PESOS

    property bool ordemCrescente: true
    property int sortColumn: 0
    property bool sortAscending: true

    Component {
        id: pesosPage

        Rectangle {
            color: "transparent"
            Layout.fillWidth: true
            Layout.fillHeight: true

            ColumnLayout {
                anchors.centerIn: parent
                spacing: 24
                width: 1100

                Text {
                    text: "Atribuição de Pesos"
                    font.pixelSize: 28
                    font.bold: true
                    color: "#3cb3e6"
                    horizontalAlignment: Text.AlignHCenter
                    Layout.alignment: Qt.AlignHCenter
                }

                Rectangle {
                    width: 1100
                    height: 700
                    color: "#fff"
                    border.color: "#3cb3e6"
                    border.width: 2
                    radius: 18
                    clip: true
                    Layout.alignment: Qt.AlignHCenter

                    Column {
                        anchors.fill: parent
                        spacing: 0

                        Item {
                            width: parent.width
                            height: 44
                            RowLayout {
                                anchors.fill: parent
                                spacing: 0

                                Item {
                                    width: 800
                                    height: 44
                                    Text {
                                        text: "Tipo de Agendamento"
                                        anchors.centerIn: parent
                                        font.pixelSize: 18
                                        font.bold: true
                                        color: "#3cb3e6"
                                    }
                                }

                                Item {
                                    width: 260
                                    height: 44
                                    Text {
                                        text: "Peso"
                                        anchors.centerIn: parent
                                        font.pixelSize: 18
                                        font.bold: true
                                        color: "#3cb3e6"
                                    }
                                }
                            }

                            Rectangle {
                                anchors.bottom: parent.bottom
                                width: parent.width
                                height: 1
                                color: "#3cb3e6"
                            }
                        }

                        Flickable {
                            id: flick
                            width: parent.width
                            height: parent.height - 44
                            contentHeight: column.height
                            clip: true

                            Column {
                                id: column
                                width: parent.width

                                Repeater {
                                    model: pesosModel

                                    Item {
                                        width: column.width
                                        height: 44

                                        RowLayout {
                                            anchors.fill: parent
                                            spacing: 0

                                            Item {
                                                width: 800
                                                height: 44
                                                Text {
                                                    text: model.tipo
                                                    anchors.centerIn: parent
                                                    font.pixelSize: 17
                                                    color: "#232946"
                                                    elide: Text.ElideRight
                                                }
                                            }

                                            Item {
                                                width: 260
                                                height: 44
                                                RowLayout {
                                                    anchors.centerIn: parent
                                                    spacing: 8

                                                    // Botão -
                                                    Rectangle {
                                                        width: 32
                                                        height: 32
                                                        radius: 8
                                                        color: "#e3f2fd"
                                                        border.color: "#3cb3e6"
                                                        border.width: 1
                                                        MouseArea {
                                                            anchors.fill: parent
                                                            cursorShape: Qt.PointingHandCursor
                                                            onClicked: {
                                                                let value = parseFloat(textField.text)
                                                                value = isNaN(value) ? 1.0 : Math.max(1.0, value - 1.0)
                                                                textField.text = value.toFixed(1)
                                                                pesosModel.setProperty(index, "peso", value)
                                                                backend.atualizarPeso(index, value)
                                                            }
                                                        }
                                                        Text {
                                                            text: "-"
                                                            anchors.centerIn: parent
                                                            font.pixelSize: 18
                                                            font.bold: true
                                                            color: "#3cb3e6"
                                                        }
                                                    }

                                                    // CAMPO DE DIGITAÇÃO (largura fixa para "10.0")
                                                    TextField {
                                                        id: textField
                                                        text: model.peso !== undefined ? Number(model.peso).toFixed(1) : ""
                                                        Layout.preferredWidth: 70
                                                        height: 32
                                                        font.pixelSize: 16
                                                        horizontalAlignment: TextInput.AlignHCenter
                                                        verticalAlignment: TextInput.AlignVCenter
                                                        validator: IntValidator { bottom: 1; top: 10 }
                                                        inputMethodHints: Qt.ImhDigitsOnly
                                                        background: Rectangle {
                                                            radius: 6
                                                            border.color: "#3cb3e6"
                                                            border.width: 1
                                                            color: "#fff"
                                                        }
                                                        onEditingFinished: {
                                                            let value = parseInt(text)
                                                            value = Math.min(10, Math.max(1, isNaN(value) ? 1 : value))
                                                            text = value.toFixed(1)
                                                            pesosModel.setProperty(index, "peso", value)
                                                            backend.atualizarPeso(index, value)
                                                        }
                                                    }

                                                    // Botão +
                                                    Rectangle {
                                                        width: 32
                                                        height: 32
                                                        radius: 8
                                                        color: "#e3f2fd"
                                                        border.color: "#3cb3e6"
                                                        border.width: 1
                                                        MouseArea {
                                                            anchors.fill: parent
                                                            cursorShape: Qt.PointingHandCursor
                                                            onClicked: {
                                                                let value = parseFloat(textField.text)
                                                                value = isNaN(value) ? 1.0 : Math.min(10.0, value + 1.0)
                                                                textField.text = value.toFixed(1)
                                                                pesosModel.setProperty(index, "peso", value)
                                                                backend.atualizarPeso(index, value)
                                                            }
                                                        }
                                                        Text {
                                                            text: "+"
                                                            anchors.centerIn: parent
                                                            font.pixelSize: 18
                                                            font.bold: true
                                                            color: "#3cb3e6"
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        Rectangle {
                                            anchors.bottom: parent.bottom
                                            width: parent.width
                                            height: 1
                                            color: "#3cb3e6"
                                            visible: index < pesosModel.count - 1
                                        }
                                    }
                                }
                            }

                            ScrollBar.vertical: ScrollBar {
                                width: 12
                                policy: ScrollBar.AlwaysOn
                                contentItem: Rectangle {
                                    radius: 6
                                    color: "#3cb3e6"
                                }
                            }
                        }
                    }
                }

                RowLayout {
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 18

                    Rectangle {
                        width: 220
                        height: 48
                        radius: 14
                        gradient: Gradient {
                            GradientStop { position: 0.0; color: "#3cb3e6" }
                            GradientStop { position: 1.0; color: "#1976d2" }
                        }

                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                backend.salvarPesos()
                                saveConfirmation.open()
                            }
                        }

                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 8
                            Image {
                                source: "assets/icons/peso2.png"
                                Layout.preferredWidth: 30
                                Layout.preferredHeight: 30
                                Layout.minimumWidth: 30
                                Layout.minimumHeight: 30
                                Layout.maximumWidth: 30
                                Layout.maximumHeight: 30
                                fillMode: Image.PreserveAspectFit
                                Layout.alignment: Qt.AlignVCenter
                            }
                            Text {
                                text: "Salvar Pesos"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 17
                                verticalAlignment: Text.AlignVCenter
                            }
                        }
                    }

                    // Botão Restaurar Padrão
                    Rectangle {
                        width: 220
                        height: 48
                        radius: 14
                        gradient: Gradient {
                            GradientStop { position: 0.0; color: "#3cb3e6" }
                            GradientStop { position: 1.0; color: "#1976d2" }
                        }
                        property bool hovered: false
                        scale: hovered ? 1.03 : 1.0
                        Behavior on scale { NumberAnimation { duration: 120; easing.type: Easing.OutQuad } }
                        layer.enabled: true
                        layer.effect: MultiEffect {
                            shadowEnabled: true
                            shadowColor: "#b8d6f3"
                            shadowBlur: 1.0
                            shadowHorizontalOffset: 2
                            shadowVerticalOffset: 2
                        }
                        MouseArea {
                            anchors.fill: parent
                            hoverEnabled: true
                            onEntered: parent.hovered = true
                            onExited: parent.hovered = false
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                backend.restaurarPesosPadrao()
                                // Forçar reload completo do modelo da tabela
                                pesosModel.clear()
                                for (var i = 0; i < backend.tabela_pesos.length; ++i) {
                                    pesosModel.append(backend.tabela_pesos[i])
                                }
                            }
                        }
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 8
                            Image {
                                source: "assets/icons/peso4.png"
                                Layout.preferredWidth: 30
                                Layout.preferredHeight: 30
                                Layout.minimumWidth: 30
                                Layout.minimumHeight: 30
                                Layout.maximumWidth: 30
                                Layout.maximumHeight: 30
                                fillMode: Image.PreserveAspectFit
                                Layout.alignment: Qt.AlignVCenter
                            }
                            Text {
                                text: "Restaurar Padrão"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 17
                                verticalAlignment: Text.AlignVCenter
                            }
                        }
                    }
                }
            }
        }
    }
}
