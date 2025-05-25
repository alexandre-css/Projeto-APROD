import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Effects
import QtCharts
import QtQuick.Controls.Material

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
                            { icon: "semana.png", label: "Semana", page: "semana" },
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
                                    : currentPage === "comparar" ? compararPage
                                    : currentPage === "semana" ? semanaPage
                                    : currentPage === "pesos" ? pesosPage
                                    : currentPage === "config" ? configPage
                                    : dashboardPage
                    
                    
                    onLoaded: {
                        if (currentPage === "dashboard" && backend && backend.atualizar_grafico) {
                            console.log("Chamando atualizar_grafico");
                            backend.atualizar_grafico()
                        }
                        if (currentPage === "pesos" && backend && backend.atualizar_tabela_pesos) {
                            console.log("Chamando atualizar_tabela_pesos");
                            backend.atualizar_tabela_pesos()
                            pesosModel.clear()
                            for (var i = 0; i < backend.tabela_pesos.length; ++i) {
                                pesosModel.append(backend.tabela_pesos[i])
                            }
                        }
                        if (currentPage === "semana" && backend && typeof backend.gerarTabelaSemana === "function") {
                            console.log("Página semana selecionada - atualizando dados");
                            backend.gerarTabelaSemana();
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
                            width: Math.min(mesesFlow.width + 32, 1100) // 32 = margens internas + bordas
                            height: 85
                            radius: 14
                            color: "#f5faff"
                            border.width: 2
                            border.color: "#3cb3e6"

                            // Texto de instrução
                            Text {
                                id: chipLegend
                                text: "Clique nos meses para ativar/desativar"
                                font.pixelSize: 12
                                color: "#3cb3e6"
                                anchors.top: parent.top
                                anchors.left: parent.left
                                anchors.margins: 8
                                height: 18
                            }

                            Flickable {
                                id: mesesFlickable
                                width: Math.min(mesesFlow.width + 20, 1060) // 20 = margens internas
                                height: parent.height - chipLegend.height - 12
                                anchors.left: parent.left
                                anchors.top: chipLegend.bottom
                                anchors.leftMargin: 8
                                anchors.topMargin: 4
                                contentWidth: mesesFlow.width
                                flickableDirection: Flickable.HorizontalFlick
                                clip: true

                                Flow {
                                    id: mesesFlow
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
                    id: chartContainer
                    width: Math.min(parent.width - 160, 1000)
                    height: Math.min(Math.max(400, backend.nomes && backend.nomes.length > 0 ? backend.nomes.length * 38 : 400), 730)
                    radius: 22
                    border.color: "#3cb3e6"
                    border.width: 2
                    anchors.horizontalCenter: parent.horizontalCenter

                    Loader {
                        id: chartLoader
                        anchors.fill: parent
                        sourceComponent: chartComponent
                        property bool updateTrigger: false

                        Connections {
                            target: backend
                            function onValoresChanged() {
                                chartLoader.updateTrigger = !chartLoader.updateTrigger
                                chartLoader.active = false
                                chartLoader.active = true
                            }
                            function onNomesChanged() {
                                chartLoader.updateTrigger = !chartLoader.updateTrigger
                                chartLoader.active = false
                                chartLoader.active = true
                            }
                        }
                    }

                    Component {
                        id: chartComponent
                        ChartView {
                            anchors.fill: parent
                            antialiasing: true
                            backgroundColor: "#fff"
                            legend.visible: false
                            animationOptions: ChartView.SeriesAnimations
                            animationDuration: 900

                            Component.onCompleted: {
                                console.log("Forçando geração da tabela semana...");
                                if (backend) {
                                    backend.gerarTabelaSemana();
                                }
                            }

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
                                    values: backend.valores && backend.valores.length > 0 ? backend.valores : []
                                    brushFilename: "assets/gradients/bargradient.png"
                                    labelFont: Qt.font({ pixelSize: 16, bold: true, family: "Arial" })
                                    labelColor: "#1976d2"
                                }
                            }

                            BarCategoryAxis {
                                id: eixoUsuarios
                                categories: backend.nomes && backend.nomes.length > 0 ? backend.nomes : [" "]
                                labelsFont.pixelSize: 14
                                labelsFont.bold: true
                                labelsColor: "#232946"
                            }

                            ValueAxis {
                                id: eixoValores
                                min: 0
                                max: backend.valores && backend.valores.length > 0 ? Math.max.apply(Math, backend.valores) * 1.1 : 1
                                labelFormat: "%.0f"
                            }

                            Text {
                                anchors.centerIn: parent
                                visible: !backend.valores || backend.valores.length === 0
                                text: "Nenhum dado disponível para os meses selecionados."
                                color: "#3cb3e6"
                                font.pixelSize: 20
                                font.bold: true
                            }

                            title: "Minutas por Usuário"
                            titleFont: Qt.font({ pixelSize: 12, bold: true, family: "Arial" })
                            titleColor: "#1976d2"
                        }
                    }
                }
            }
        }
    }

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


                                                    // CAMPO DE DIGITAÇÃO CORRIGIDO
                                                    TextField {
                                                        id: textField
                                                        property real valorAtual: model.peso !== undefined ? Number(model.peso) : 1.0
                                                        
                                                        text: valorAtual.toFixed(1)
                                                        Layout.preferredWidth: 70
                                                        height: 32
                                                        font.pixelSize: 16
                                                        horizontalAlignment: TextInput.AlignHCenter
                                                        verticalAlignment: TextInput.AlignVCenter
                                                        validator: DoubleValidator { 
                                                            bottom: 0.1; 
                                                            top: 10.0;
                                                            decimals: 1;
                                                            notation: DoubleValidator.StandardNotation
                                                        }
                                                        inputMethodHints: Qt.ImhFormattedNumbersOnly
                                                        
                                                        background: Rectangle {
                                                            radius: 6
                                                            border.color: "#3cb3e6"
                                                            border.width: 1
                                                            color: "#fff"
                                                        }
                                                        
                                                        // SOLUÇÃO: Guardar o valor no foco e só atualizar se mudou
                                                        property real valorNoFoco: 0.0
                                                        
                                                        onFocusChanged: {
                                                            if (focus) {
                                                                valorNoFoco = valorAtual
                                                            } else {
                                                                // Quando perde o foco, verifica se o valor realmente mudou
                                                                let valorDigitado = parseFloat(text)
                                                                if (isNaN(valorDigitado)) {
                                                                    // Restaurar valor original se inválido
                                                                    text = valorNoFoco.toFixed(1)
                                                                } else {
                                                                    valorDigitado = Math.min(10, Math.max(0.1, valorDigitado))
                                                                    // Só atualiza se realmente houve mudança
                                                                    if (Math.abs(valorDigitado - valorNoFoco) > 0.01) {
                                                                        text = valorDigitado.toFixed(1)
                                                                        pesosModel.setProperty(index, "peso", valorDigitado)
                                                                        backend.atualizarPeso(index, valorDigitado)
                                                                        valorAtual = valorDigitado
                                                                    }
                                                                }
                                                            }
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

        // Página semanaPage otimizada com alinhamentos, rolagem e ordenação
    Component {
        id: semanaPage
    
        Rectangle {
            color: "transparent"
            Layout.fillWidth: true
            Layout.fillHeight: true
    
            // Propriedades para ordenação
            property bool sortAscending: true
            property string sortColumn: "usuario"
    
            // Modelo de dados para o ranking
            ListModel {
                id: rankingModel
            }
            
            // Função para atualizar o modelo de ranking
            function updateRankingModel() {
                rankingModel.clear();
                if (backend && backend.rankingSemana) {
                    var lines = backend.rankingSemana.split('\n');
                    for (var i = 0; i < lines.length; i++) {
                        if (lines[i].trim().length > 0) {
                            rankingModel.append({"rankText": lines[i]});
                        }
                    }
                }
            }
            
            // Função para ordenar a tabela
            function sortTable(column) {
                if (sortColumn === column) {
                    sortAscending = !sortAscending;
                } else {
                    sortColumn = column;
                    sortAscending = true;
                }
                
                if (backend) {
                    backend.ordenarTabelaSemana(sortColumn, sortAscending);
                }
            }
            
            Component.onCompleted: {
                updateRankingModel();
            }
            
            Connections {
                target: backend
                function onRankingSemanaChanged() {
                    updateRankingModel();
                }
            }
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
                spacing: 16
    
                Text {
                    text: "Comparação de Produtividade por Dia da Semana"
                    font.pixelSize: 28
                    font.bold: true
                    color: "#3cb3e6"
                    horizontalAlignment: Text.AlignHCenter
                    Layout.alignment: Qt.AlignHCenter
                    Layout.topMargin: 8
                    Layout.bottomMargin: 4
                }
    
                // Container principal - usa mais espaço da tela
                Rectangle {
                    id: mainContainer
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    Layout.margins: 10
                    color: "#fff"
                    border.color: "#3cb3e6"
                    border.width: 2
                    radius: 16
                    clip: true
                    Layout.preferredWidth: Math.min(parent.width - 40, 1200)
                    Layout.alignment: Qt.AlignHCenter
    
                    // Mensagem para quando não houver dados
                    Text {
                        anchors.centerIn: parent
                        visible: !(backend && backend.tabelaSemana && backend.tabelaSemana.length > 0)
                        text: "Nenhum dado disponível para os meses selecionados."
                        font.pixelSize: 18
                        color: "#3cb3e6"
                        font.bold: true
                    }
    
                    // Layout da tabela
                    Column {
                        id: tableLayout
                        anchors.fill: parent
                        anchors.bottomMargin: 100 // Espaço para o painel inferior
                        spacing: 0
                        visible: backend && backend.tabelaSemana && backend.tabelaSemana.length > 0
    
                        // Cabeçalho da tabela fixo
                        Rectangle {
                            id: headerRect
                            width: parent.width
                            height: 50
                            color: "#e3f2fd"
                            z: 10 // Garantir que fique acima durante a rolagem
    
                            Row {
                                anchors.fill: parent
    
                                // Coluna Usuário
                                Rectangle {
                                    width: 220
                                    height: parent.height
                                    color: "transparent"
    
                                    RowLayout {
                                        anchors.centerIn: parent
                                        spacing: 5
    
                                        Text {
                                            text: "Usuário"
                                            font.pixelSize: 18
                                            font.bold: true
                                            color: "#1976d2"
                                        }
                                        
                                        // Ícone de ordenação
                                        Text {
                                            text: sortColumn === "usuario" ? (sortAscending ? "▲" : "▼") : ""
                                            font.pixelSize: 14
                                            color: "#3cb3e6"
                                            font.bold: true
                                        }
                                    }
                                    
                                    // Área clicável para ordenação
                                    MouseArea {
                                        anchors.fill: parent
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: sortTable("usuario")
                                    }
                                }
    
                                // Colunas para dias da semana
                                Repeater {
                                    model: [
                                        {name: "Segunda", key: "Segunda"},
                                        {name: "Terça", key: "Terça"},
                                        {name: "Quarta", key: "Quarta"},
                                        {name: "Quinta", key: "Quinta"},
                                        {name: "Sexta", key: "Sexta"},
                                        {name: "Sábado", key: "Sábado"},
                                        {name: "Domingo", key: "Domingo"}
                                    ]
                                    
                                    Rectangle {
                                        width: (parent.width - 220) / 7
                                        height: parent.height
                                        color: "transparent"
                                        
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 5
                                
                                            Text {
                                                text: modelData.name
                                                font.pixelSize: 16
                                                font.bold: true
                                                color: "#1976d2"
                                            }
                                            
                                            // Ícone de ordenação
                                            Text {
                                                text: sortColumn === modelData.key ? (sortAscending ? "▲" : "▼") : ""
                                                font.pixelSize: 14
                                                color: "#3cb3e6"
                                                font.bold: true
                                            }
                                        }
                                        
                                        // Área clicável para ordenação
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: sortTable(modelData.key)
                                        }
                                    }
                                }
                            }
    
                            // Linha separadora
                            Rectangle {
                                anchors.bottom: parent.bottom
                                width: parent.width
                                height: 2
                                color: "#3cb3e6"
                            }
                        }
    
                        // ScrollView MELHORADO para os dados da tabela
                        ScrollView {
                            id: tableScrollView
                            width: parent.width
                            height: parent.height - headerRect.height
                            clip: true
                            ScrollBar.horizontal.policy: ScrollBar.AlwaysOff
                            ScrollBar.vertical.policy: ScrollBar.AsNeeded
                            contentWidth: width
                        
                            Column {
                                id: tableContent
                                width: tableScrollView.width
                                spacing: 0
                        
                                Repeater {
                                    model: backend && backend.tabelaSemana ? backend.tabelaSemana : []

                                    Rectangle {
                                        id: userRowItem
                                        width: tableContent.width
                                        height: 40
                                        color: index % 2 === 0 ? "#f5faff" : "#ffffff" // Cores alternadas
                                        
                                        Component.onCompleted: {
                                            console.log("userRowItem modelData:", JSON.stringify(modelData))
                                        }
                                        
                                        // Efeito de hover como primeira camada
                                        Rectangle {
                                            id: hoverHighlight
                                            anchors.fill: parent
                                            color: "#2196f3"
                                            opacity: 0
                                            Behavior on opacity { NumberAnimation { duration: 100 } }
                                            z: 1
                                        }
                                        
                                        // MouseArea FORA da Row, mas DENTRO do Rectangle pai
                                        MouseArea {
                                            id: rowHoverArea
                                            anchors.fill: parent
                                            hoverEnabled: true
                                            z: 100
                                            onEntered: hoverHighlight.opacity = 0.15
                                            onExited: hoverHighlight.opacity = 0
                                            propagateComposedEvents: false
                                        }
                                        
                                        // Row principal com o conteúdo
                                        Row {
                                            z: 2
                                            width: parent.width
                                            height: parent.height
                                            
                                            // Coluna Usuário
                                            Rectangle {
                                                width: 220
                                                height: parent.height
                                                color: "transparent"
                                    
                                                Text {
                                                    anchors.verticalCenter: parent.verticalCenter
                                                    anchors.left: parent.left
                                                    anchors.leftMargin: 16
                                                    text: modelData.usuario
                                                    font.pixelSize: 15
                                                    font.bold: true
                                                    color: "#232946"
                                                    elide: Text.ElideRight
                                                    width: 200
                                                }
                                            }

                                            
                                            Repeater {
                                                model: ["segunda", "terca", "quarta", "quinta", "sexta", "sabado", "domingo"]
                                                Rectangle {
                                                    id: dayColumnDelegate
                                                    width: (parent.width - 220) / 7
                                                    height: 40
                                                    color: "transparent"
                                                    
                                                    Rectangle {
                                                        anchors.centerIn: parent
                                                        width: 70
                                                        height: 28
                                                        radius: 6
                                                        
                                                        property var itemData: parent.parent.parent.parent.modelData
                                                        property string currentDay: dayColumnDelegate.modelData
                                                        
                                                        property real dayValue: {
                                                            if (!itemData) return 0.0;
                                                            return Number(itemData[currentDay] || 0)
                                                        }
                                                        
                                                        color: dayValue > 50 ? "#a3d4ff" : dayValue > 20 ? "#d9eaf7" : "#f7f7f7"
                                                        border.width: 1
                                                        border.color: dayValue > 50 ? "#1976d2" : dayValue > 20 ? "#a0c8e4" : "#e0e0e0"
                                                        
                                                        Text {
                                                            anchors.centerIn: parent
                                                            text: parent.dayValue.toFixed(1)
                                                            font.pixelSize: 15
                                                            color: parent.dayValue > 0 ? "#232946" : "#9e9e9e"
                                                            font.bold: parent.dayValue > 50
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // Barra de rolagem personalizada com tom azul
                            ScrollBar {
                                id: vScrollBar
                                orientation: Qt.Vertical
                                size: 0.3
                                position: 0.2
                                active: true
                                parent: tableScrollView
                                anchors.top: tableScrollView.top
                                anchors.right: tableScrollView.right
                                anchors.bottom: tableScrollView.bottom
                                policy: ScrollBar.AsNeeded
                                visible: tableScrollView.contentHeight > tableScrollView.height
                                width: 12
                                
                                contentItem: Rectangle {
                                    implicitWidth: 12
                                    radius: 6
                                    color: vScrollBar.pressed ? "#1976d2" : vScrollBar.hovered ? "#3cb3e6" : "#98cdf0"
                                }
                                
                                background: Rectangle {
                                    implicitWidth: 12
                                    radius: 6
                                    color: "#e3f2fd"
                                }
                            }
                        }
                    }
    
                    // Container inferior para ranking e botão
                    Rectangle {
                        id: bottomPanel
                        anchors.bottom: parent.bottom
                        anchors.left: parent.left
                        anchors.right: parent.right
                        height: 100
                        color: "transparent"
    
                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 10
                            spacing: 20
    
                            // Painel de Ranking - visual otimizado
                            Rectangle {
                                Layout.fillWidth: true
                                Layout.preferredHeight: 80
                                radius: 12
                                color: "#f5faff"
                                border.color: "#3cb3e6"
                                border.width: 1
    
                                // Título do ranking
                                Text {
                                    id: rankingTitle
                                    text: "Produtividade geral por dia da semana"
                                    font.pixelSize: 14
                                    font.bold: true
                                    color: "#3cb3e6"
                                    anchors {
                                        top: parent.top
                                        topMargin: 8
                                        horizontalCenter: parent.horizontalCenter
                                    }
                                }
    
                                // ScrollView horizontal para o ranking
                                ScrollView {
                                    id: rankingScroll
                                    anchors {
                                        top: rankingTitle.bottom
                                        topMargin: 4
                                        bottom: parent.bottom
                                        bottomMargin: 8
                                        left: parent.left
                                        right: parent.right
                                        leftMargin: 16
                                        rightMargin: 16
                                    }
                                    clip: true
                                    ScrollBar.horizontal.policy: ScrollBar.AlwaysOff
                                    ScrollBar.vertical.policy: ScrollBar.AlwaysOff
    
                                    ScrollBar {
                                        id: hScrollBar
                                        orientation: Qt.Horizontal
                                        size: 0.3
                                        position: 0.0
                                        active: true
                                        policy: ScrollBar.AsNeeded
                                        parent: rankingScroll
                                        anchors.left: rankingScroll.left
                                        anchors.right: rankingScroll.right
                                        anchors.bottom: rankingScroll.bottom
                                        height: 8
                                        visible: rankingRow.width > rankingScroll.width
                                        
                                        contentItem: Rectangle {
                                            implicitHeight: 8
                                            radius: 4
                                            color: hScrollBar.pressed ? "#1976d2" : hScrollBar.hovered ? "#3cb3e6" : "#98cdf0"
                                        }
    
                                        background: Rectangle {
                                            implicitHeight: 8
                                            radius: 4
                                            color: "#e3f2fd"
                                        }
                                    }
                                    
                                    Row {
                                        id: rankingRow
                                        spacing: 16
                                        anchors.verticalCenter: parent.verticalCenter
    
                                        // Repeater usando o modelo do ListModel
                                        Repeater {
                                            model: rankingModel
    
                                            Rectangle {
                                                height: 40
                                                width: rankItemText.width + 20
                                                color: index < 3 ? "#e3f2fd" : "transparent"
                                                radius: 8
                                                border.width: 1
                                                border.color: index < 3 ? "#3cb3e6" : "#e0e0e0"
                                                
                                                Text {
                                                    id: rankItemText
                                                    anchors.centerIn: parent
                                                    text: model.rankText
                                                    font.pixelSize: 14
                                                    color: "#232946"
                                                    font.bold: index < 3
                                                    padding: 5
                                                }
                                            }
                                        }
                                    }
                                }
                            }
    
                            // Botão Atualizar Dados - corrigido usando o mesmo padrão dos outros botões
                            Rectangle {
                                Layout.preferredWidth: 180
                                Layout.preferredHeight: 50
                                radius: 12
                                gradient: Gradient {
                                    GradientStop { position: 0.0; color: "#3cb3e6" }
                                    GradientStop { position: 1.0; color: "#1976d2" }
                                }
                                
                                // Adicionar efeito de hover
                                property bool hovered: false
                                scale: hovered ? 1.03 : 1.0
                                Behavior on scale { NumberAnimation { duration: 120 } }
                            
                                MouseArea {
                                    anchors.fill: parent
                                    hoverEnabled: true
                                    cursorShape: Qt.PointingHandCursor
                                    onEntered: parent.hovered = true
                                    onExited: parent.hovered = false
                                    onClicked: {
                                        if (backend) {
                                            console.log("Gerando tabela da semana...");
                                            backend.gerarTabelaSemana();
                                            updateRankingModel();
                                        }
                                    }
                                }
                            
                                // Use RowLayout como nos outros botões
                                RowLayout {
                                    anchors.centerIn: parent
                                    spacing: 8
                            
                                    Image {
                                        source: "assets/icons/peso4.png"
                                        Layout.preferredWidth: 24
                                        Layout.preferredHeight: 24
                                        fillMode: Image.PreserveAspectFit
                                        Layout.alignment: Qt.AlignVCenter
                                    }

                                    Text {
                                        text: "Atualizar Dados"
                                        color: "#fff"
                                        font.bold: true
                                        font.pixelSize: 15
                                        verticalAlignment: Text.AlignVCenter
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}