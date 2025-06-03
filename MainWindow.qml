import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Effects
import QtCharts
import QtQuick.Window

Window {
    width: 1440
    height: 900
    visible: true
    visibility: Window.Maximized

    property string currentPage: "dashboard"
    property bool temaEscuro: backend ? backend.temaEscuro : true

    property var coresClaro: {
        "background": "#f9fafb",
        "backgroundSecundario": "#e5e7ea",
        "cardBackground": "#fff",
        "bordaCard": "#3cb3e6",
        "textoTitulo": "#1976d2",
        "textoPrimario": "#232946",
        "textoSecundario": "#666666",
        "acento": "#3cb3e6",
        "hover": "#e3f2fd",
        "sidebar": "#fff",
        "bordaSidebar": "#e0e0e0",
        "fundoInput": "#ffffff",
        "textoInput": "#232946",
        "bordaInput": "#e0e0e0",
        "fundoTabela": "#ffffff",
        "linhaTabela": "#f8f9fa",
        "headerTabela": "#e3f2fd",
        "fundoPopup": "#ffffff",
        "fundoGrafico": "#ffffff"
    }
    
    property var coresEscuro: {
        "background": "#1a1a1a",
        "backgroundSecundario": "#2d2d2d",
        "cardBackground": "#2d2d2d",
        "bordaCard": "#4a90e2",
        "textoTitulo": "#4a90e2",
        "textoPrimario": "#ffffff",
        "textoSecundario": "#b0b0b0",
        "acento": "#4a90e2",
        "hover": "#3d3d3d",
        "sidebar": "#262626",
        "bordaSidebar": "#404040",
        "fundoInput": "#3d3d3d",
        "textoInput": "#ffffff",
        "bordaInput": "#404040",
        "fundoTabela": "#2d2d2d",
        "linhaTabela": "#353535",
        "headerTabela": "#3d3d3d",
        "fundoPopup": "#2d2d2d",
        "fundoGrafico": "#2d2d2d"
    }

    property var cores: temaEscuro ? coresEscuro : coresClaro

    Connections {
        target: backend
        function onTemaChanged() {
            temaEscuro = backend.temaEscuro
        }
    }

    Rectangle {
        anchors.fill: parent
        gradient: Gradient {
            GradientStop { position: 0.0; color: cores.background }
            GradientStop { position: 1.0; color: cores.backgroundSecundario }
        }

        RowLayout {
            anchors.fill: parent
            spacing: 0

            Rectangle {
                id: sidebar
                width: hovered ? 220 : 110
                color: cores.sidebar
                border.color: cores.bordaSidebar
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
                    anchors.margins: 10
                    spacing: 10
                
                    Item { Layout.fillHeight: true }
                
                    Repeater {
                        model: [
                            { icon: "dashboard.png", label: "Dashboard", page: "dashboard" },
                            { icon: "compara.png", label: "Comparar", page: "comparar" },
                            { icon: "semana.png", label: "Semana", page: "semana" },
                            { icon: "peso.png", label: "Peso", page: "pesos" },
                            { icon: "settings.png", label: "Configurações", page: "config" }
                        ]
                        
                        Rectangle {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 70
                            Layout.topMargin: 15
                            Layout.bottomMargin: 15
                            color: currentPage === modelData.page ? cores.hover : "transparent"
                            radius: 12
                            border.width: currentPage === modelData.page ? 2 : 0
                            border.color: cores.acento
                            
                            MouseArea {
                                anchors.fill: parent
                                hoverEnabled: true
                                cursorShape: Qt.PointingHandCursor
                                onClicked: currentPage = modelData.page
                            }
                            
                            RowLayout {
                                anchors.fill: parent
                                anchors.leftMargin: 20
                                anchors.rightMargin: 20
                                spacing: 15
                
                                Image {
                                    source: modelData.icon ? "assets/icons/" + modelData.icon : ""
                                    Layout.preferredWidth: 32
                                    Layout.preferredHeight: 32
                                    fillMode: Image.PreserveAspectFit
                                    Layout.alignment: Qt.AlignCenter
                                }
                                
                                Text {
                                    text: modelData.label ? modelData.label : ""
                                    visible: sidebar.hovered
                                    opacity: sidebar.hovered ? 1 : 0
                                    font.pixelSize: 16
                                    font.weight: Font.Medium
                                    color: currentPage === modelData.page ? cores.textoTitulo : cores.textoPrimario
                                    Layout.fillWidth: true
                                    Layout.alignment: Qt.AlignVCenter
                                    Behavior on opacity { NumberAnimation { duration: 200 } }
                                }
                            }
                        }
                    }

                    Item { Layout.fillHeight: true }
                }
            }

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
                        if (backend) {
                            backend.handle_page_loaded(currentPage)
                            if (currentPage === "dashboard") {
                                Qt.callLater(function() {
                                    backend.atualizar_grafico()
                                })
                            }
                        }
                    }
                    
                    Connections {
                        target: backend
                        function onChartUpdateRequired() {
                            if (currentPage === "dashboard") {
                                if (pageLoader.item && pageLoader.item.children.length > 0) {
                                    var chartView = pageLoader.item.children[0].children[2]
                                    if (chartView) {
                                        Qt.callLater(function() {
                                            backend.atualizar_grafico()
                                        })
                                    }
                                }
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
            color: cores.fundoPopup
            radius: 8
            border.color: cores.bordaCard
            border.width: 2
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
    
                Text {
                    text: "Pesos salvos com sucesso!"
                    font.pixelSize: 18
                    color: cores.acento
                    Layout.alignment: Qt.AlignHCenter
                }
    
                Text {
                    text: "Arquivo disponível na pasta:\n'Documentos/Analyzer-Dev-APROD/pesos salvos'"
                    font.pixelSize: 14
                    color: cores.textoPrimario
                    wrapMode: Text.WordWrap
                    Layout.fillWidth: true
                }
            }
        }
    }
    
    Popup {
        id: restoreConfirmation
        width: 400
        height: 160
        modal: true
        focus: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent
    
        Rectangle {
            anchors.fill: parent
            color: cores.fundoPopup
            radius: 8
            border.color: cores.bordaCard
            border.width: 2
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
                spacing: 20
    
                Text {
                    text: "Você tem certeza que deseja restaurar todos os pesos para 1?"
                    font.pixelSize: 16
                    color: cores.textoPrimario
                    wrapMode: Text.WordWrap
                    Layout.fillWidth: true
                    horizontalAlignment: Text.AlignHCenter
                }
    
                RowLayout {
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 20
    
                    Rectangle {
                        width: 100
                        height: 40
                        radius: 8
                        color: cores.acento
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                restoreConfirmation.close()
                                if (backend) {
                                    backend.restore_pesos_complete()
                                }
                            }
                        }
    
                        Text {
                            anchors.centerIn: parent
                            text: "Sim"
                            color: "#fff"
                            font.bold: true
                            font.pixelSize: 14
                        }
                    }
    
                    Rectangle {
                        width: 100
                        height: 40
                        radius: 8
                        color: cores.textoSecundario
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: restoreConfirmation.close()
                        }
    
                        Text {
                            anchors.centerIn: parent
                            text: "Cancelar"
                            color: cores.textoPrimario
                            font.bold: true
                            font.pixelSize: 14
                        }
                    }
                }
            }
        }
    }
    
    Popup {
        id: exportPopup
        width: 420
        height: 280
        modal: true
        focus: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent
    
        Rectangle {
            anchors.fill: parent
            color: cores.fundoPopup
            radius: 12
            border.color: cores.bordaCard
            border.width: 2
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 25
                spacing: 25
    
                Text {
                    text: "Selecione o formato para exportação:"
                    font.pixelSize: 18
                    color: cores.textoPrimario
                    font.bold: true
                    Layout.alignment: Qt.AlignHCenter
                }
    
                ColumnLayout {
                    Layout.alignment: Qt.AlignHCenter
                    Layout.fillWidth: true
                    spacing: 15
    
                    Rectangle {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 50
                        radius: 10
                        color: cores.acento
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                exportPopup.close()
                                if (backend) {
                                    backend.exportarTabelaSemana("csv")
                                }
                            }
                        }
    
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 12
    
                            Image {
                                source: "assets/icons/excel.png"
                                Layout.preferredWidth: 28
                                Layout.preferredHeight: 28
                                fillMode: Image.PreserveAspectFit
                            }
    
                            Text {
                                text: "Exportar como CSV"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 16
                            }
                        }
                    }
    
                    Rectangle {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 50
                        radius: 10
                        color: "#4caf50"
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                exportPopup.close()
                                if (backend) {
                                    backend.exportarTabelaSemana("xlsx")
                                }
                            }
                        }
    
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 12
    
                            Image {
                                source: "assets/icons/excel.png"
                                Layout.preferredWidth: 28
                                Layout.preferredHeight: 28
                                fillMode: Image.PreserveAspectFit
                            }
    
                            Text {
                                text: "Exportar como XLSX"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 16
                            }
                        }
                    }
    
                    Rectangle {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 50
                        radius: 10
                        color: "#ff5722"
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                exportPopup.close()
                                if (backend) {
                                    backend.exportarTabelaSemana("pdf")
                                }
                            }
                        }
    
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 12
    
                            Image {
                                source: "assets/icons/excel.png"
                                Layout.preferredWidth: 28
                                Layout.preferredHeight: 28
                                fillMode: Image.PreserveAspectFit
                            }
    
                            Text {
                                text: "Exportar como PDF"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 16
                            }
                        }
                    }
                }
            }
        }
    }
    
    Popup {
        id: exportSuccessPopup
        width: 450
        height: 200
        modal: true
        focus: true
        closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
        anchors.centerIn: parent
    
        Connections {
            target: backend
            function onExportSuccessSignal() {
                exportSuccessPopup.open()
            }
        }
    
        Rectangle {
            anchors.fill: parent
            color: cores.fundoPopup
            radius: 12
            border.color: "#4caf50"
            border.width: 2
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 25
                spacing: 20
    
                Text {
                    text: "Sucesso ao exportar tabela!"
                    font.pixelSize: 18
                    color: "#4caf50"
                    font.bold: true
                    Layout.alignment: Qt.AlignHCenter
                }
    
                Text {
                    text: "Arquivo salvo em: exports/tabela_semana"
                    font.pixelSize: 14
                    color: cores.textoPrimario
                    Layout.alignment: Qt.AlignHCenter
                    horizontalAlignment: Text.AlignHCenter
                }
    
                RowLayout {
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 15
    
                    Rectangle {
                        width: 120
                        height: 40
                        radius: 8
                        color: "#4caf50"
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                if (backend) {
                                    backend.abrirPastaExports()
                                }
                                exportSuccessPopup.close()
                            }
                        }
    
                        Text {
                            anchors.centerIn: parent
                            text: "Abrir Pasta"
                            color: "#fff"
                            font.bold: true
                            font.pixelSize: 14
                        }
                    }
    
                    Rectangle {
                        width: 80
                        height: 40
                        radius: 8
                        color: cores.textoSecundario
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: exportSuccessPopup.close()
                        }
    
                        Text {
                            anchors.centerIn: parent
                            text: "Fechar"
                            color: cores.textoPrimario
                            font.bold: true
                            font.pixelSize: 14
                        }
                    }
                }
            }
        }
    }

    Component {
        id: dashboardPage
    
        Rectangle {
            color: "transparent"
            anchors.fill: parent
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
                spacing: 20
    
                RowLayout {
                    Layout.fillWidth: true
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 20
                
                    Rectangle {
                        Layout.preferredWidth: 280
                        Layout.preferredHeight: 120
                        radius: 16
                        color: cores.cardBackground
                        border.color: cores.bordaCard
                        border.width: 2
                
                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 16
                            spacing: 12
                
                            Image {
                                source: "assets/icons/sum.png"
                                Layout.preferredWidth: 30
                                Layout.preferredHeight: 30
                                fillMode: Image.PreserveAspectFit
                                Layout.alignment: Qt.AlignVCenter
                            }
                
                            Column {
                                Layout.alignment: Qt.AlignVCenter
                                spacing: 4
                                Text {
                                    text: "Total de Minutas"
                                    font.pixelSize: 16
                                    color: cores.acento
                                    font.bold: true
                                }
                                Text {
                                    text: kpis ? kpis["minutas"] : "0"
                                    font.pixelSize: 28
                                    color: cores.textoPrimario
                                    font.bold: true
                                }
                            }
                        }
                    }
                
                    Rectangle {
                        Layout.preferredWidth: 280
                        Layout.preferredHeight: 120
                        radius: 16
                        color: cores.cardBackground
                        border.color: cores.bordaCard
                        border.width: 2
                    
                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 16
                            spacing: 12
                    
                            Image {
                                source: "assets/icons/calendar.png"
                                Layout.preferredWidth: 30
                                Layout.preferredHeight: 30
                                fillMode: Image.PreserveAspectFit
                                Layout.alignment: Qt.AlignVCenter
                            }
                    
                            Column {
                                Layout.alignment: Qt.AlignVCenter
                                Layout.fillWidth: true
                                spacing: 4
                                Text {
                                    text: "Dia Mais Produtivo"
                                    font.pixelSize: 16
                                    color: cores.acento
                                    font.bold: true
                                }
                                Text {
                                    text: backend && backend.kpis ? backend.kpis["dia_produtivo"] : "--"
                                    font.pixelSize: backend && backend.kpis && backend.kpis["dia_produtivo"] && backend.kpis["dia_produtivo"].includes("\n") ? 20 : 28
                                    color: cores.textoPrimario
                                    font.bold: true
                                    wrapMode: Text.WordWrap
                                    width: parent.width
                                    horizontalAlignment: Text.AlignLeft
                                }
                            }
                        }
                    }
                
                    Rectangle {
                        Layout.preferredWidth: 320
                        Layout.preferredHeight: 120
                        radius: 16
                        color: cores.cardBackground
                        border.color: cores.bordaCard
                        border.width: 2
                
                        RowLayout {
                            anchors.fill: parent
                            anchors.margins: 16
                            spacing: 12
                
                            Image {
                                source: "assets/icons/top.png"
                                Layout.preferredWidth: 30
                                Layout.preferredHeight: 30
                                fillMode: Image.PreserveAspectFit
                                Layout.alignment: Qt.AlignVCenter
                            }
                
                            Column {
                                Layout.alignment: Qt.AlignVCenter
                                Layout.fillWidth: true
                                spacing: 4
                
                                Text {
                                    text: "Top 3 Usuários"
                                    font.pixelSize: 16
                                    color: cores.acento
                                    font.bold: true
                                }
                
                                Column {
                                    spacing: 2
                                    
                                    Repeater {
                                        model: backend && backend.kpis && backend.kpis["top3"] ? backend.kpis["top3"].split(" | ") : []
                                        
                                        Text {
                                            text: modelData
                                            font.pixelSize: 12
                                            color: cores.textoPrimario
                                            font.bold: index === 0
                                            wrapMode: Text.WordWrap
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
    
                RowLayout {
                    Layout.fillWidth: true
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 20
    
                    Rectangle {
                        Layout.preferredWidth: 200
                        Layout.preferredHeight: 50
                        Layout.alignment: Qt.AlignLeft
                        radius: 12
                        gradient: Gradient {
                            GradientStop { position: 0.0; color: cores.acento }
                            GradientStop { position: 1.0; color: cores.textoTitulo }
                        }
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: if (backend) backend.importar_arquivo_excel()
                        }
    
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 8
    
                            Image {
                                source: "assets/icons/excel.png"
                                Layout.preferredWidth: 24
                                Layout.preferredHeight: 24
                                fillMode: Image.PreserveAspectFit
                            }
    
                            Text {
                                text: "Importar Excel"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 16
                            }
                        }
                    }
    
                    Rectangle {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 50
                        Layout.alignment: Qt.AlignCenter
                        radius: 12
                        color: cores.cardBackground
                        border.color: cores.bordaCard
                        border.width: 2
    
                        Row {
                            anchors.fill: parent
                            anchors.margins: 16
                            spacing: 12
                            anchors.verticalCenter: parent.verticalCenter
    
                            Text {
                                text: "Meses Ativos:"
                                font.pixelSize: 16
                                font.bold: true
                                color: cores.acento
                                anchors.verticalCenter: parent.verticalCenter
                            }
    
                            Repeater {
                                model: backend ? backend.mesesDisponiveis : []
    
                                Rectangle {
                                    width: mesText.width + 24
                                    height: 32
                                    radius: 8
                                    anchors.verticalCenter: parent.verticalCenter
    
                                    property bool isActive: backend && backend.mes_ativo_status(modelData)
    
                                    color: isActive ? cores.acento : cores.hover
                                    border.color: isActive ? cores.textoTitulo : cores.bordaCard
                                    border.width: 2
    
                                    opacity: isActive ? 1.0 : 0.5
    
                                    Behavior on color { ColorAnimation { duration: 200 } }
                                    Behavior on border.color { ColorAnimation { duration: 200 } }
                                    Behavior on opacity { NumberAnimation { duration: 200 } }
    
                                    scale: mouseArea.containsMouse ? 1.05 : 1.0
                                    Behavior on scale { NumberAnimation { duration: 100 } }
    
                                    Text {
                                        id: mesText
                                        text: modelData
                                        anchors.centerIn: parent
                                        font.pixelSize: 12
                                        color: parent.isActive ? "#ffffff" : cores.textoSecundario
                                        font.bold: parent.isActive
    
                                        Behavior on color { ColorAnimation { duration: 200 } }
                                    }
    
                                    MouseArea {
                                        id: mouseArea
                                        anchors.fill: parent
                                        hoverEnabled: true
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: {
                                            if (backend) {
                                                backend.toggleMesAtivo(modelData)
                                            }
                                        }
                                    }
    
                                    Connections {
                                        target: backend
                                        function onMesesAtivosChanged() {
                                            isActive = backend.mes_ativo_status(modelData)
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
    
                Rectangle {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    radius: 16
                    color: cores.cardBackground
                    border.color: cores.bordaCard
                    border.width: 2
    
                    ChartView {
                        id: chartView
                        anchors.fill: parent
                        anchors.margins: 10
                        antialiasing: true
                        backgroundColor: cores.fundoGrafico
                        legend.visible: false
                        title: "Minutas por Usuário"
                        titleFont: Qt.font({ pixelSize: 16, bold: true })
                        titleColor: cores.textoTitulo
                        plotAreaColor: cores.fundoGrafico
                        
                        Component.onCompleted: {
                            if (backend) {
                                Qt.callLater(function() {
                                    backend.atualizar_grafico()
                                })
                            }
                        }
                        
                        Connections {
                            target: backend
                            function onNomesChanged() {
                                if (backend && backend.nomes && backend.valores) {
                                    eixoUsuarios.categories = backend.nomes
                                    var barSet = serieMinutas.at(0)
                                    if (barSet) {
                                        barSet.values = backend.valores
                                    }
                                }
                            }
                            function onValoresChanged() {
                                if (backend && backend.nomes && backend.valores) {
                                    var barSet = serieMinutas.at(0)
                                    if (barSet) {
                                        barSet.values = backend.valores
                                    }
                                    eixoValores.max = backend.maxAxisValue
                                }
                            }
                        }
    
                        HorizontalBarSeries {
                            id: serieMinutas
                            axisY: eixoUsuarios
                            axisX: eixoValores
                            barWidth: 0.7
                            labelsVisible: true
                            labelsFormat: "@value"
                            labelsPosition: AbstractBarSeries.LabelsOutsideEnd
    
                            BarSet {
                                label: "Minutas"
                                values: backend && backend.valores ? backend.valores : []
                                color: cores.acento
                                borderColor: cores.acento
                                borderWidth: 0
                                labelColor: cores.textoPrimario
                                labelFont.pixelSize: 12
                            }
                        }
    
                        BarCategoryAxis {
                            id: eixoUsuarios
                            categories: backend && backend.nomes ? backend.nomes : []
                            labelsFont.pixelSize: 12
                            labelsColor: cores.textoPrimario
                            gridVisible: false
                        }
    
                        ValueAxis {
                            id: eixoValores
                            min: 0
                            max: backend ? backend.maxAxisValue : 100
                            tickCount: 5
                            labelFormat: "%d"
                            labelsFont.pixelSize: 12
                            labelsColor: cores.textoPrimario
                            gridLineColor: cores.bordaInput
                            minorGridVisible: false
                        }
    
                        Text {
                            anchors.centerIn: parent
                            visible: !backend || !backend.valores || backend.valores.length === 0
                            text: "Nenhum dado disponível"
                            color: cores.acento
                            font.pixelSize: 18
                            font.bold: true
                        }
                    }
                }
            }
        }
    }

    property bool ordemCrescente: true
    property int sortColumn: 0
    property bool sortAscending: true

    Component {
        id: pesosPage
    
        Rectangle {
            color: "transparent"
            Layout.fillWidth: true
            Layout.fillHeight: true
    
            Connections {
                target: backend
                function onPesosModelUpdate() {
                    console.log("Modelo de pesos atualizado pelo backend")
                }
            }
    
            ColumnLayout {
                anchors.centerIn: parent
                spacing: 24
                width: 1100
    
                Text {
                    text: "Atribuição de Pesos"
                    font.pixelSize: 28
                    font.bold: true
                    color: cores.acento
                    horizontalAlignment: Text.AlignHCenter
                    Layout.alignment: Qt.AlignHCenter
                }
    
                Rectangle {
                    width: 1100
                    height: 600
                    color: cores.cardBackground
                    border.color: cores.bordaCard
                    border.width: 2
                    radius: 18
                    clip: true
                    Layout.alignment: Qt.AlignHCenter
    
                    Column {
                        anchors.fill: parent
                        spacing: 0
    
                        Rectangle {
                            width: parent.width
                            height: 50
                            color: cores.headerTabela
                            radius: 0
                            
                            RowLayout {
                                anchors.fill: parent
                                anchors.margins: 10
                                spacing: 0
    
                                Text {
                                    Layout.preferredWidth: 800
                                    text: "Tipo de Agendamento"
                                    font.pixelSize: 18
                                    font.bold: true
                                    color: cores.textoTitulo
                                    horizontalAlignment: Text.AlignHCenter
                                }
    
                                Text {
                                    Layout.preferredWidth: 260
                                    text: "Peso"
                                    font.pixelSize: 18
                                    font.bold: true
                                    color: cores.textoTitulo
                                    horizontalAlignment: Text.AlignHCenter
                                }
                            }
    
                            Rectangle {
                                anchors.bottom: parent.bottom
                                width: parent.width
                                height: 2
                                color: cores.bordaCard
                            }
                        }
    
                        Flickable {
                            id: flick
                            width: parent.width
                            height: parent.height - 50
                            contentHeight: column.height
                            clip: true
    
                            Column {
                                id: column
                                width: parent.width
    
                                Repeater {
                                    model: backend ? backend.pesosModel : []
                                    
                                    Rectangle {
                                        width: column.width
                                        height: 50
                                        color: index % 2 === 0 ? cores.linhaTabela : cores.cardBackground
                                        
                                        RowLayout {
                                            anchors.fill: parent
                                            anchors.margins: 10
                                            spacing: 0
                                            
                                            Text {
                                                Layout.preferredWidth: 800
                                                text: modelData ? modelData.tipo : ""
                                                font.pixelSize: 16
                                                color: cores.textoPrimario
                                                elide: Text.ElideRight
                                                horizontalAlignment: Text.AlignLeft
                                                verticalAlignment: Text.AlignVCenter
                                                leftPadding: 20
                                            }
                                            
                                            RowLayout {
                                                Layout.preferredWidth: 260
                                                Layout.alignment: Qt.AlignHCenter
                                                spacing: 8
                                                
                                                Rectangle {
                                                    width: 30
                                                    height: 30
                                                    radius: 6
                                                    color: cores.acento
    
                                                    MouseArea {
                                                        anchors.fill: parent
                                                        cursorShape: Qt.PointingHandCursor
                                                        onClicked: {
                                                            if (backend) {
                                                                backend.decrementar_peso(index)
                                                            }
                                                        }
                                                    }
    
                                                    Text {
                                                        anchors.centerIn: parent
                                                        text: "-"
                                                        color: "#fff"
                                                        font.bold: true
                                                        font.pixelSize: 18
                                                    }
                                                }
                                                
                                                Rectangle {
                                                    width: 60
                                                    height: 30
                                                    radius: 6
                                                    color: cores.fundoInput
                                                    border.color: cores.bordaInput
                                                    border.width: 1
    
                                                    TextInput {
                                                        anchors.centerIn: parent
                                                        text: modelData ? modelData.peso.toFixed(1) : "1.0"
                                                        font.pixelSize: 14
                                                        font.bold: true
                                                        color: cores.textoInput
                                                        horizontalAlignment: TextInput.AlignHCenter
                                                        selectByMouse: true
                                                        validator: DoubleValidator {
                                                            bottom: 0.1
                                                            top: 10.0
                                                            decimals: 1
                                                        }
                                                        
                                                        onEditingFinished: {
                                                            if (backend) {
                                                                backend.validar_peso_input(index, text)
                                                            }
                                                        }
                                                    }
                                                }
                                                
                                                Rectangle {
                                                    width: 30
                                                    height: 30
                                                    radius: 6
                                                    color: cores.acento
    
                                                    MouseArea {
                                                        anchors.fill: parent
                                                        cursorShape: Qt.PointingHandCursor
                                                        onClicked: {
                                                            if (backend) {
                                                                backend.incrementar_peso(index)
                                                            }
                                                        }
                                                    }
    
                                                    Text {
                                                        anchors.centerIn: parent
                                                        text: "+"
                                                        color: "#fff"
                                                        font.bold: true
                                                        font.pixelSize: 18
                                                    }
                                                }
                                            }
                                        }
                                        
                                        Rectangle {
                                            anchors.bottom: parent.bottom
                                            width: parent.width
                                            height: 1
                                            color: cores.bordaInput
                                            visible: index < (backend && backend.pesosModel ? backend.pesosModel.length - 1 : 0)
                                        }
                                    }
                                }
                            }
                            
                            ScrollBar.vertical: ScrollBar {
                                width: 12
                                policy: ScrollBar.AlwaysOn
                                contentItem: Rectangle {
                                    radius: 6
                                    color: cores.acento
                                }
                            }
                        }
                    }
                }
    
                RowLayout {
                    Layout.alignment: Qt.AlignHCenter
                    spacing: 18
                
                    Rectangle {
                        width: 180
                        height: 48
                        radius: 14
                        gradient: Gradient {
                            GradientStop { position: 0.0; color: cores.acento }
                            GradientStop { position: 1.0; color: cores.textoTitulo }
                        }
                
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                if (backend) {
                                    backend.salvarPesos()
                                    saveConfirmation.open()
                                }
                            }
                        }
                
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 8
                            Image {
                                source: "assets/icons/peso2.png"
                                Layout.preferredWidth: 24
                                Layout.preferredHeight: 24
                                fillMode: Image.PreserveAspectFit
                            }
                            Text {
                                text: "Salvar Pesos"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 15
                            }
                        }
                    }
                
                    Rectangle {
                        width: 180
                        height: 48
                        radius: 14
                        gradient: Gradient {
                            GradientStop { position: 0.0; color: cores.acento }
                            GradientStop { position: 1.0; color: cores.textoTitulo }
                        }
                
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: {
                                if (backend) {
                                    backend.aplicarPesos()
                                }
                            }
                        }
                
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 8
                            Image {
                                source: "assets/icons/flecha.png"
                                Layout.preferredWidth: 24
                                Layout.preferredHeight: 24
                                fillMode: Image.PreserveAspectFit
                            }
                            Text {
                                text: "Aplicar Pesos"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 15
                            }
                        }
                    }
                
                    Rectangle {
                        width: 180
                        height: 48
                        radius: 14
                        gradient: Gradient {
                            GradientStop { position: 0.0; color: cores.acento }
                            GradientStop { position: 1.0; color: cores.textoTitulo }
                        }
                        
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: restoreConfirmation.open()
                        }
    
                        RowLayout {
                            anchors.centerIn: parent
                            spacing: 8
                            Image {
                                source: "assets/icons/peso4.png"
                                Layout.preferredWidth: 24
                                Layout.preferredHeight: 24
                                fillMode: Image.PreserveAspectFit
                            }
                            Text {
                                text: "Restaurar Padrão"
                                color: "#fff"
                                font.bold: true
                                font.pixelSize: 15
                            }
                        }
                    }
                }
            }
        }
    }


    Component {
        id: semanaPage
    
        Rectangle {
            color: "transparent"
            Layout.fillWidth: true
            Layout.fillHeight: true
    
            property bool sortAscending: backend ? backend.sortAscending : true
            property string sortColumn: backend ? backend.sortColumn : "usuario"
            property string filtroUsuario: ""
            
            Component.onCompleted: {
                if (backend) {
                    backend.gerarTabelaSemana()
                    backend.atualizar_ranking_model()
                }
            }
            
            Connections {
                target: backend
                function onRankingSemanaChanged() {
                    backend.atualizar_ranking_model()
                }
                function onTabelaSemanaChanged() {
                    console.log("Tabela semana atualizada")
                }
                function onSortColumnChanged(column) {
                    sortColumn = column
                }
                function onSortAscendingChanged(ascending) {
                    sortAscending = ascending
                }
                function onRankingModelChanged(dados) {
                    console.log("Ranking model atualizado pelo backend")
                }
            }
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
                spacing: 0
            
                Text {
                    text: "Comparação de Produtividade por Dia da Semana"
                    font.pixelSize: 28
                    font.bold: true
                    color: cores.acento
                    horizontalAlignment: Text.AlignHCenter
                    Layout.alignment: Qt.AlignHCenter
                    Layout.topMargin: 0
                    Layout.bottomMargin: 10
                }
                
                Rectangle {
                    Layout.fillWidth: true
                    Layout.preferredHeight: 60
                    Layout.preferredWidth: Math.min(parent.width - 40, 1200)
                    Layout.alignment: Qt.AlignHCenter
                    color: cores.cardBackground
                    radius: 12
                    border.width: 1
                    border.color: cores.bordaCard
                    
                    RowLayout {
                        anchors.fill: parent
                        anchors.margins: 10
                        spacing: 20
                        
                        Text {
                            text: "Filtrar por usuário:"
                            font.pixelSize: 15
                            color: cores.textoTitulo
                            font.bold: true
                        }
                        
                        Rectangle {
                            Layout.preferredWidth: 250
                            Layout.preferredHeight: 36
                            color: cores.fundoInput
                            radius: 6
                            border.width: 1
                            border.color: cores.bordaCard
                            
                            TextInput {
                                id: filtroInput
                                anchors.fill: parent
                                anchors.margins: 8
                                verticalAlignment: TextInput.AlignVCenter
                                font.pixelSize: 14
                                color: cores.textoInput
                                selectByMouse: true
                                
                                onTextChanged: {
                                    backend.setFiltroUsuario(text)
                                }
                                
                                Text {
                                    anchors.fill: parent
                                    text: "Digite para filtrar..."
                                    font.pixelSize: 14
                                    color: cores.textoSecundario
                                    visible: !parent.text
                                    verticalAlignment: Text.AlignVCenter
                                }
                            }
                        }
                        
                        Item { Layout.fillWidth: true }
                        
                        Rectangle {
                            Layout.preferredWidth: 140
                            Layout.preferredHeight: 36
                            radius: 8
                            gradient: Gradient {
                                GradientStop { position: 0.0; color: cores.acento }
                                GradientStop { position: 1.0; color: cores.textoTitulo }
                            }
                            
                            property bool hovered: false
                            scale: hovered ? 1.03 : 1.0
                            Behavior on scale { NumberAnimation { duration: 120 } }
                            
                            MouseArea {
                                anchors.fill: parent
                                hoverEnabled: true
                                cursorShape: Qt.PointingHandCursor
                                onEntered: parent.hovered = true
                                onExited: parent.hovered = false
                                onClicked: exportPopup.open()
                            }
                            
                            RowLayout {
                                anchors.centerIn: parent
                                spacing: 8
                                
                                Image {
                                    source: "assets/icons/excel.png"
                                    Layout.preferredWidth: 20
                                    Layout.preferredHeight: 20
                                    fillMode: Image.PreserveAspectFit
                                    Layout.alignment: Qt.AlignVCenter
                                }
                                
                                Text {
                                    text: "Exportar"
                                    color: "#fff"
                                    font.bold: true
                                    font.pixelSize: 13
                                    verticalAlignment: Text.AlignVCenter
                                }
                            }
                        }
                    }
                }
                
                Rectangle {
                    id: mainContainer
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    Layout.maximumHeight: parent.height - 300
                    Layout.topMargin: -2
                    color: cores.cardBackground
                    border.color: cores.bordaCard
                    border.width: 2
                    radius: 16
                    clip: true
                    Layout.preferredWidth: Math.min(parent.width - 40, 1200)
                    Layout.alignment: Qt.AlignHCenter
                                
                    Text {
                        anchors.centerIn: parent
                        visible: !(backend && backend.tabelaSemana && backend.tabelaSemana.length > 0)
                        text: "Nenhum dado disponível para os meses selecionados."
                        font.pixelSize: 18
                        color: cores.acento
                        font.bold: true
                    }
                
                    Rectangle {
                        anchors.fill: parent
                        anchors.margins: 2
                        color: "transparent"
                        border.color: cores.bordaCard
                        border.width: 2
                        radius: 14
                        visible: backend && backend.tabelaSemana && backend.tabelaSemana.length > 0
                
                        Column {
                            id: tableLayout
                            anchors.fill: parent
                            anchors.margins: 2
                            spacing: 0
                
                            Rectangle {
                                id: headerRect
                                width: parent.width
                                height: 50
                                color: cores.headerTabela
                                radius: 0
                                border.width: 0
                                z: 10
                            
                                Rectangle {
                                    anchors.bottom: parent.bottom
                                    anchors.left: parent.left
                                    anchors.right: parent.right
                                    height: parent.radius
                                    color: parent.color
                                }
                            
                                Row {
                                    anchors.fill: parent
                                
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
                                                color: cores.textoTitulo
                                            }
                                            
                                            Text {
                                                text: (backend && backend.sortColumn === "usuario") ? (backend.sortAscending ? "▲" : "▼") : ""
                                                font.pixelSize: 14
                                                color: cores.acento
                                                font.bold: true
                                            }
                                        }
                                        
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.atualizar_sort_table("usuario")
                                                }
                                            }
                                        }
                                    }
                            
                                    Repeater {
                                        model: [
                                            {name: "Segunda", key: "segunda"},
                                            {name: "Terça", key: "terca"},
                                            {name: "Quarta", key: "quarta"},
                                            {name: "Quinta", key: "quinta"},
                                            {name: "Sexta", key: "sexta"},
                                            {name: "Sábado", key: "sabado"},
                                            {name: "Domingo", key: "domingo"}
                                        ]
                            
                                        Rectangle {
                                            width: (headerRect.width - 220 - 120) / 7
                                            height: parent.height
                                            color: "transparent"
                                            border.width: 1
                                            border.color: cores.bordaCard
                                            
                                            RowLayout {
                                                anchors.centerIn: parent
                                                spacing: 5
                                        
                                                Text {
                                                    text: modelData.name
                                                    font.pixelSize: 16
                                                    font.bold: true
                                                    color: cores.textoTitulo
                                                }
                                                
                                                Text {
                                                    text: (backend && backend.sortColumn === modelData.key) ? (backend.sortAscending ? "▲" : "▼") : ""
                                                    font.pixelSize: 14
                                                    color: cores.acento
                                                    font.bold: true
                                                }
                                            }
                                            
                                            MouseArea {
                                                anchors.fill: parent
                                                cursorShape: Qt.PointingHandCursor
                                                onClicked: {
                                                    if (backend) {
                                                        backend.atualizar_sort_table(modelData.key)
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    
                                    Rectangle {
                                        width: 120
                                        height: parent.height
                                        color: "transparent"
                                        
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 5
                                            
                                            Text {
                                                text: "Total"
                                                font.pixelSize: 16
                                                font.bold: true
                                                color: cores.textoTitulo
                                            }
                                            
                                            Text {
                                                text: (backend && backend.sortColumn === "Total") ? (backend.sortAscending ? "▲" : "▼") : ""
                                                font.pixelSize: 14
                                                color: cores.acento
                                                font.bold: true
                                            }
                                        }
                                        
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.atualizar_sort_table("Total")
                                                }
                                            }
                                        }
                                    }
                                }
            
                                Rectangle {
                                    anchors.bottom: parent.bottom
                                    width: parent.width
                                    height: 2
                                    color: cores.bordaCard
                                }
                            }
            
                            Flickable {
                                id: tableScrollView
                                width: parent.width
                                height: parent.height - headerRect.height
                                contentHeight: tableContent.height
                                clip: true
                                
                                Column {
                                    id: tableContent
                                    width: tableScrollView.width
                                    spacing: 0
                                
                                    Component.onCompleted: {
                                        console.log("Dados carregados")
                                    }
                                
                                    Repeater {
                                        model: backend && backend.tabelaSemana ? backend.tabelaSemana : []
                                        
                                        Rectangle {
                                            id: userRowItem
                                            width: tableContent.width
                                            height: 40
                                            color: index % 2 === 0 ? cores.linhaTabela : cores.cardBackground
                                            border.width: 0
                                            
                                            property var itemData: modelData
                                            property real rowTotal: backend.calculate_row_total(itemData)
                                        
                                            Row {
                                                anchors.fill: parent
                                        
                                                Rectangle {
                                                    width: 220
                                                    height: parent.height
                                                    color: "transparent"
                                                    
                                                    Text {
                                                        anchors.verticalCenter: parent.verticalCenter
                                                        anchors.left: parent.left
                                                        anchors.leftMargin: 16
                                                        text: itemData ? itemData.usuario : ""
                                                        font.pixelSize: 15
                                                        font.bold: true
                                                        color: cores.textoPrimario
                                                        elide: Text.ElideRight
                                                        width: 200
                                                    }
                                                }
                                
                                                Repeater {
                                                    model: ["segunda", "terca", "quarta", "quinta", "sexta", "sabado", "domingo"]
                                                    
                                                    Rectangle {
                                                        width: (userRowItem.width - 220 - 120) / 7
                                                        height: 40
                                                        color: "transparent"
                                                        border.width: 1
                                                        border.color: cores.bordaInput
                                                        
                                                        Text {
                                                            anchors.centerIn: parent
                                                            text: backend.get_day_value_formatted(userRowItem.itemData, modelData)
                                                            font.pixelSize: 15
                                                            color: cores.textoPrimario
                                                            font.bold: backend.get_day_value(userRowItem.itemData, modelData) > 50
                                                        }
                                                    }
                                                }
                                
                                                Rectangle {
                                                    width: 120
                                                    height: 40
                                                    color: "transparent"
                                                    
                                                    Text {
                                                        anchors.centerIn: parent
                                                        text: backend.get_row_total_formatted(userRowItem.itemData)
                                                        font.pixelSize: 15
                                                        font.bold: backend.get_row_total(userRowItem.itemData) > 200
                                                        color: cores.textoPrimario
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            
                                ScrollBar.vertical: ScrollBar {
                                    width: 12
                                    policy: ScrollBar.AlwaysOn
                                    contentItem: Rectangle {
                                        radius: 6
                                        color: cores.acento
                                    }
                                }
                            }
                        }
                    }
                }
            
                Rectangle {
                    id: bottomPanel
                    Layout.fillWidth: true
                    Layout.preferredHeight: 80
                    Layout.topMargin: 10
                    color: "transparent"
            
                    RowLayout {
                        anchors.fill: parent
                        spacing: 20
            
                        Rectangle {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 80
                            radius: 12
                            color: cores.cardBackground
                            border.color: cores.bordaCard
                            border.width: 1
            
                            Text {
                                id: rankingTitle
                                text: "Produtividade geral por dia da semana"
                                font.pixelSize: 14
                                font.bold: true
                                color: cores.acento
                                anchors {
                                    top: parent.top
                                    topMargin: 2
                                    horizontalCenter: parent.horizontalCenter
                                }
                            }
            
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
            
                                Row {
                                    id: rankingRow
                                    spacing: 16
                                    anchors.verticalCenter: parent.verticalCenter
            
                                    Repeater {
                                        model: backend ? backend.rankingData : []
            
                                        Rectangle {
                                            height: 40
                                            width: rankItemText.width + 20
                                            color: index < 3 ? cores.hover : "transparent"
                                            radius: 8
                                            border.width: 1
                                            border.color: index < 3 ? cores.acento : cores.bordaInput
                                            
                                            Text {
                                                id: rankItemText
                                                anchors.centerIn: parent
                                                text: modelData.rankText
                                                font.pixelSize: 14
                                                color: cores.textoPrimario
                                                font.bold: index < 3
                                                padding: 5
                                            }
                                        }
                                    }
                                }
                            }
                        }
            
                        Rectangle {
                            Layout.preferredWidth: 180
                            Layout.preferredHeight: 50
                            radius: 12
                            gradient: Gradient {
                                GradientStop { position: 0.0; color: cores.acento }
                                GradientStop { position: 1.0; color: cores.textoTitulo }
                            }
                            
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
                                        backend.gerarTabelaSemana()
                                        backend.atualizar_ranking_model()
                                    }
                                }
                            }
                            
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


    Component {
        id: compararPage
    
        Rectangle {
            color: "transparent"
            anchors.fill: parent
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
                spacing: 20
    
                Text {
                    text: "Comparação de Produtividade entre Usuários"
                    font.pixelSize: 28
                    font.bold: true
                    color: cores.acento
                    horizontalAlignment: Text.AlignHCenter
                    Layout.alignment: Qt.AlignHCenter
                }
    
                RowLayout {
                    Layout.fillWidth: true
                    spacing: 20
    
                    Rectangle {
                        Layout.preferredWidth: 400
                        Layout.fillHeight: true
                        color: cores.cardBackground
                        border.color: cores.bordaCard
                        border.width: 2
                        radius: 16
    
                        ColumnLayout {
                            anchors.fill: parent
                            anchors.margins: 15
                            spacing: 10
    
                            Text {
                                text: "Selecionar Usuários para Comparação"
                                font.pixelSize: 18
                                font.bold: true
                                color: cores.textoTitulo
                                Layout.alignment: Qt.AlignHCenter
                            }
    
                            Rectangle {
                                Layout.fillWidth: true
                                Layout.preferredHeight: 35
                                color: cores.fundoInput
                                border.color: cores.bordaCard
                                border.width: 1
                                radius: 8
    
                                TextInput {
                                    id: filtroUsuarios
                                    anchors.fill: parent
                                    anchors.margins: 8
                                    font.pixelSize: 14
                                    color: cores.textoInput
                                    verticalAlignment: TextInput.AlignVCenter
                                    selectByMouse: true
    
                                    onTextChanged: {
                                        if (backend) {
                                            backend.filtrarUsuariosComparacao(text)
                                        }
                                    }
    
                                    Text {
                                        anchors.fill: parent
                                        text: "Filtrar usuários..."
                                        font.pixelSize: 14
                                        color: cores.textoSecundario
                                        visible: !parent.text
                                        verticalAlignment: Text.AlignVCenter
                                    }
                                }
                            }
    
                            ScrollView {
                                Layout.fillWidth: true
                                Layout.fillHeight: true
                                clip: true
    
                                ListView {
                                    id: listaUsuarios
                                    model: backend ? backend.usuariosComparacao : []
                                    spacing: 2
    
                                    delegate: Rectangle {
                                        width: listaUsuarios.width
                                        height: 35
                                        color: modelData.selecionado ? cores.hover : (index % 2 === 0 ? cores.linhaTabela : cores.cardBackground)
                                        border.color: modelData.selecionado ? cores.acento : "transparent"
                                        border.width: modelData.selecionado ? 2 : 0
                                        radius: 6
    
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.toggleUsuarioComparacao(modelData.nome)
                                                }
                                            }
                                        }
    
                                        RowLayout {
                                            anchors.fill: parent
                                            anchors.margins: 8
                                            spacing: 8
    
                                            Rectangle {
                                                Layout.preferredWidth: 18
                                                Layout.preferredHeight: 18
                                                radius: 3
                                                color: modelData.selecionado ? cores.acento : "transparent"
                                                border.color: cores.acento
                                                border.width: 2
    
                                                Text {
                                                    anchors.centerIn: parent
                                                    text: "✓"
                                                    color: "#ffffff"
                                                    font.bold: true
                                                    font.pixelSize: 12
                                                    visible: modelData.selecionado
                                                }
                                            }
    
                                            Text {
                                                text: modelData.nome
                                                font.pixelSize: 13
                                                color: cores.textoPrimario
                                                Layout.fillWidth: true
                                                elide: Text.ElideRight
                                            }
    
                                            Text {
                                                text: modelData.total.toFixed(1)
                                                font.pixelSize: 12
                                                color: cores.acento
                                                font.bold: true
                                            }
                                        }
                                    }
                                }
                            }
    
                            RowLayout {
                                Layout.fillWidth: true
                                spacing: 10
                            
                                Rectangle {
                                    Layout.fillWidth: true
                                    Layout.preferredHeight: 35
                                    radius: 8
                                    color: cores.acento
                            
                                    MouseArea {
                                        anchors.fill: parent
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: {
                                            if (backend) {
                                                backend.selecionarTodosUsuarios()
                                            }
                                        }
                                    }
                            
                                    Text {
                                        anchors.centerIn: parent
                                        text: "Selecionar Todos"
                                        color: "#ffffff"
                                        font.bold: true
                                        font.pixelSize: 12
                                    }
                                }
                            
                                Rectangle {
                                    Layout.fillWidth: true
                                    Layout.preferredHeight: 35
                                    radius: 8
                                    color: cores.textoSecundario
                            
                                    MouseArea {
                                        anchors.fill: parent
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: {
                                            if (backend) {
                                                backend.limparSelecaoUsuarios()
                                            }
                                        }
                                    }
                            
                                    Text {
                                        anchors.centerIn: parent
                                        text: "Limpar Seleção"
                                        color: cores.textoPrimario
                                        font.bold: true
                                        font.pixelSize: 12
                                    }
                                }
                            }
                        }
                    }
    
                    Rectangle {
                        Layout.fillWidth: true
                        Layout.fillHeight: true
                        color: cores.cardBackground
                        border.color: cores.bordaCard   
                        border.width: 2
                        radius: 16
    
                        ColumnLayout {
                            anchors.fill: parent
                            anchors.margins: 15
                            spacing: 10
    
                            Text {
                                text: "Gráfico de Comparação"
                                font.pixelSize: 18
                                font.bold: true
                                color: cores.textoTitulo
                                Layout.alignment: Qt.AlignHCenter
                            }
    
                            ChartView {
                                id: graficoComparacao
                                Layout.fillWidth: true
                                Layout.fillHeight: true
                                antialiasing: true
                                backgroundColor: cores.fundoGrafico
                                legend.visible: true
                                legend.alignment: Qt.AlignBottom
                                title: "Produtividade por Dia da Semana"
                                titleFont: Qt.font({ pixelSize: 14, bold: true })
                                titleColor: cores.textoTitulo
                                plotAreaColor: cores.fundoGrafico
                                
                                BarCategoryAxis {
                                    id: eixoX
                                    categories: ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
                                    labelsFont.pixelSize: 10
                                    labelsColor: cores.textoPrimario
                                    gridLineColor: cores.bordaInput
                                    labelsAngle: -45
                                }
                                
                                ValueAxis {
                                    id: eixoY
                                    min: 0
                                    max: backend ? backend.maxComparacaoValue : 100
                                    tickCount: 6
                                    labelFormat: "%d"
                                    labelsFont.pixelSize: 10
                                    labelsColor: cores.textoPrimario
                                    gridLineColor: cores.bordaInput
                                }
                            
                                property var activeSeries: []
                            
                                function limparGrafico() {
                                    console.log("Limpando gráfico...")
                                    for (var i = 0; i < activeSeries.length; i++) {
                                        graficoComparacao.removeSeries(activeSeries[i])
                                    }
                                    activeSeries = []
                                }
                            
                                function atualizarGrafico() {
                                    console.log("=== ATUALIZANDO GRÁFICO DE COMPARAÇÃO ===")
                                    
                                    limparGrafico()
                                    
                                    if (!backend || !backend.dadosComparacao || backend.dadosComparacao.length === 0) {
                                        console.log("Nenhum dado de comparação disponível")
                                        return
                                    }
                                
                                    var cores = ["#3cb3e6", "#4caf50", "#ff9800", "#f44336", "#9c27b0", "#607d8b", "#795548", "#e91e63"]
                                    
                                    eixoY.max = backend.maxComparacaoValue
                                    
                                    for (var i = 0; i < backend.dadosComparacao.length; i++) {
                                        var userData = backend.dadosComparacao[i]
                                        console.log("Criando série para:", userData.nome, "com valores:", userData.valores)
                                        
                                        var series = graficoComparacao.createSeries(ChartView.SeriesTypeLine, userData.nome, eixoX, eixoY)
                                        
                                        if (series) {
                                            series.color = cores[i % cores.length]
                                            series.width = 3
                                            series.pointsVisible = true
                                            series.pointLabelsVisible = false
                                            
                                            for (var j = 0; j < userData.valores.length && j < 7; j++) {
                                                var valor = userData.valores[j] || 0
                                                series.append(j, valor)
                                            }
                                            
                                            activeSeries.push(series)
                                            console.log("Série criada com sucesso:", userData.nome)
                                        } else {
                                            console.log("Erro ao criar série para:", userData.nome)
                                        }
                                    }
                                    
                                    console.log("Total de séries ativas:", activeSeries.length)
                                }
                                
                                Connections {
                                    target: backend
                                    function onComparacaoDataChanged() {
                                        console.log("Sinal comparacaoDataChanged recebido")
                                        Qt.callLater(function() {
                                            graficoComparacao.atualizarGrafico()
                                        })
                                    }
                                    function onMaxComparacaoValueChanged() {
                                        console.log("Max value changed:", backend.maxComparacaoValue)
                                        eixoY.max = backend.maxComparacaoValue
                                    }
                                    function onUsuariosComparacaoChanged() {
                                        console.log("Lista de usuários mudou")
                                        listaUsuarios.model = backend.usuariosComparacao
                                    }
                                }
                                
                                Component.onCompleted: {
                                    console.log("ChartView inicializado")
                                    if (backend) {
                                        backend.gerarDadosComparacao()
                                    }
                                }
                            
                                Text {
                                    anchors.centerIn: parent
                                    visible: !backend || !backend.dadosComparacao || backend.dadosComparacao.length === 0
                                    text: "Selecione usuários para comparar"
                                    color: cores.acento
                                    font.pixelSize: 16
                                    font.bold: true
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    Component {
        id: configPage
    
        Rectangle {
            color: "transparent"
            anchors.fill: parent
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
                spacing: 30
    
                Text {
                    text: "Configurações do Sistema"
                    font.pixelSize: 28
                    font.bold: true
                    color: cores.acento
                    horizontalAlignment: Text.AlignHCenter
                    Layout.alignment: Qt.AlignHCenter
                }
    
                ScrollView {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    clip: true
    
                    ColumnLayout {
                        width: parent.width
                        spacing: 25
    
                        Rectangle {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 200
                            color: cores.cardBackground
                            border.color: cores.bordaCard
                            border.width: 2
                            radius: 16
                        
                            ColumnLayout {
                                anchors.fill: parent
                                anchors.margins: 20
                                spacing: 15
                        
                                Text {
                                    text: "Informações do Sistema"
                                    font.pixelSize: 20
                                    font.bold: true
                                    color: cores.textoTitulo
                                    Layout.alignment: Qt.AlignHCenter
                                }
                        
                                GridLayout {
                                    Layout.fillWidth: true
                                    columns: 2
                                    rowSpacing: 12
                                    columnSpacing: 30
                        
                                    Text {
                                        text: "Versão do Aplicativo:"
                                        font.pixelSize: 14
                                        color: cores.textoPrimario
                                        font.bold: true
                                    }
                                    Text {
                                        text: backend ? backend.versaoApp : "1.0.0"
                                        font.pixelSize: 14
                                        color: cores.acento
                                        font.bold: true
                                    }
                        
                                    Text {
                                        text: "Data de Compilação:"
                                        font.pixelSize: 14
                                        color: cores.textoPrimario
                                        font.bold: true
                                    }
                                    Text {
                                        text: backend ? backend.dataCompilacao : "Junho 2025"
                                        font.pixelSize: 14
                                        color: cores.acento
                                        font.bold: true
                                    }
                        
                                    Text {
                                        text: "Sistema Operacional:"
                                        font.pixelSize: 14
                                        color: cores.textoPrimario
                                        font.bold: true
                                    }
                                    Text {
                                        text: backend ? backend.statusSistema : "Windows"
                                        font.pixelSize: 14
                                        color: cores.acento
                                        font.bold: true
                                    }
                                    
                                    Text {
                                        text: "Arquivos de Dados:"
                                        font.pixelSize: 14
                                        color: cores.textoPrimario
                                        font.bold: true
                                    }
                                    Text {
                                        text: backend ? backend.totalArquivosDados + " arquivos" : "0 arquivos"
                                        font.pixelSize: 14
                                        color: cores.acento
                                        font.bold: true
                                    }
                                }
                            }
                        }
                        
                        Rectangle {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 160
                            color: cores.cardBackground
                            border.color: cores.bordaCard
                            border.width: 2
                            radius: 16
                        
                            ColumnLayout {
                                anchors.fill: parent
                                anchors.margins: 20
                                spacing: 12
                        
                                Text {
                                    text: "Personalização"
                                    font.pixelSize: 20
                                    font.bold: true
                                    color: cores.textoTitulo
                                    Layout.alignment: Qt.AlignHCenter
                                }
                        
                                RowLayout {
                                    Layout.fillWidth: true
                                    Layout.alignment: Qt.AlignHCenter
                                    spacing: 15
                        
                                    Text {
                                        text: "Tema do Sistema:"
                                        font.pixelSize: 15
                                        color: cores.textoPrimario
                                        font.bold: true
                                    }
                        
                                    Rectangle {
                                        Layout.preferredWidth: 180
                                        Layout.preferredHeight: 45
                                        radius: 10
                                        color: temaEscuro ? cores.acento : cores.hover
                                        border.color: cores.bordaCard
                                        border.width: 2
                        
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.toggleTema()
                                                }
                                            }
                                        }
                        
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 10
                        
                                            Text {
                                                text: temaEscuro ? "🌙" : "☀️"
                                                font.pixelSize: 16
                                            }
                        
                                            Text {
                                                text: temaEscuro ? "Modo Escuro" : "Modo Claro"
                                                color: temaEscuro ? "#ffffff" : cores.textoPrimario
                                                font.bold: true
                                                font.pixelSize: 13
                                            }
                                        }
                                    }
                                }
                        
                                Text {
                                    text: "Clique para alternar entre os modos claro e escuro"
                                    font.pixelSize: 11
                                    color: cores.textoSecundario
                                    Layout.alignment: Qt.AlignHCenter
                                }
                            }
                        }
    
                        Rectangle {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 180
                            color: cores.cardBackground
                            border.color: cores.bordaCard
                            border.width: 2
                            radius: 16
    
                            ColumnLayout {
                                anchors.fill: parent
                                anchors.margins: 20
                                spacing: 15
    
                                Text {
                                    text: "Gerenciar Pastas"
                                    font.pixelSize: 20
                                    font.bold: true
                                    color: cores.textoTitulo
                                    Layout.alignment: Qt.AlignHCenter
                                }
    
                                GridLayout {
                                    Layout.fillWidth: true
                                    columns: 2
                                    rowSpacing: 12
                                    columnSpacing: 15
    
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 40
                                        radius: 8
                                        color: cores.acento
                                        border.color: cores.bordaCard
                                        border.width: 1
    
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.abrirPastaApp()
                                                }
                                            }
                                        }
    
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 8
    
                                            Image {
                                                source: "assets/icons/dashboard.png"
                                                Layout.preferredWidth: 16
                                                Layout.preferredHeight: 16
                                                fillMode: Image.PreserveAspectFit
                                            }
    
                                            Text {
                                                text: "Pasta do Aplicativo"
                                                color: "#ffffff"
                                                font.bold: true
                                                font.pixelSize: 13
                                            }
                                        }
                                    }
    
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 40
                                        radius: 8
                                        color: cores.acento
                                        border.color: cores.bordaCard
                                        border.width: 1
    
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.abrirPastaDados()
                                                }
                                            }
                                        }
    
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 8
    
                                            Image {
                                                source: "assets/icons/excel.png"
                                                Layout.preferredWidth: 16
                                                Layout.preferredHeight: 16
                                                fillMode: Image.PreserveAspectFit
                                            }
    
                                            Text {
                                                text: "Dados Mensais"
                                                color: "#ffffff"
                                                font.bold: true
                                                font.pixelSize: 13
                                            }
                                        }
                                    }
    
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 40
                                        radius: 8
                                        color: cores.acento
                                        border.color: cores.bordaCard
                                        border.width: 1
    
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.abrirPastaSalvos()
                                                }
                                            }
                                        }
    
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 8
    
                                            Image {
                                                source: "assets/icons/peso.png"
                                                Layout.preferredWidth: 16
                                                Layout.preferredHeight: 16
                                                fillMode: Image.PreserveAspectFit
                                            }
    
                                            Text {
                                                text: "Configurações"
                                                color: "#ffffff"
                                                font.bold: true
                                                font.pixelSize: 13
                                            }
                                        }
                                    }
    
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 40
                                        radius: 8
                                        color: cores.acento
                                        border.color: cores.bordaCard
                                        border.width: 1
    
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                if (backend) {
                                                    backend.abrirPastaExports()
                                                }
                                            }
                                        }
    
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 8
    
                                            Image {
                                                source: "assets/icons/settings.png"
                                                Layout.preferredWidth: 16
                                                Layout.preferredHeight: 16
                                                fillMode: Image.PreserveAspectFit
                                            }
    
                                            Text {
                                                text: "Pasta de Exports"
                                                color: "#ffffff"
                                                font.bold: true
                                                font.pixelSize: 13
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        Rectangle {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 160
                            color: cores.cardBackground
                            border.color: cores.bordaCard
                            border.width: 2
                            radius: 16
    
                            ColumnLayout {
                                anchors.fill: parent
                                anchors.margins: 20
                                spacing: 15
    
                                Text {
                                    text: "Manutenção do Sistema"
                                    font.pixelSize: 20
                                    font.bold: true
                                    color: cores.textoTitulo
                                    Layout.alignment: Qt.AlignHCenter
                                }
    
                                Text {
                                    text: "Cache atual: " + (backend ? backend.tamanhoCache : "0 itens")
                                    font.pixelSize: 13
                                    color: cores.textoSecundario
                                    Layout.alignment: Qt.AlignHCenter
                                }
    
                                RowLayout {
                                    Layout.fillWidth: true
                                    spacing: 15
    
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 45
                                        radius: 8
                                        color: cores.textoSecundario
                                        border.color: cores.bordaCard
                                        border.width: 1
    
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                limparCachePopup.open()
                                            }
                                        }
    
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 8
    
                                            Image {
                                                source: "assets/icons/peso4.png"
                                                Layout.preferredWidth: 18
                                                Layout.preferredHeight: 18
                                                fillMode: Image.PreserveAspectFit
                                            }
    
                                            Text {
                                                text: "Limpar Cache"
                                                color: cores.textoPrimario
                                                font.bold: true
                                                font.pixelSize: 14
                                            }
                                        }
                                    }
    
                                    Rectangle {
                                        Layout.fillWidth: true
                                        Layout.preferredHeight: 45
                                        radius: 8
                                        color: cores.textoSecundario
                                        border.color: cores.bordaCard
                                        border.width: 1
    
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: {
                                                resetConfigPopup.open()
                                            }
                                        }
    
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 8
    
                                            Image {
                                                source: "assets/icons/settings.png"
                                                Layout.preferredWidth: 18
                                                Layout.preferredHeight: 18
                                                fillMode: Image.PreserveAspectFit
                                            }
    
                                            Text {
                                                text: "Resetar Configurações"
                                                color: cores.textoPrimario
                                                font.bold: true
                                                font.pixelSize: 14
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        Rectangle {
                            Layout.fillWidth: true
                            Layout.preferredHeight: 120
                            color: cores.cardBackground
                            border.color: cores.bordaCard
                            border.width: 2
                            radius: 16
    
                            ColumnLayout {
                                anchors.fill: parent
                                anchors.margins: 20
                                spacing: 10
    
                                Text {
                                    text: "Sobre o Projeto"
                                    font.pixelSize: 20
                                    font.bold: true
                                    color: cores.textoTitulo
                                    Layout.alignment: Qt.AlignHCenter
                                }
    
                                Text {
                                    text: "Sistema de Análise de Produtividade APROD\nDesenvolvido para gerenciar e analisar dados de produtividade com ferramentas avançadas de visualização e relatórios."
                                    font.pixelSize: 14
                                    color: cores.textoPrimario
                                    horizontalAlignment: Text.AlignHCenter
                                    Layout.alignment: Qt.AlignHCenter
                                    wrapMode: Text.WordWrap
                                    Layout.fillWidth: true
                                }
                            }
                        }
                    }
                }
            }
    
            Popup {
                id: limparCachePopup
                width: 400
                height: 160
                modal: true
                focus: true
                closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
                anchors.centerIn: parent
    
                Rectangle {
                    anchors.fill: parent
                    color: cores.fundoPopup
                    radius: 12
                    border.color: cores.bordaCard
                    border.width: 2
    
                    ColumnLayout {
                        anchors.fill: parent
                        anchors.margins: 20
                        spacing: 20
    
                        Text {
                            text: "Deseja limpar o cache do sistema?"
                            font.pixelSize: 16
                            color: cores.textoPrimario
                            wrapMode: Text.WordWrap
                            Layout.fillWidth: true
                            horizontalAlignment: Text.AlignHCenter
                            font.bold: true
                        }
    
                        Text {
                            text: "Isso irá remover todos os dados temporários carregados."
                            font.pixelSize: 12
                            color: cores.textoSecundario
                            wrapMode: Text.WordWrap
                            Layout.fillWidth: true
                            horizontalAlignment: Text.AlignHCenter
                        }
    
                        RowLayout {
                            Layout.alignment: Qt.AlignHCenter
                            spacing: 20
    
                            Rectangle {
                                width: 100
                                height: 40
                                radius: 8
                                color: cores.acento
    
                                MouseArea {
                                    anchors.fill: parent
                                    cursorShape: Qt.PointingHandCursor
                                    onClicked: {
                                        limparCachePopup.close()
                                        if (backend) {
                                            backend.limparDadosCache()
                                        }
                                    }
                                }
    
                                Text {
                                    anchors.centerIn: parent
                                    text: "Sim"
                                    color: "#fff"
                                    font.bold: true
                                    font.pixelSize: 14
                                }
                            }
    
                            Rectangle {
                                width: 100
                                height: 40
                                radius: 8
                                color: cores.textoSecundario
    
                                MouseArea {
                                    anchors.fill: parent
                                    cursorShape: Qt.PointingHandCursor
                                    onClicked: limparCachePopup.close()
                                }
    
                                Text {
                                    anchors.centerIn: parent
                                    text: "Cancelar"
                                    color: cores.textoPrimario
                                    font.bold: true
                                    font.pixelSize: 14
                                }
                            }
                        }
                    }
                }
            }
            
            Popup {
                id: resetConfigPopup
                width: 450
                height: 180
                modal: true
                focus: true
                closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside
                anchors.centerIn: parent
    
                Rectangle {
                    anchors.fill: parent
                    color: cores.fundoPopup
                    radius: 12
                    border.color: cores.bordaCard
                    border.width: 2
    
                    ColumnLayout {
                        anchors.fill: parent
                        anchors.margins: 20
                        spacing: 20
    
                        Text {
                            text: "⚠️ Resetar todas as configurações?"
                            font.pixelSize: 16
                            color: cores.textoPrimario
                            wrapMode: Text.WordWrap
                            Layout.fillWidth: true
                            horizontalAlignment: Text.AlignHCenter
                            font.bold: true
                        }
    
                        Text {
                            text: "Isso irá:\n• Limpar todos os meses ativos\n• Restaurar pesos padrão\n• Remover filtros salvos\n• Resetar ordenações\n• Voltar ao tema padrão"
                            font.pixelSize: 12
                            color: cores.textoSecundario
                            wrapMode: Text.WordWrap
                            Layout.fillWidth: true
                            horizontalAlignment: Text.AlignLeft
                        }
    
                        RowLayout {
                            Layout.alignment: Qt.AlignHCenter
                            spacing: 20
    
                            Rectangle {
                                width: 120
                                height: 40
                                radius: 8
                                color: cores.acento
    
                                MouseArea {
                                    anchors.fill: parent
                                    cursorShape: Qt.PointingHandCursor
                                    onClicked: {
                                        resetConfigPopup.close()
                                        if (backend) {
                                            backend.resetarConfiguracoes()
                                        }
                                    }
                                }
    
                                Text {
                                    anchors.centerIn: parent
                                    text: "Resetar"
                                    color: "#fff"
                                    font.bold: true
                                    font.pixelSize: 14
                                }
                            }
    
                            Rectangle {
                                width: 100
                                height: 40
                                radius: 8
                                color: cores.textoSecundario
    
                                MouseArea {
                                    anchors.fill: parent
                                    cursorShape: Qt.PointingHandCursor
                                    onClicked: resetConfigPopup.close()
                                }
    
                                Text {
                                    anchors.centerIn: parent
                                    text: "Cancelar"
                                    color: cores.textoPrimario
                                    font.bold: true
                                    font.pixelSize: 14
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

