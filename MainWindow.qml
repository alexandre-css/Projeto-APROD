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
                            color: currentPage === modelData.page ? "#e3f2fd" : "transparent"
                            radius: 12
                            border.width: currentPage === modelData.page ? 2 : 0
                            border.color: "#3cb3e6"
                            
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
                                    color: currentPage === modelData.page ? "#1976d2" : "#424242"
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
                        if (backend) {
                            backend.handle_page_loaded(currentPage)
                        }
                    }
                    
                    Connections {
                        target: backend
                        function onChartUpdateRequired() {
                            if (currentPage === "dashboard") {
                                pageLoader.item.forceLayout()
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
            color: "#ffffff"
            radius: 8
            border.color: "#3cb3e6"
            border.width: 2

            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 20
                spacing: 20

                Text {
                    text: "Você tem certeza que deseja restaurar todos os pesos para 1?"
                    font.pixelSize: 16
                    color: "#232946"
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
                        color: "#3cb3e6"

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
                        color: "#e0e0e0"

                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: restoreConfirmation.close()
                        }

                        Text {
                            anchors.centerIn: parent
                            text: "Cancelar"
                            color: "#232946"
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
            color: "#ffffff"
            radius: 12
            border.color: "#3cb3e6"
            border.width: 2
    
            ColumnLayout {
                anchors.fill: parent
                anchors.margins: 25
                spacing: 25
    
                Text {
                    text: "Selecione o formato para exportação:"
                    font.pixelSize: 18
                    color: "#232946"
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
                        color: "#3cb3e6"
    
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
            color: "#ffffff"
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
                    color: "#232946"
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
                        color: "#e0e0e0"
    
                        MouseArea {
                            anchors.fill: parent
                            cursorShape: Qt.PointingHandCursor
                            onClicked: exportSuccessPopup.close()
                        }
    
                        Text {
                            anchors.centerIn: parent
                            text: "Fechar"
                            color: "#232946"
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
                        color: "#f5faff"
                        border.color: "#3cb3e6"
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
                                    color: "#3cb3e6"
                                    font.bold: true
                                }
                                Text {
                                    text: kpis ? kpis["minutas"] : "0"
                                    font.pixelSize: 28
                                    color: "#232946"
                                    font.bold: true
                                }
                            }
                        }
                    }
                
                    Rectangle {
                        Layout.preferredWidth: 280
                        Layout.preferredHeight: 120
                        radius: 16
                        color: "#f5faff"
                        border.color: "#3cb3e6"
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
                                    color: "#3cb3e6"
                                    font.bold: true
                                }
                                Text {
                                    text: backend && backend.kpis ? backend.kpis["dia_produtivo"] : "--"
                                    font.pixelSize: backend && backend.kpis && backend.kpis["dia_produtivo"] && backend.kpis["dia_produtivo"].includes("\n") ? 20 : 28
                                    color: "#232946"
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
                        color: "#f5faff"
                        border.color: "#3cb3e6"
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
                                    color: "#3cb3e6"
                                    font.bold: true
                                }
                
                                Column {
                                    spacing: 2
                                    
                                    Repeater {
                                        model: {
                                            if (!kpis || !kpis["top3"]) return []
                                            return kpis["top3"].split(" | ").slice(0, 3)
                                        }
                                        
                                        Text {
                                            text: modelData
                                            font.pixelSize: 12
                                            color: "#232946"
                                            font.bold: true
                                            elide: Text.ElideRight
                                            width: 260
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
                            GradientStop { position: 0.0; color: "#3cb3e6" }
                            GradientStop { position: 1.0; color: "#1976d2" }
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
                        color: "#f5faff"
                        border.color: "#3cb3e6"
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
                                color: "#3cb3e6"
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
    
                                    color: isActive ? "#3cb3e6" : "#f5f5f5"
                                    border.color: isActive ? "#1976d2" : "#d0d0d0"
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
                                        color: parent.isActive ? "#ffffff" : "#999999"
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
                    color: "#fff"
                    border.color: "#3cb3e6"
                    border.width: 2
    
                    ChartView {
                        anchors.fill: parent
                        anchors.margins: 10
                        antialiasing: true
                        backgroundColor: "#fff"
                        legend.visible: false
                        title: "Minutas por Usuário"
                        titleFont: Qt.font({ pixelSize: 16, bold: true })
                        titleColor: "#1976d2"
                        plotAreaColor: "#ffffff"
                        margins {
                            top: 20
                            bottom: 20
                            left: 20
                            right: 20
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
                                color: "#3cb3e6"
                                borderColor: "#3cb3e6"
                                borderWidth: 0
                                labelColor: "#232946"
                                labelFont.pixelSize: 12
                            }
                        }
    
                        BarCategoryAxis {
                            id: eixoUsuarios
                            categories: backend && backend.nomes ? backend.nomes : []
                            labelsFont.pixelSize: 12
                            labelsColor: "#232946"
                            gridVisible: false
                        }
    
                        ValueAxis {
                            id: eixoValores
                            min: 0
                            max: backend ? backend.maxAxisValue : 100
                            tickCount: 5
                            labelFormat: "%d"
                            labelsFont.pixelSize: 12
                            labelsColor: "#232946"
                            gridLineColor: "#e0e0e0"
                            minorGridVisible: false
                        }
    
                        Text {
                            anchors.centerIn: parent
                            visible: !backend || !backend.valores || backend.valores.length === 0
                            text: "Nenhum dado disponível"
                            color: "#3cb3e6"
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
                    color: "#3cb3e6"
                    horizontalAlignment: Text.AlignHCenter
                    Layout.alignment: Qt.AlignHCenter
                }
    
                Rectangle {
                    width: 1100
                    height: 600
                    color: "#fff"
                    border.color: "#3cb3e6"
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
                            color: "#e3f2fd"
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
                                    color: "#1976d2"
                                    horizontalAlignment: Text.AlignHCenter
                                }
    
                                Text {
                                    Layout.preferredWidth: 260
                                    text: "Peso"
                                    font.pixelSize: 18
                                    font.bold: true
                                    color: "#1976d2"
                                    horizontalAlignment: Text.AlignHCenter
                                }
                            }
    
                            Rectangle {
                                anchors.bottom: parent.bottom
                                width: parent.width
                                height: 2
                                color: "#3cb3e6"
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
                                        color: index % 2 === 0 ? "#f5faff" : "#ffffff"
                                        
                                        RowLayout {
                                            anchors.fill: parent
                                            anchors.margins: 10
                                            spacing: 0
                                            
                                            Text {
                                                Layout.preferredWidth: 800
                                                text: modelData ? modelData.tipo : ""
                                                font.pixelSize: 16
                                                color: "#232946"
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
                                                    Layout.preferredWidth: 40
                                                    Layout.preferredHeight: 35
                                                    radius: 8
                                                    color: "#e3f2fd"
                                                    border.color: "#3cb3e6"
                                                    border.width: 1
                                                    
                                                    MouseArea {
                                                        anchors.fill: parent
                                                        cursorShape: Qt.PointingHandCursor
                                                        onClicked: {
                                                            if (backend) {
                                                                let currentValue = modelData ? modelData.peso : 1.0
                                                                let newValue = Math.max(1.0, currentValue - 1.0)
                                                                backend.atualizarPeso(index, newValue)
                                                            }
                                                        }
                                                    }
                                                    
                                                    Text {
                                                        text: "-"
                                                        anchors.centerIn: parent
                                                        font.pixelSize: 20
                                                        font.bold: true
                                                        color: "#3cb3e6"
                                                    }
                                                }
                                                
                                                Rectangle {
                                                    Layout.preferredWidth: 80
                                                    Layout.preferredHeight: 35
                                                    color: "#ffffff"
                                                    border.color: "#3cb3e6"
                                                    border.width: 1
                                                    radius: 6
                                                    
                                                    TextInput {
                                                        anchors.fill: parent
                                                        anchors.margins: 8
                                                        text: modelData ? Number(modelData.peso).toFixed(0) : "1"
                                                        font.pixelSize: 16
                                                        color: "#232946"
                                                        horizontalAlignment: TextInput.AlignHCenter
                                                        verticalAlignment: TextInput.AlignVCenter
                                                        selectByMouse: true
                                                        
                                                        onEditingFinished: {
                                                            if (backend) {
                                                                let value = Math.max(1, Math.min(100, parseInt(text) || 1))
                                                                backend.atualizarPeso(index, value)
                                                            }
                                                        }
                                                    }
                                                }
                                                
                                                Rectangle {
                                                    Layout.preferredWidth: 40
                                                    Layout.preferredHeight: 35
                                                    radius: 8
                                                    color: "#e3f2fd"
                                                    border.color: "#3cb3e6"
                                                    border.width: 1
                                                    
                                                    MouseArea {
                                                        anchors.fill: parent
                                                        cursorShape: Qt.PointingHandCursor
                                                        onClicked: {
                                                            if (backend) {
                                                                let currentValue = modelData ? modelData.peso : 1.0
                                                                let newValue = Math.min(10.0, currentValue + 1.0)
                                                                backend.atualizarPeso(index, newValue)
                                                            }
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
                                        
                                        Rectangle {
                                            anchors.bottom: parent.bottom
                                            width: parent.width
                                            height: 1
                                            color: "#e0e0e0"
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
                        width: 180
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
                            GradientStop { position: 0.0; color: "#3cb3e6" }
                            GradientStop { position: 1.0; color: "#1976d2" }
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
                            GradientStop { position: 0.0; color: "#3cb3e6" }
                            GradientStop { position: 1.0; color: "#1976d2" }
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
    
            property bool sortAscending: backend.sortAscending
            property string sortColumn: backend.sortColumn
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
                    color: "#3cb3e6"
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
                    color: "#f5faff"
                    radius: 12
                    border.width: 1
                    border.color: "#3cb3e6"
                    
                    RowLayout {
                        anchors.fill: parent
                        anchors.margins: 10
                        spacing: 20
                        
                        Text {
                            text: "Filtrar por usuário:"
                            font.pixelSize: 15
                            color: "#1976d2"
                            font.bold: true
                        }
                        
                        Rectangle {
                            Layout.preferredWidth: 250
                            Layout.preferredHeight: 36
                            color: "#ffffff"
                            radius: 6
                            border.width: 1
                            border.color: "#3cb3e6"
                            
                            TextInput {
                                id: filtroInput
                                anchors.fill: parent
                                anchors.margins: 8
                                verticalAlignment: TextInput.AlignVCenter
                                font.pixelSize: 14
                                color: "#232946"
                                selectByMouse: true
                                
                                onTextChanged: {
                                    backend.setFiltroUsuario(text)
                                }
                                
                                Text {
                                    anchors.fill: parent
                                    text: "Digite para filtrar..."
                                    font.pixelSize: 14
                                    color: "#aaaaaa"
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
                                GradientStop { position: 0.0; color: "#3cb3e6" }
                                GradientStop { position: 1.0; color: "#1976d2" }
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
                    Layout.maximumHeight: parent.height - 250
                    Layout.topMargin: -50
                    color: "#fff"
                    border.color: "#3cb3e6"
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
                        color: "#3cb3e6"
                        font.bold: true
                    }
                
                    Rectangle {
                        anchors.fill: parent
                        anchors.margins: 2
                        color: "transparent"
                        border.color: "#3cb3e6"
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
                                color: "#e3f2fd"
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
                                            color: "#1976d2"
                                        }
                                        
                                        Text {
                                            text: backend.sortColumn === "usuario" ? (backend.sortAscending ? "▲" : "▼") : ""
                                            font.pixelSize: 14
                                            color: "#3cb3e6"
                                            font.bold: true
                                        }
                                    }
                                    
                                    MouseArea {
                                        anchors.fill: parent
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: backend.atualizar_sort_table("usuario")
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
                                        border.color: "#3cb3e6"
                                        
                                        RowLayout {
                                            anchors.centerIn: parent
                                            spacing: 5
                                    
                                            Text {
                                                text: modelData.name
                                                font.pixelSize: 16
                                                font.bold: true
                                                color: "#1976d2"
                                            }
                                            
                                            Text {
                                                text: backend.sortColumn === modelData.key ? (backend.sortAscending ? "▲" : "▼") : ""
                                                font.pixelSize: 14
                                                color: "#3cb3e6"
                                                font.bold: true
                                            }
                                        }
                                        
                                        MouseArea {
                                            anchors.fill: parent
                                            cursorShape: Qt.PointingHandCursor
                                            onClicked: backend.atualizar_sort_table(modelData.key)
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
                                            color: "#1976d2"
                                        }
                                        
                                        Text {
                                            text: backend.sortColumn === "Total" ? (backend.sortAscending ? "▲" : "▼") : ""
                                            font.pixelSize: 14
                                            color: "#3cb3e6"
                                            font.bold: true
                                        }
                                    }
                                    
                                    MouseArea {
                                        anchors.fill: parent
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: backend.atualizar_sort_table("Total")
                                    }
                                }
                            }
            
                            Rectangle {
                                anchors.bottom: parent.bottom
                                width: parent.width
                                height: 2
                                color: "#3cb3e6"
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
                                        color: index % 2 === 0 ? "#f5faff" : "#ffffff"
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
                                                    color: "#232946"
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
                                                    border.color: "#e3f2fd"
                                                    
                                                    Text {
                                                        anchors.centerIn: parent
                                                        text: backend.get_day_value_formatted(userRowItem.itemData, modelData)
                                                        font.pixelSize: 15
                                                        color: "#232946"
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
                                                    color: "#232946"
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
                                    color: "#3cb3e6"
                                }
                            }
                        }
                    }
                }
            }
        }
            
            Rectangle {
                id: bottomPanel
                anchors.bottom: parent.bottom
                anchors.left: parent.left
                anchors.right: parent.right
                anchors.margins: 5
                height: 100
                color: "transparent"
            
                RowLayout {
                    anchors.fill: parent
                    spacing: 20
            
                    Rectangle {
                        Layout.fillWidth: true
                        Layout.preferredHeight: 80
                        radius: 12
                        color: "#f5faff"
                        border.color: "#3cb3e6"
                        border.width: 1
            
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
                                        color: index < 3 ? "#e3f2fd" : "transparent"
                                        radius: 8
                                        border.width: 1
                                        border.color: index < 3 ? "#3cb3e6" : "#e0e0e0"
                                        
                                        Text {
                                            id: rankItemText
                                            anchors.centerIn: parent
                                            text: modelData.rankText
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
            
                    Rectangle {
                        Layout.preferredWidth: 180
                        Layout.preferredHeight: 50
                        radius: 12
                        gradient: Gradient {
                            GradientStop { position: 0.0; color: "#3cb3e6" }
                            GradientStop { position: 1.0; color: "#1976d2" }
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



    Component {
        id: compararPage
        Rectangle {
            color: "white"
            Text {
                text: "Página Comparar"
                anchors.centerIn: parent
                font.pixelSize: 24
                color: "black"
            }
        }
    }


    Component {
        id: configPage
        Rectangle {
            color: "white"
            Text {
                text: "Página de Configurações"
                anchors.centerIn: parent
                font.pixelSize: 24
                color: "black"
            }
        }
    }
}

