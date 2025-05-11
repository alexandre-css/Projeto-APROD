import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.Material
import QtCharts
import QtQuick.Effects

ApplicationWindow {
    width: 1440
    height: 900
    visible: true
    visibility: ApplicationWindow.Maximized
    Material.theme: Material.Light
    Material.primary: "#3cb3e6"
    Material.accent: "#1976d2"

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
            // Dentro do seu Rectangle do menu lateral:
            Rectangle {
                id: sidebar
                width: hovered ? 220 : 110
                color: "#fff"
                border.color: "#e0e0e0"
                Layout.fillHeight: true
                z: 3
                property bool hovered: false

                Behavior on width { NumberAnimation { duration: 180; easing.type: Easing.OutQuad } }

                MouseArea {
                    anchors.fill: parent
                    hoverEnabled: true
                    onEntered: sidebar.hovered = true
                    onExited: sidebar.hovered = false
                    cursorShape: Qt.PointingHandCursor
                }

                ColumnLayout {
                    anchors.fill: parent
                    spacing: 40

                    Item { Layout.fillHeight: true }

                    Repeater {
                        model: [
                            { icon: "dashboard.png", label: "Dashboard" },
                            { icon: "compara.png", label: "Comparar" },
                            { icon: "bars.png", label: "Gráficos" },
                            { icon: "peso.png", label: "Peso" },
                            { icon: "settings.png", label: "Configurações" }
                        ]
                        RowLayout {
                            Layout.alignment: Qt.AlignLeft
                            spacing: 14
                            width: parent.width

                            // Espaçador à esquerda para afastar do canto do menu
                            Item {
                                width: 15 // ajuste conforme necessário para o visual institucional
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
                            // Espaçador invisível para ocupar o resto do espaço
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

                // CONTAINER INSTITUCIONAL ANCORADO NO TOPO (sem height fixo)
                ColumnLayout {
                    anchors.top: parent.top
                    anchors.topMargin: 48
                    anchors.horizontalCenter: parent.horizontalCenter
                    width: Math.min(parent.width - 160, 1200)
                    spacing: 24

                    // KPIs centralizados, espaçamento menor, textos protegidos
                    RowLayout {
                        Layout.fillWidth: true
                        Layout.alignment: Qt.AlignHCenter
                        spacing: 18

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
                    Item { height: 6 }

                    // Botões e campo de meses centralizados
                    RowLayout {
                        Layout.fillWidth: true
                        Layout.alignment: Qt.AlignHCenter
                        spacing: 18

                        Rectangle {
                            id: importButton
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

                        Rectangle {
                            id: arquivosCarregadosBox
                            width: 440
                            height: 48
                            radius: 14
                            border.color: "#3cb3e6"
                            border.width: 2
                            color: hovered ? "#e1f5fe" : "#f5faff"
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
                                onEntered: arquivosCarregadosBox.hovered = true
                                onExited: arquivosCarregadosBox.hovered = false
                                cursorShape: Qt.PointingHandCursor
                            }
                            RowLayout {
                                anchors.fill: parent
                                anchors.margins: 10
                                spacing: 14

                                Image {
                                    source: "assets/icons/xls.png"
                                    Layout.preferredWidth: 30
                                    Layout.preferredHeight: 30
                                    Layout.minimumWidth: 30
                                    Layout.minimumHeight: 30
                                    Layout.maximumWidth: 30
                                    Layout.maximumHeight: 30
                                    fillMode: Image.PreserveAspectFit
                                }

                                Text {
                                    text: backend && backend.arquivosCarregados ? backend.arquivosCarregados : ""
                                    font.pixelSize: 17
                                    font.bold: true
                                    color: "#232946"
                                    verticalAlignment: Text.AlignVCenter
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                    Layout.fillWidth: true
                                }
                            }
                        }
                    }

                    // Espaço entre botões/campo de meses e gráfico
                    Item { height: 6 }

                    // Gráfico centralizado (altura garantida)
                    Rectangle {
                        width: Math.min(parent.width - 160, 1000)
                        height: 650 // ou mais, conforme sua preferência
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
                                categories: backend.nomes
                                labelsFont.pixelSize: 12
                            }
                            ValueAxis {
                                id: eixoValores
                                min: 0
                                max: 280
                                labelFormat: "%.0f"
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
}
