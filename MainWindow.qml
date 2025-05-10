import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.Material
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
                width: 110
                color: "#fff"
                border.color: "#e0e0e0"
                Layout.fillHeight: true
                z: 3

                ColumnLayout {
                    anchors.fill: parent
                    spacing: 50

                    Item { Layout.fillHeight: true }

                    Repeater {
                        model: [
                            "dashboard.png",
                            "compara.png",
                            "bars.png",
                            "peso.png",
                            "settings.png"
                        ]
                        Image {
                            source: "assets/icons/" + modelData
                            width: 56
                            height: 56
                            fillMode: Image.PreserveAspectFit
                            anchors.horizontalCenter: parent.horizontalCenter
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

                ColumnLayout {
                    anchors.fill: parent
                    anchors.margins: 0
                    spacing: 7000

                    // BLOCO DE CONTEÚDO PRINCIPAL COM MARGEM À ESQUERDA
                    ColumnLayout {
                        width: 1220
                        anchors.left: parent.left
                        anchors.leftMargin: 32
                        spacing: 5

                        Item { height: 5 }

                        // === KPIs DINÂMICOS NO TOPO ===
                        RowLayout {
                            Layout.fillWidth: true
                            spacing: 16

                            Rectangle {
                                width: 340; height: 100
                                color: "#fff"
                                radius: 14
                                border.color: "#3cb3e6"
                                border.width: 2
                                Column {
                                    anchors.centerIn: parent
                                    spacing: 4
                                    Text {
                                        text: "Total de Minutas"
                                        font.pixelSize: 18
                                        color: "#6ec1d6"
                                        font.bold: true
                                        horizontalAlignment: Text.AlignHCenter
                                    }
                                    Text {
                                        text: kpis.minutas
                                        font.pixelSize: 34
                                        color: "#232946"
                                        font.bold: true
                                        horizontalAlignment: Text.AlignHCenter
                                    }
                                }
                            }
                            Rectangle {
                                width: 340; height: 100
                                color: "#fff"
                                radius: 14
                                border.color: "#3cb3e6"
                                border.width: 2
                                Column {
                                    anchors.centerIn: parent
                                    spacing: 4
                                    Text {
                                        text: "Dia Mais Produtivo"
                                        font.pixelSize: 18
                                        color: "#6ec1d6"
                                        font.bold: true
                                        horizontalAlignment: Text.AlignHCenter
                                    }
                                    Text {
                                        text: kpis.dia
                                        font.pixelSize: 24
                                        color: "#232946"
                                        font.bold: true
                                        horizontalAlignment: Text.AlignHCenter
                                    }
                                }
                            }
                            Rectangle {
                                width: 340; height: 100
                                color: "#fff"
                                radius: 14
                                border.color: "#3cb3e6"
                                border.width: 2
                                Column {
                                    anchors.centerIn: parent
                                    spacing: 4
                                    Text {
                                        text: "Top 3 Usuários"
                                        font.pixelSize: 18
                                        color: "#6ec1d6"
                                        font.bold: true
                                        horizontalAlignment: Text.AlignHCenter
                                    }
                                    Text {
                                        text: kpis.top3
                                        font.pixelSize: 10
                                        color: "#232946"
                                        font.bold: true
                                        horizontalAlignment: Text.AlignHCenter
                                        wrapMode: Text.WordWrap
                                        width: parent.width - 24
                                    }
                                }
                            }
                        }
                        
                        Item { height: 40 }

                        // BOTÃO IMPORTAR EXCEL
                        Rectangle {
                            id: importButton
                            width: 220
                            height: 48
                            radius: 8
                            anchors.left: parent.left
                            gradient: Gradient {
                                GradientStop { position: 0.0; color: "#3cb3e6" }
                                GradientStop { position: 1.0; color: "#1976d2" }
                            }
                            property bool hovered: false
                            property bool pressed: false
                            scale: pressed ? 0.95 : (hovered ? 1.05 : 1.0)
                            Behavior on scale { NumberAnimation { duration: 120; easing.type: Easing.OutQuad } }
                            MouseArea {
                                anchors.fill: parent
                                hoverEnabled: true
                                onEntered: importButton.hovered = true
                                onExited: importButton.hovered = false
                                onPressed: importButton.pressed = true
                                onReleased: importButton.pressed = false
                                cursorShape: Qt.PointingHandCursor
                                onClicked: backend.importar_arquivo_excel()
                            }
                            RowLayout {
                                anchors.centerIn: parent
                                spacing: 8
                                Image {
                                    source: "assets/icons/excel.png"
                                    width: 28
                                    height: 28
                                    fillMode: Image.PreserveAspectFit
                                }
                                Text {
                                    text: "Importar Excel"
                                    color: "#fff"
                                    font.bold: true
                                    font.pixelSize: 17
                                }
                            }
                        }

                        // Botão Arquivos Carregados
                        Rectangle {
                            id: arquivosCarregadosBox
                            width: 480
                            height: 48
                            radius: 10
                            border.color: "#3cb3e6"
                            border.width: 2
                            color: "#fff"
                            anchors.left: importButton.right
                            anchors.leftMargin: 16
                            anchors.verticalCenter: importButton.verticalCenter

                            RowLayout {
                                anchors.fill: parent
                                spacing: 0
                                Text {
                                    id: arquivosCarregadosText
                                    text: backend.arquivosCarregados
                                    color: "#3cb3e6"
                                    font.pixelSize: 16
                                    font.bold: false
                                    font.family: "Arial"
                                    horizontalAlignment: Text.AlignLeft
                                    verticalAlignment: Text.AlignVCenter
                                    anchors.left: parent.left
                                    anchors.leftMargin: 18
                                    elide: Text.ElideRight
                                    wrapMode: Text.NoWrap
                                    clip: true
                                }
                            }
                        }

                        // GRÁFICO DE BARRAS HORIZONTAIS
                        Rectangle {
                            width: 1150
                            height: 700
                            radius: 22
                            border.color: "#e0e0e0"
                            anchors.left: parent.left
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
                                        values: backend.valores
                                        brushFilename: "assets/gradients/bargradient.png"
                                        labelFont: Qt.font({ pixelSize: 16, bold: true, family: "Arial" })
                                        labelColor: "#1976d2"
                                    }
                                }
                                BarCategoryAxis {
                                    id: eixoUsuarios
                                    categories: backend.nomes
                                    labelsFont.pixelSize: 10
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
}
