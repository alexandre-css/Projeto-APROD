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
                            Layout.alignment: Qt.AlignHCenter
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
                                    width: 38
                                    height: 38
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
                                        text: kpis.minutas
                                        font.pixelSize: 34
                                        color: "#232946"
                                        font.bold: true
                                        elide: Text.ElideRight
                                        wrapMode: Text.NoWrap
                                    }
                                }
                            }
                        }
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
                                    width: 38
                                    height: 38
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
                                        text: kpis.dia
                                        font.pixelSize: 20
                                        color: "#232946"
                                        font.bold: true
                                        elide: Text.ElideRight
                                        wrapMode: Text.NoWrap
                                    }
                                }
                            }
                        }
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
                                    width: 38
                                    height: 38
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
                                        text: kpis.top3
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
                                    text: backend.arquivosCarregados
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
                                    values: backend.valores
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
