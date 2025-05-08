import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.15
import QtCharts 2.15

ApplicationWindow {
    id: mainWindow
    width: 1280
    height: 800
    visible: true
    title: "Dashboard de Produtividade"

    Rectangle {
        Layout.fillWidth: true
        Layout.fillHeight: true
        color: "#f6f8fa"

        // AppBar superior
        Rectangle {
            height: 80
            color: "#232946"
            z: 2
            Text {
                text: "Dashboard de Produtividade"
                color: "#fff"
                font {
                    pixelSize: 28
                    bold: true
                }
            
                anchors.centerIn: parent
            }

            function calculateMax(values, defaultValue) {
                return values.length > 0 ? Math.max.apply(null, values.concat(defaultValue)) : defaultValue;
            }
            }
        }

        Row {
            Layout.fillWidth: true
            Layout.margins: 80
            // Sidebar vertical
            Rectangle {
                width: 60
                color: "#232946"
                Layout.fillHeight: true
                Column {
                    spacing: 12

                    // Botão Upload (⬆)
                    Button {
                        width: 44
                        height: 44
                        text: "⬆"
                        contentItem: Text {
                            text: "⬆"
                            horizontalAlignment: Text.AlignHCenter
                            verticalAlignment: Text.AlignVCenter
                            font {
                                pixelSize: 24
                            }
                        }
                        background: Rectangle { color: "#eebbc3"; radius: 12 }
                        Layout.alignment: Qt.AlignHCenter | Qt.AlignTop
                        Layout.topMargin: 16
                        ToolTip { text: "Importar Excel" }
                    }
                }
            }
        }

            // Conteúdo principal
            ColumnLayout {
                Layout.fillWidth: true
                Layout.fillHeight: true
                spacing: 0

                // KPIs em linha, centralizados
                RowLayout {
                    Layout.fillWidth: true
                    spacing: 24
                    Layout.margins: 40

                    // Card KPI 1
                    Rectangle {
                        width: 240; height: 100; radius: 22
                        color: "#fff"
                        border.color: "#e0e0e0"
                        Column {
                            Layout.alignment: Qt.AlignHCenter | Qt.AlignVCenter
                            anchors.horizontalCenter: parent.horizontalCenter
                            anchors.verticalCenter: parent.verticalCenter
                            Text { text: "Minutas"; font.pixelSize: 16; color: "#6c6c6c"; horizontalAlignment: Text.AlignHCenter }
                            Text { text: kpis ? kpis.minutas : "N/A"; font.pixelSize: 38; font.bold: true; color: "#232946"; horizontalAlignment: Text.AlignHCenter }
                            Text { text: "No período"; font.pixelSize: 13; color: "#b0b9c6"; horizontalAlignment: Text.AlignHCenter }
                        }
                    }
                    // Card KPI 2
                    Rectangle {
                        width: 240; height: 100; radius: 22
                        color: "#fff"
                        border.color: "#e0e0e0"
                        Column {
                            // Removed invalid Layout.alignment property
                            spacing: 6
                            Text { text: "Média Produtividade"; font.pixelSize: 16; color: "#6c6c6c"; horizontalAlignment: Text.AlignHCenter }
                            Text { text: kpis ? kpis.media : "N/A"; font.pixelSize: 38; font.bold: true; color: "#232946"; horizontalAlignment: Text.AlignHCenter }
                            Text { text: "por usuário"; font.pixelSize: 13; color: "#b0b9c6"; horizontalAlignment: Text.AlignHCenter }
                        }
                    }
                    // Card KPI 3
                    Rectangle {
                        width: 260; height: 100; radius: 22
                        color: "#fff"
                        border.color: "#e0e0e0"
                        Column {
                            spacing: 6
                            Text { text: "Dia Mais Produtivo"; font.pixelSize: 16; color: "#6c6c6c"; horizontalAlignment: Text.AlignHCenter }
                            Text {
                                text: kpis ? kpis.dia : "N/A"
                                font {
                                    pixelSize: 22
                                    bold: true
                                }
                                color: "#232946"
                                horizontalAlignment: Text.AlignHCenter
                                wrapMode: Text.Wrap
                            }
                            Text { 
                                text: "Maior produção diária" 
                                font {
                                    pixelSize: 22
                                }
                                color: "#b0b9c6" 
                                horizontalAlignment: Text.AlignHCenter 
                            }
                        }
                    }
                    // Card KPI 4
                    Rectangle {
                        width: 260; height: 100; radius: 22
                        color: "#fff"
                        border.color: "#e0e0e0"
                        Column {
                            Layout.alignment: Qt.AlignHCenter | Qt.AlignVCenter
                            spacing: 6
                            Text { text: "Top 3 Usuários"; font.pixelSize: 16; color: "#6c6c6c"; horizontalAlignment: Text.AlignHCenter }
                            Text {
                                text: kpis ? kpis.top3 : "N/A"
                                font {
                                    pixelSize: 17
                                    bold: true
                                }
                                color: "#232946"
                                horizontalAlignment: Text.AlignHCenter
                                wrapMode: Text.Wrap
                                elide: Text.ElideRight
                            }
                            Text { text: "Mais produtivos"; font.pixelSize: 13; color: "#b0b9c6"; horizontalAlignment: Text.AlignHCenter }
                        }
                    }
                }

                // Gráfico de barras QtCharts
                Rectangle {
                    // Removed anchors.fill: parent as it conflicts with Layout properties
                    height: 440
                    color: "#f6f8fa"
                    border.color: "transparent"
                    Layout.margins: 32

                // Botão Atualizar gráfico
                Button {
                    text: "Atualizar gráfico"
                    Layout.alignment: Qt.AlignRight | Qt.AlignTop
                    Layout.topMargin: 16
                    Layout.rightMargin: 24
                    width: 140
                    height: 38
                    font {
                        pixelSize: 16
                    }
                    background: Rectangle {
                        color: "#c7c9d9"
                        radius: 18
                    }
                    onClicked: backend.atualizar_grafico()
                    z: 2
                }

                    ChartView {
                        id: chartView
                        Layout.fillWidth: true
                        Layout.fillHeight: true
                        Layout.margins: 24
                        backgroundColor: "transparent"
                        plotAreaColor: "transparent"
                        legend.visible: false
                        antialiasing: true

                        HorizontalBarSeries {
                            BarSet {
                                label: "Minutas"
                                values: backend && backend.valores && backend.valores.length > 0 ? backend.valores : [0]
                                color: "#38a3eb"
                            }
                        }
                        BarCategoryAxis {
                            categories: backend && backend.categorias ? backend.categorias : []
                        }
                        ValueAxis {
                            min: 0
                        max: backend.valores.length > 0 ? Math.max.apply(null, backend.valores.concat(10)) : 10
                            tickCount: 7
                            gridLineColor: "#e0e0e0"
                            labelFormat: "%.0f"
                    }
                }
            }
        }
    }