import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.15

import QtQuick 2.15
import QtQuick.Controls 2.15
import QtQuick.Layouts 1.15


ApplicationWindow {
    id: mainWindow
    width: 1200
    height: 800
    visible: true
    title: "Dashboard de Produtividade"

    Rectangle {
        anchors.fill: parent
        color: "#f4f5f8"

        RowLayout {
            anchors.fill: parent

            // Sidebar
            Rectangle {
                width: 220
                color: "#1a232e"
                Layout.fillHeight: true
                ColumnLayout {
                    anchors.fill: parent
                    spacing: 12
                    // Logo
                    Image {
                        source: "logo.png"
                        width: 38
                        height: 38
                        anchors.horizontalCenter: parent.horizontalCenter
                        fillMode: Image.PreserveAspectFit
                        // Adapte o caminho do logo conforme necessário
                    }
                    // Menu
                    Repeater {
                        model: ["Home", "Audicance", "Analysis", "Retention", "Conversion", "Pages", "Events", "Behaviour", "Monetization", "Settings"]
                        delegate: Button {
                            text: modelData
                            font.pixelSize: 16
                            background: Rectangle {
                                color: hovered ? "#232f3e" : "transparent"
                                radius: 6
                            }
                            anchors.horizontalCenter: parent.horizontalCenter
                            width: parent.width * 0.85
                            height: 38
                            leftPadding: 16
                            rightPadding: 16
                            font.bold: text === "Home"
                            contentItem: Text {
                                text: parent.text
                                color: text === "Home" ? "#fff" : "#b0b9c6"
                                font.pixelSize: 16
                                font.bold: text === "Home"
                            }
                        }
                    }
                    // Espaço para expandir
                    Rectangle { Layout.fillHeight: true; color: "transparent" }
                }
            }

            // Conteúdo principal
            ColumnLayout {
                Layout.fillWidth: true
                Layout.fillHeight: true
                spacing: 0

                // Barra superior
                Rectangle {
                    color: "#fff"
                    height: 60
                    Layout.fillWidth: true
                    RowLayout {
                        anchors.fill: parent
                        spacing: 16
                        TextField {
                            placeholderText: "Search here ..."
                            width: 220
                            height: 36
                            font.pixelSize: 15
                            background: Rectangle { color: "#f4f5f8"; radius: 8 }
                        }
                        Text {
                            text: "Home"
                            font.pixelSize: 28
                            font.bold: true
                            color: "#1a232e"
                            verticalAlignment: Text.AlignVCenter
                        }
                    }
                }

                // Cards de KPIs
                RowLayout {
                    Layout.fillWidth: true
                    spacing: 24
                    anchors.margins: 32
                    anchors.leftMargin: 0
                    anchors.rightMargin: 0

                    // Exemplo de um card de KPI funcional
                    Repeater {
                        model: [
                            {titulo: "Minutas", valor: kpiModel.minutas, descricao: "Minutas produzidas no período"},
                            {titulo: "Média Produtividade", valor: kpiModel.media, descricao: "Média por usuário"},
                            {titulo: "Dia Mais Produtivo", valor: kpiModel.dia, descricao: "Maior produção diária"},
                            {titulo: "Top 3 Usuários", valor: kpiModel.top3, descricao: "Usuários mais produtivos"}
                        ]
                        delegate: Rectangle {
                            width: 200
                            height: 110
                            radius: 16
                            color: "#fff"
                            border.color: "#e0e0e0"
                            ColumnLayout {
                                anchors.centerIn: parent
                                spacing: 2
                                Text { text: model.titulo; font.pixelSize: 15; color: "#6c6c6c" }
                                Text { text: model.valor; font.pixelSize: 32; font.bold: true; color: "#1a232e" }
                                Text { text: model.descricao; font.pixelSize: 12; color: "#aaaaaa" }
                            }
                        }
                    }
                }

                // Área de gráficos e tabelas
                Rectangle {
                    Layout.fillWidth: true
                    Layout.fillHeight: true
                    color: "#fff"
                    radius: 14
                    anchors.margins: 36

                    Loader {
                        id: conteudoLoader
                        anchors.fill: parent
                        source: "GraficoView.qml"  // Troque dinamicamente conforme a navegação
                    }
                }
            }
        }
    }
}
