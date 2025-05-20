import QtQuick 2.15
import QtCharts 2.15

ChartView {
    anchors.fill: parent
    antialiasing: true
    PieSeries {
        PieSlice { label: "A"; value: 30 }
        PieSlice { label: "B"; value: 70 }
    }
}
