import sys
import os
import os

os.environ["QT_QUICK_CONTROLS_STYLE"] = "Fusion"

from PyQt6.QtCore import QLibraryInfo

print("Qt Plugins Path:", QLibraryInfo.path(QLibraryInfo.LibraryPath.PluginsPath))
print("Arquivo QML existe?", os.path.isfile("MainWindow.qml"))
from PyQt6.QtWidgets import QApplication
from PyQt6.QtQml import QQmlApplicationEngine
from PyQt6.QtCore import QObject, pyqtProperty, pyqtSignal, QTimer

app = QApplication(sys.argv)
engine = QQmlApplicationEngine()


class KpiModel(QObject):
    def __init__(self, parent=None):
        super().__init__(parent)


kpi_model = KpiModel()
engine.rootContext().setContextProperty("kpiModel", kpi_model)

engine.load("MainWindow.qml")
app.exec()


class KPIs(QObject):
    kpisChanged = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._minutas = "--"
        self._media = "--"
        self._dia = "--"
        self._top3 = "--"

    @pyqtProperty(str, notify=kpisChanged)
    def minutas(self):
        return self._minutas

    @pyqtProperty(str, notify=kpisChanged)
    def media(self):
        return self._media

    @pyqtProperty(str, notify=kpisChanged)
    def dia(self):
        return self._dia

    @pyqtProperty(str, notify=kpisChanged)
    def top3(self):
        return self._top3

    def atualizar_kpis(self, minutas, media, dia, top3):
        self._minutas = minutas
        self._media = media
        self._dia = dia
        self._top3 = top3
        self.kpisChanged.emit()


if __name__ == "__main__":
    os.environ["QT_QUICK_CONTROLS_STYLE"] = "Fusion"

    from app_v7 import MainWindow

    kpis = KPIs()

    # TESTE: Força valores visíveis nos KPIs ANTES de carregar o QML
    kpis.atualizar_kpis("FORCE", "FORCE", "FORCE", "FORCE")

    engine.rootContext().setContextProperty("kpis", kpis)
    engine.load(os.path.abspath("MainWindow.qml"))
    if not engine.rootObjects():
        sys.exit(-1)

    win = MainWindow(kpis_qml=kpis)
    win.hide()

    QTimer.singleShot(0, win.atualizar_kpis)

    sys.exit(app.exec())
