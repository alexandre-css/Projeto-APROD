from PyQt6.QtCore import QObject, pyqtSignal, pyqtProperty, pyqtSlot, QVariant
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox, QMainWindow, QListWidget
from PyQt6.QtQml import QQmlApplicationEngine
from xls2xlsx import XLS2XLSX
import app_v7
from app_v7 import GerenciadorPesosAgendamento, carregar_planilha
from typing import List
import sys, os, shutil
import pandas as pd
import re

os.environ["QT_QUICK_CONTROLS_STYLE"] = "Material"

class Backend(QObject):
    nomesChanged = pyqtSignal()
    valoresChanged = pyqtSignal()

    def __init__(self, mainwindow: QMainWindow = None):
        super().__init__()
        self._nomes: List[str] = []
        self._valores: List[QVariant] = []
        if mainwindow is not None and isinstance(mainwindow, QMainWindow):
            self.mainwindow: QMainWindow = mainwindow
        else:
            self.mainwindow: QMainWindow = QMainWindow()  # Default to a new QMainWindow instance

    @pyqtProperty('QStringList', notify=nomesChanged)
    def nomes(self):
        return self._nomes

    @pyqtProperty('QVariantList', notify=valoresChanged)
    def valores(self):
        return self._valores

    @pyqtSlot()
    def atualizar_grafico(self):
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")] if os.path.exists(pasta) else []
        dfs = []
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            df = app_v7.carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            if not df.empty:
                dfs.append(df)
        from PyQt6.QtCore import QStringListModel
        nomes_list = []
        valores_list = []
        if dfs:
            df_total = pd.concat(dfs, ignore_index=True)
            if "Usuário" in df_total.columns:
                contagem = df_total["Usuário"].value_counts()
                nomes_list = list(contagem.index)
                valores_list = [int(v) for v in contagem.values]
        
        self._nomes = nomes_list
        self._valores = [QVariant(v) for v in valores_list]
        if isinstance(self._nomes, list) and isinstance(self._valores, list):
            self.nomesChanged.emit()
            self.valoresChanged.emit()

    def abrir_dialogo_excel(self):
        file_dialog = QFileDialog()
        file_dialog.setWindowTitle("Selecione o arquivo Excel")
        file_dialog.setNameFilter("Arquivos Excel (*.xlsx *.xls)")
        if file_dialog.exec():
            selected_file = file_dialog.selectedFiles()[0]
            print("Arquivo selecionado:", selected_file)
    
    @pyqtSlot()
    def importar_arquivo_excel(self):
        self.mainwindow.importar_arquivo_excel()

class KPIs(QObject):
    kpisChanged = pyqtSignal()
    def __init__(self):
        super().__init__()
        self._minutas = "0"
        self._media = "0"
        self._dia = "--"
        self._top3 = "--"

    @pyqtProperty(str, notify=kpisChanged)
    def minutas(self): return self._minutas
    @pyqtProperty(str, notify=kpisChanged)
    def media(self): return self._media
    @pyqtProperty(str, notify=kpisChanged)
    def dia(self): return self._dia
    @pyqtProperty(str, notify=kpisChanged)
    def top3(self): return self._top3

    @pyqtSlot(str, str, str, str)
    def atualizar_kpis(self, minutas, media, dia, top3):
        self._minutas = minutas
        self._media = media
        self._dia = dia
        self._top3 = top3
        self.kpisChanged.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Main Window")

    def importar_arquivo_excel(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Importar arquivos Excel",
            "",
            "Todos Arquivos (*);;Excel Files (*.xlsx *.xls)",
        )
        if paths:
            pasta = "dados_mensais"
            os.makedirs(pasta, exist_ok=True)
            arquivos_importados = 0
            for path in paths:
                try:
                    nome_destino = os.path.basename(path)
                    destino = os.path.join(pasta, nome_destino)
                    if not os.path.exists(destino):
                        shutil.copy2(path, destino)
                        arquivos_importados += 1
                except Exception as e:
                    QMessageBox.warning(
                        self,
                        "Erro",
                        f"Erro ao importar o arquivo '{path}': {str(e)}",
                    )
            if arquivos_importados:
                QMessageBox.information(
                    self,
                    "Importação",
                    f"{arquivos_importados} arquivo(s) importado(s) com sucesso!",
                )
            self.carregar_dados_mensais()
            self.atualizar_listbox_meses()
            self.atualizar_filtros_grafico()
            self.atualizar_kpis()
        else:
            QMessageBox.warning(
                self,
                "Aviso",
                "Nenhum arquivo novo foi importado (todos já existem na pasta de dados).",
            )

    def carregar_dados_mensais(self):
        pasta = "dados_mensais"
        arquivos = (
            [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            if os.path.exists(pasta)
            else []
        )
        dfs = []
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            df = self.carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            df = df.loc[:, ~df.columns.duplicated()]
            if not df.empty:
                dfs.append(df)
        self.df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    def carregar_planilha(self, file_path):
        if file_path.lower().endswith(".xls"):
            x2x = XLS2XLSX(file_path)
            file_path = x2x.to_xlsx()

        header = pd.read_excel(file_path, nrows=0)
        if "peso" in header.columns:
            df = pd.read_excel(file_path, usecols="A:H")
        else:
            df = pd.read_excel(file_path, skiprows=1, usecols="A:G")
            if "Agendamento" in df.columns:
                tipo_limpo = df["Agendamento"].apply(
                    lambda x: re.sub(r"\s*\(.*?\)", "", str(x)).strip()
                )
                gerenciador_pesos = GerenciadorPesosAgendamento()
                df["peso"] = tipo_limpo.apply(gerenciador_pesos.obter_peso)
            else:
                df["peso"] = 1.0
        # Padroniza nome do usuário
        if "Usuário" in df.columns:
            df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
        # Remove colunas duplicadas
        df = df.loc[:, ~df.columns.duplicated()]
        # Garante nome correto da coluna
        for col in df.columns:
            if isinstance(col, str) and col.lower().startswith("usu") and col != "Usuário":
                df.rename(columns={col: "Usuário"}, inplace=True)
        return df
        
    def atualizar_listbox_meses(self):
        pasta = "dados_mensais"
        if not hasattr(self, 'listbox_meses'):
            self.listbox_meses = QListWidget()
            self.setCentralWidget(self.listbox_meses)
        arquivos = (
            [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            if os.path.exists(pasta)
            else []
        )
        opcoes_amigaveis = []
        mapa_arquivo_meses = {}
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            mes_ano = self.extrair_mes_ano_do_arquivo(file_path)
            if mes_ano and "/" in mes_ano:
                opcoes_amigaveis.append(mes_ano)
                mapa_arquivo_meses[mes_ano] = arq
        if hasattr(self, 'listbox_meses'):
            self.listbox_meses.clear()
            for mes in sorted(opcoes_amigaveis, key=self.chave_mes_ano):
                self.listbox_meses.addItem(mes)
        else:
            print("listbox_meses is not defined.")
        self.mapa_arquivo_meses = mapa_arquivo_meses

    def atualizar_filtros_grafico(self):
        if self.df is None or self.df.empty:
            for lb in [
                self.filtro_usuario,
                self.filtro_mes,
                self.filtro_tipo,
                self.filtro_agendamento,
            ]:
                lb.clear()
            return

    def atualizar_kpis(self):
        pasta = "dados_mensais"
        arquivos = (
            [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            if os.path.exists(pasta)
            else []
        )

        dfs = []
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            df = self.carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            df = df.loc[:, ~df.columns.duplicated()]
            if not df.empty:
                dfs.append(df)

        if not dfs:
            # Atualiza QML com valor em branco
            if hasattr(self, "kpis_qml") and self.kpis_qml is not None:
                self.kpis_qml.atualizar_kpis("--", "--", "--", "--")
            return

        df_total = pd.concat(dfs, ignore_index=True)
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]

        # Padroniza coluna Usuário
        for col in df_total.columns:
            if (
                isinstance(col, str)
                and col.lower().startswith("usu")
                and col != "Usuário"
            ):
                df_total.rename(columns={col: "Usuário"}, inplace=True)

        if "Usuário" in df_total.columns:
            df_total["Usuário"] = (
                df_total["Usuário"].astype(str).str.strip().str.upper()
            )

        # Total de minutas
        minutas = len(df_total)
        minutas_qml = str(minutas)

        # Média de produtividade por usuário
        media_qml = "--"
        top3_qml = "--"
        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            produtividade = df_total.groupby("Usuário")["peso"].sum()
            media = produtividade.mean() if not produtividade.empty else 0
            media_qml = f"{media:.2f}"

            top3 = produtividade.sort_values(ascending=False).head(3)
            if not top3.empty:
                top3_qml = "\n".join(
                    [f"{i+1}º {u}: {p:.1f}" for i, (u, p) in enumerate(top3.items())]
                )
            else:
                top3_qml = "--"
        else:
            media_qml = "--"
            top3_qml = "--"

        # Dia mais produtivo
        dia_qml = "--"
        if "Data criação" in df_total.columns and "peso" in df_total.columns:
            df_total["Data criação"] = pd.to_datetime(
                df_total["Data criação"], errors="coerce"
            )
            df_total = df_total.dropna(subset=["Data criação"])
            dias = df_total.groupby(df_total["Data criação"].dt.date)["peso"].sum()
            if not dias.empty:
                dia_mais = dias.idxmax()
                valor_mais = dias.max()
                dia_qml = f"{dia_mais.strftime('%d/%m/%Y')} ({valor_mais:.2f})"
            else:
                dia_qml = "--"
        else:
            dia_qml = "--"

        # Atualiza KPIs no QML
        if hasattr(self, "kpis_qml") and self.kpis_qml is not None:
            self.kpis_qml.atualizar_kpis(minutas_qml, media_qml, dia_qml, top3_qml)


    def extrair_mes_ano_do_arquivo(self, file_path):
        # Extrai mês/ano do arquivo Excel, conforme estrutura padronizada
        try:
            df = carregar_planilha(file_path)
            if "Data criação" in df.columns:
                datas = pd.to_datetime(
                    df["Data criação"], errors="coerce", dayfirst=True
                )
                if not datas.isnull().all():
                    data = datas.dropna().iloc[0]
                    return data.strftime("%b/%Y")
        except Exception as e:
            print(f"Erro: {e}")

    def chave_mes_ano(self, mes_ano):
        try:
            mes, ano = mes_ano.split("/")
            meses = [
                "jan",
                "fev",
                "mar",
                "abr",
                "mai",
                "jun",
                "jul",
                "ago",
                "set",
                "out",
                "nov",
                "dez",
            ]
            return int(ano), meses.index(mes.lower())
        except Exception:
            return (0, 0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("KPI App")
    engine = QQmlApplicationEngine()

    mainwindow = MainWindow()
    backend = Backend(mainwindow)
    kpis = KPIs()
    mainwindow.kpis_qml = kpis

    # REGISTRE OS CONTEXTOS ANTES DO LOAD!
    engine.rootContext().setContextProperty("mainwindow", mainwindow)
    engine.rootContext().setContextProperty("backend", backend)
    engine.rootContext().setContextProperty("kpis", kpis)

    backend.atualizar_grafico()  # Pode chamar aqui, sem problema

    qml_file = os.path.join(os.path.dirname(__file__), "MainWindow.qml")
    if not os.path.exists(qml_file):
        QMessageBox.critical(None, "Error", f"QML file not found at {qml_file}")
        print(f"Expected QML file path: {qml_file}")
        sys.exit(1)

    engine.load(qml_file)
    if not engine.rootObjects():
        QMessageBox.critical(None, "Error", "Failed to load QML file.")
        sys.exit(0)

    sys.exit(app.exec())

