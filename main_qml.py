from PyQt6.QtCore import QObject, pyqtSignal, pyqtProperty, pyqtSlot, QVariant
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox, QMainWindow, QListWidget
from PyQt6.QtQml import QQmlApplicationEngine
from xls2xlsx import XLS2XLSX
from app_v7 import GerenciadorPesosAgendamento, carregar_planilha, MainWindow
from typing import List
import sys, os, shutil
import pandas as pd
import re
import json
from collections import defaultdict

os.environ["QT_QUICK_CONTROLS_STYLE"] = "Material"

MESES_PT = {
    'Jan': 'jan', 'Feb': 'fev', 'Mar': 'mar', 'Apr': 'abr', 'May': 'mai', 'Jun': 'jun',
    'Jul': 'jul', 'Aug': 'ago', 'Sep': 'set', 'Oct': 'out', 'Nov': 'nov', 'Dec': 'dez'
}

class Backend(QObject):
    nomesChanged = pyqtSignal()
    valoresChanged = pyqtSignal()
    arquivosCarregadosChanged = pyqtSignal()
    tabelaPesosChanged = pyqtSignal()

    def __init__(self, mainwindow: QMainWindow = None):
        super().__init__()
        self._nomes: List[str] = []
        self._valores: List[QVariant] = []
        self._arquivosCarregados = ""
        if mainwindow is not None and isinstance(mainwindow, QMainWindow):
            self.mainwindow: QMainWindow = mainwindow
        else:
            self.mainwindow: QMainWindow = QMainWindow()

    @pyqtProperty('QStringList', notify=nomesChanged)
    def nomes(self):
        return self._nomes

    @pyqtProperty('QVariantList', notify=valoresChanged)
    def valores(self):
        return self._valores

    @pyqtProperty('QVariantList', notify=tabelaPesosChanged)
    def tabela_pesos(self):
        return [
            {'tipo': nome, 'peso': valor}
            for nome, valor in zip(self._nomes, self._valores)
        ]

    @pyqtProperty(str, notify=arquivosCarregadosChanged)
    def arquivosCarregados(self):
        return self._arquivosCarregados

    @pyqtSlot()
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
            df = self.mainwindow.carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            df = df.loc[:, ~df.columns.duplicated()]
            if not df.empty:
                dfs.append(df)
        if not dfs:
            if hasattr(self.mainwindow, "kpis_qml") and self.mainwindow.kpis_qml is not None:
                self.mainwindow.kpis_qml.atualizar_kpis("--", "--", "--", "--")
            return
        df_total = pd.concat(dfs, ignore_index=True)
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]
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
        minutas = len(df_total)
        minutas_qml = str(minutas)
        media_qml = "--"
        top3_qml = "--"
        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            produtividade = df_total.groupby("Usuário")["peso"].sum()
            media = produtividade.mean() if not produtividade.empty else 0
            media_qml = f"{media:.2f}"
            top3 = produtividade.sort_values(ascending=False).head(3)
            if not top3.empty:
                top3_qml = "\n".join([f"{u}: {p:.1f}" for u, p in top3.items()])
            else:
                top3_qml = "--"
        elif "Usuário" in df_total.columns:
            contagem = df_total.groupby("Usuário").size()
            top3 = contagem.sort_values(ascending=False).head(3)
            top3_qml = "\n".join([f"{u}: {p}" for u, p in top3.items()])
        if "Data criação" in df_total.columns and "peso" in df_total.columns:
            df_total["Data criação"] = pd.to_datetime(
                df_total["Data criação"], errors="coerce", dayfirst=True
            )
            df_total = df_total.dropna(subset=["Data criação"])
            df_total["dia_semana"] = df_total["Data criação"].dt.dayofweek
            dias_semana = {0: "segunda", 1: "terça", 2: "quarta", 3: "quinta", 4: "sexta", 5: "sábado", 6: "domingo"}
            produtividade_semana = df_total.groupby("dia_semana")["peso"].sum()
            if not produtividade_semana.empty:
                dia_mais = produtividade_semana.idxmax()
                valor_mais = produtividade_semana.max()
                dia_qml = f"{dias_semana[dia_mais].capitalize()} ({valor_mais:.2f})"
            else:
                dia_qml = "--"
        if hasattr(self.mainwindow, "kpis_qml") and self.mainwindow.kpis_qml is not None:
            self.mainwindow.kpis_qml.atualizar_kpis(minutas_qml, media_qml, dia_qml, top3_qml)

    @pyqtSlot()
    def atualizar_arquivos_carregados(self):
        pasta = "dados_mensais"
        if not os.path.exists(pasta):
            self._arquivosCarregados = "Nenhum arquivo carregado"
        else:
            arquivos = sorted([arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")])
            periodos = []
            for arq in arquivos:
                file_path = os.path.join(pasta, arq)
                periodo = self.mainwindow.extrair_mes_ano_do_arquivo(file_path)
                if periodo and "/" in periodo:
                    periodos.append(periodo)
            self._arquivosCarregados = (
                " - ".join(periodos) if periodos else "Nenhum arquivo carregado"
            )
        self.arquivosCarregadosChanged.emit()

    @pyqtSlot()
    def atualizar_grafico(self):
        pasta = "dados_mensais"
        arquivos = (
            [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            if os.path.exists(pasta)
            else []
        )
        dfs = []
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            df = self.mainwindow.carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            df = df.loc[:, ~df.columns.duplicated()]
            if not df.empty:
                dfs.append(df)
        if not dfs:
            self._nomes = []
            self._valores = []
            self.nomesChanged.emit()
            self.valoresChanged.emit()
            self.atualizar_arquivos_carregados()
            return

        df_total = pd.concat(dfs, ignore_index=True)
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]
        for col in df_total.columns:
            if (
                isinstance(col, str)
                and col.lower().startswith("usu")
                and col != "Usuário"
            ):
                df_total.rename(columns={col: "Usuário"}, inplace=True)

        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            produtividade = df_total.groupby("Usuário")["peso"].sum()
            produtividade = produtividade.sort_values(ascending=False)
            self._nomes = list(produtividade.index)
            self._valores = [float(v) for v in produtividade.values]
        elif "Usuário" in df_total.columns:
            produtividade = df_total.groupby("Usuário").size()
            produtividade = produtividade.sort_values(ascending=False)
            self._nomes = list(produtividade.index)
            self._valores = [float(v) for v in produtividade.values]
        else:
            self._nomes = []
            self._valores = []

        self.nomesChanged.emit()
        self.valoresChanged.emit()
        self.atualizar_arquivos_carregados()

    @pyqtSlot()
    def importar_arquivo_excel(self):
        self.mainwindow.importar_arquivo_excel()
        self.atualizar_arquivos_carregados()


    #& Aqui começam os métodos de atribuição de pesos.

    @pyqtSlot()
    def atualizar_tabela_pesos(self):
        """
        Popula self._nomes com os tipos de agendamento únicos encontrados nos arquivos Excel
        e self._valores com o peso atribuído a cada tipo (via GerenciadorPesosAgendamento).
        """
        pasta = "dados_mensais"
        arquivos = (
            [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            if os.path.exists(pasta)
            else []
        )
        tipos_agendamento = set()
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            df = self.mainwindow.carregar_planilha(file_path)
            if "Agendamento" in df.columns:
                tipos = df["Agendamento"].dropna().astype(str).str.strip()
                tipos_agendamento.update(tipos)
        # Ordena alfabeticamente para exibir sempre igual
        tipos_agendamento = sorted(tipos_agendamento)
        # Pega os pesos atuais do gerenciador
        pesos = [
            self.mainwindow.gerenciador_pesos.obter_peso(tipo)
            for tipo in tipos_agendamento
        ]
        self._nomes = list(tipos_agendamento)
        self._valores = pesos
        self.nomesChanged.emit()
        self.valoresChanged.emit()

    @pyqtSlot(int, float)
    def atualizarPeso(self, index, novo_peso):
        """
        Atualiza o peso de um tipo de agendamento, dado o índice na lista de nomes.
        """
        if hasattr(self.mainwindow, "gerenciador_pesos"):
            nome = self.nomes[index] if 0 <= index < len(self.nomes) else None
            if nome:
                self.mainwindow.gerenciador_pesos.atualizar_peso(nome, novo_peso)
                self.mainwindow.gerenciador_pesos.salvar_pesos()
                self.atualizar_grafico()
                self.atualizar_kpis()
                self.nomesChanged.emit()
                self.valoresChanged.emit()
    
    @pyqtSlot()
    def salvarPesos(self):
        """
        Salva os pesos atuais no arquivo JSON.
        """
        if hasattr(self.mainwindow, "gerenciador_pesos"):
            self.mainwindow.gerenciador_pesos.salvar_pesos()
            self.atualizar_tabela_pesos()
    @pyqtSlot()
    def carregarPesos(self):
        """
        Carrega pesos do arquivo JSON e atualiza a tabela.
        """
        if hasattr(self.mainwindow, "gerenciador_pesos"):
            self.mainwindow.gerenciador_pesos.carregar_pesos()
            self.atualizar_grafico()
            self.atualizar_kpis()
            self.nomesChanged.emit()
            self.valoresChanged.emit()
            self.atualizar_tabela_pesos()
    @pyqtSlot()
    def restaurarPesosPadrao(self):
        """
        Restaura todos os pesos para 1.0 e salva.
        """
        if hasattr(self.mainwindow, "gerenciador_pesos"):
            self.mainwindow.gerenciador_pesos.pesos = defaultdict(lambda: 1.0)
            self.mainwindow.gerenciador_pesos.salvar_pesos()
            self.atualizar_grafico()
            self.atualizar_kpis()
            self.nomesChanged.emit()
            self.valoresChanged.emit()
            self.atualizar_tabela_pesos()

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
        self.gerenciador_pesos = GerenciadorPesosAgendamento()

    def chave_mes_ano(self, mes_ano):
        try:
            mes, ano = mes_ano.split("/")
            meses = [
                "jan", "fev", "mar", "abr", "mai", "jun",
                "jul", "ago", "set", "out", "nov", "dez"
            ]
            return int(ano), meses.index(mes.lower())
        except Exception:
            return (0, 0)

    def extrair_mes_ano_do_arquivo(self, file_path):
        try:
            df = carregar_planilha(file_path)
            if "Data criação" in df.columns and not df.empty:
                #&@ Converte explicitamente para datetime, sempre com dayfirst=True
                datas = pd.to_datetime(df["Data criação"], errors="coerce", dayfirst=True)
                datas = datas.dropna()
                if not datas.empty:
                    #& Usa o mês/ano mais frequente (modo), que é o mais representativo do arquivo
                    mes_anos = datas.dt.strftime("%m/%Y")
                    mes_ano_mais_frequente = mes_anos.mode()[0]
                    mes, ano = mes_ano_mais_frequente.split("/")
                    #& Mapeamento para português
                    MESES_PT = [
                        "jan", "fev", "mar", "abr", "mai", "jun",
                        "jul", "ago", "set", "out", "nov", "dez"
                    ]
                    mes_pt = MESES_PT[int(mes) - 1]
                    return f"{mes_pt}/{ano}"
        except Exception as e:
            print(f"Erro ao extrair mês/ano: {e}")
        #&& Fallback: retorna o nome do arquivo sem extensão
        return os.path.basename(file_path).replace(".xlsx", "")

    def importar_arquivo_excel(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Importar arquivos Excel",
            "",
            "Todos Arquivos (*);;Excel Files (*.xlsx *.xls)",
        )
        if not paths:
            # Usuário cancelou, não faz nada
            return

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
        else:
            QMessageBox.warning(
                self,
                "Aviso",
                "Nenhum arquivo novo foi importado (todos já existem na pasta de dados).",
            )
        self.carregar_dados_mensais()
        self.atualizar_listbox_meses()
        self.atualizar_filtros_grafico()
        self.atualizar_kpis()
        if hasattr(self, "backend_qml") and self.backend_qml is not None:
            self.backend_qml.atualizar_grafico()
            self.backend_qml.atualizar_tabela_pesos()
            self.backend_qml.atualizar_kpis()
            self.backend_qml.atualizar_arquivos_carregados()
            self.backend_qml.atualizar_filtros_grafico()

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
                df["peso"] = tipo_limpo.apply(self.gerenciador_pesos.obter_peso)
            else:
                df["peso"] = 1.0
        if "Usuário" in df.columns:
            df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
        df = df.loc[:, ~df.columns.duplicated()]
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
        if getattr(self, 'df', None) is None or self.df.empty:
            for lb in [
                getattr(self, 'filtro_usuario', None),
                getattr(self, 'filtro_mes', None),
                getattr(self, 'filtro_tipo', None),
                getattr(self, 'filtro_agendamento', None),
            ]:
                if lb:
                    lb.clear()
            return

class GerenciadorPesosAgendamento:
    """
    Gerencia os pesos atribuídos a cada tipo de agendamento.
    Os pesos são persistidos em um arquivo JSON.
    """
    def __init__(self, arquivo_pesos="pesos.json"):
        self.arquivo_pesos = arquivo_pesos
        self.pesos = defaultdict(lambda: 1.0)
        self.carregar_pesos()

    def carregar_pesos(self):
        """
        Carrega os pesos do arquivo JSON. Se não existir, inicializa tudo com 1.0.
        """
        try:
            with open(self.arquivo_pesos, "r", encoding="utf-8") as f:
                dados = json.load(f)
                self.pesos.clear()
                self.pesos.update(dados)
        except Exception:
            self.pesos = defaultdict(lambda: 1.0)
            self.salvar_pesos()

    def salvar_pesos(self):
        """
        Salva os pesos atuais no arquivo JSON.
        """
        with open(self.arquivo_pesos, "w", encoding="utf-8") as f:
            json.dump(dict(self.pesos), f, indent=4)

    def atualizar_peso(self, tipo_agendamento, novo_peso):
        """
        Atualiza o peso de um tipo de agendamento e salva imediatamente.
        """
        try:
            self.pesos[tipo_agendamento] = float(novo_peso)
            self.salvar_pesos()
            return True
        except ValueError:
            return False

    def obter_peso(self, tipo_agendamento):
        """
        Retorna o peso do tipo de agendamento informado (padrão: 1.0).
        """
        return self.pesos.get(tipo_agendamento, 1.0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("KPI App")
    engine = QQmlApplicationEngine()

    mainwindow = MainWindow()
    backend = Backend(mainwindow)
    kpis = KPIs()
    mainwindow.kpis_qml = kpis
    mainwindow.backend_qml = backend

    engine.rootContext().setContextProperty("mainwindow", mainwindow)
    engine.rootContext().setContextProperty("backend", backend)
    engine.rootContext().setContextProperty("kpis", kpis)

    backend.atualizar_grafico()
    backend.atualizar_kpis()
    backend.arquivosCarregados

    qml_file = os.path.join(os.path.dirname(__file__), "MainWindow.qml")
    if not os.path.exists(qml_file):
        QMessageBox.critical(None, "Error", f"QML file not found at {qml_file}")
        sys.exit(1)

    engine.load(qml_file)
    if not engine.rootObjects():
        QMessageBox.critical(None, "Error", "Failed to load QML file.")
        sys.exit(0)
    sys.exit(app.exec())

