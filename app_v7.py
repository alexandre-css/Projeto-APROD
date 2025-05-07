from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QPixmap, QColor, QIcon, QPalette, QAction
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QTabWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QListWidget,
    QFileDialog,
    QGroupBox,
    QComboBox,
    QCheckBox,
    QHeaderView,
    QFrame,
    QAbstractItemView,
    QMessageBox,
    QSplashScreen,
    QGraphicsDropShadowEffect,
    QToolBar,
    QSizePolicy,
)

import sys
import os
import pandas as pd
import calendar
import tempfile
import shutil
import re
import json
import unicodedata
import numpy as np
import time
import matplotlib

matplotlib.use("QtAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

plt.close("all")
from collections import defaultdict
from xls2xlsx import XLS2XLSX

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

app = QApplication.instance()
QApplication.setFont(QFont("Georgia, Segoe UI, Arial", 12))
if not app:
    app = QApplication(sys.argv)


def carregar_planilha(file_path):
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


def formatar_valor(v):
    try:
        if isinstance(v, pd.Series):
            v = v.iloc[0] if not v.empty else 0.0
        return f"{float(v):.2f}"
    except Exception:
        return "0.00"


STYLE_OPTIONS = {
    "classic": "classic",
    "dark_background": "dark_background",
    "fivethirtyeight": "fivethirtyeight",
    "ggplot": "ggplot",
    "grayscale": "grayscale",
    "bmh": "bmh",
    "Solarize_Light2": "Solarize_Light2",
    "tableau-colorblind10": "tableau-colorblind10",
    "seaborn": "seaborn-v0_8",
}


class GerenciadorPesosAgendamento:
    def __init__(self, arquivo_pesos="pesos.json"):
        self.arquivo_pesos = arquivo_pesos
        self.pesos = defaultdict(lambda: 1.0)
        self.carregar_pesos()

    def carregar_pesos(self):
        try:
            with open(self.arquivo_pesos, "r", encoding="utf-8") as f:
                dados = json.load(f)
                self.pesos.clear()
                self.pesos.update(dados)
        except Exception:
            self.pesos = defaultdict(lambda: 1.0)
            self.salvar_pesos()

    def salvar_pesos(self):
        with open(self.arquivo_pesos, "w", encoding="utf-8") as f:
            json.dump(dict(self.pesos), f, indent=4)

    def atualizar_peso(self, tipo_agendamento, novo_peso):
        try:
            self.pesos[tipo_agendamento] = float(novo_peso)
            self.salvar_pesos()
            return True
        except ValueError:
            return False

    def obter_peso(self, tipo_agendamento):
        return self.pesos.get(tipo_agendamento, 1.0)


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
        super().__init__(fig)


class MainWindow(QMainWindow):
    def __init__(self, kpis_qml=None):
        super().__init__()
        self.kpis_qml = kpis_qml
        if app:
            app.setStyleSheet(
                app.styleSheet()
                + """
                QTableWidget::item:selected, QTableView::item:selected {
                    background: rgba(189, 158, 87, 0.25);  /* dourado claro com transparência */
                    color: #1a232e;
                }
                QTableWidget::item:focus, QTableView::item:focus {
                    outline: none;
                }
            """
            )

        self.setWindowTitle("Sistema de Análise e Produtividade")
        self.resize(1200, 800)
        self.df = None
        self.gerenciador_pesos = GerenciadorPesosAgendamento()
        self.tabs = QTabWidget()

        palette = self.palette()
        palette.setColor(
            self.backgroundRole(), QColor("#dad9d5")
        )  # tom de cinza quente
        self.setPalette(palette)
        self.setAutoFillBackground(True)

        # Barra de ferramentas (toolbar) com botão de importação
        toolbar = QToolBar("Barra de ferramentas")
        toolbar.setMovable(False)
        toolbar.setStyleSheet(
            """
            QToolBar {
                background: #f7f7f7;
                border-bottom: 1px solid #d1d5db;
                spacing: 8px;
                padding: 4px 12px;
            }
        """
        )
        icon_upload = QIcon.fromTheme("document-open")
        btn_importar = QPushButton("Importar Excel")
        btn_importar.setStyleSheet(
            """
            QPushButton {
                background-color: #1a232e;
                color: white;
                border-radius: 8px;
                padding: 8px 24px;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #a97b50;
                color: #fff;
            }
            """
        )
        btn_importar.clicked.connect(self.importar_arquivo_excel)
        toolbar.addWidget(btn_importar)

        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, toolbar)

        # Painel de KPIs moderno e centralizado
        self.kpi_panel = QWidget()
        kpi_layout = QHBoxLayout()
        kpi_layout.setSpacing(24)
        kpi_layout.setContentsMargins(24, 24, 24, 12)
        self.kpiwidgets = {}

        for title, key in [
            ("Minutas", "minutas"),
            ("Média Produtividade", "media"),
            ("Dia Mais Produtivo", "dia"),
            ("Top 3 Usuários", "top3"),
        ]:
            card = QFrame()
            card.setStyleSheet(
                """
                QFrame {
                    background: #fff;
                    border: 2px solid #a97b50;
                    border-radius: 10px;
                    min-width: 170px;
                    max-width: 220px;
                }
                """
            )
            vbox = QVBoxLayout()
            vbox.setSpacing(2)
            vbox.setContentsMargins(12, 12, 12, 12)
            lbl_value = QLabel("--")
            lbl_value.setStyleSheet(
                "font-size: 22px; font-weight: bold; color: #1a232e;"
            )
            lbl_value.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl_value.setWordWrap(True)
            lbl_title = QLabel(title)
            lbl_title.setStyleSheet(
                "color: #1a232e; font-size: 13px; font-weight: 400;"
            )
            lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl_title.setWordWrap(True)
            vbox.addWidget(lbl_value)
            vbox.addWidget(lbl_title)
            card.setLayout(vbox)
            kpi_layout.addWidget(card)
            self.kpiwidgets[key] = lbl_value

        kpi_layout.addStretch()
        self.kpi_panel.setLayout(kpi_layout)

        # Layout principal
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(24, 18, 24, 18)
        main_layout.setSpacing(18)
        main_layout.addWidget(self.kpi_panel)
        main_layout.addWidget(self.tabs)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # Abas
        self.tab_graficos = QWidget()
        self.tab_comparacao_meses = QWidget()
        self.tab_config = QWidget()
        self.tabs.addTab(self.tab_graficos, "Análise e Gráficos")
        self.tabs.addTab(self.tab_comparacao_meses, "Comparação de Meses")
        self.tabs.addTab(self.tab_config, "Atribuir Pesos")

        self.setup_tab_graficos()
        self.setup_tab_comparacao_meses()
        self.setup_tab_config()
        self.carregar_dados_mensais()
        self.atualizar_listbox_meses()
        self.atualizar_filtros_grafico()
        self.atualizar_kpis()

        self._atualizando_grafico = False
        self._canvas_valido = True
        self._debounce_timer = QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self._atualizar_grafico_debounced)
        self.tabs.currentChanged.connect(self.tab_changed)

    def __del__(self):
        self._debounce_timer.stop()
        self._canvas_valido = False
        try:
            self.combo_tipo_grafico.currentIndexChanged.disconnect()
            self.combo_design_grafico.currentIndexChanged.disconnect()
            self.check_cores_diferentes.stateChanged.disconnect()
            self.filtro_usuario.itemSelectionChanged.disconnect()
            self.filtro_mes.itemSelectionChanged.disconnect()
            self.filtro_tipo.itemSelectionChanged.disconnect()
            self.filtro_agendamento.itemSelectionChanged.disconnect()
        except Exception:
            pass

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
                nome_destino = os.path.basename(path)
                destino = os.path.join(pasta, nome_destino)
                if not os.path.exists(destino):
                    shutil.copy2(path, destino)
                    arquivos_importados += 1
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

    def setup_tab_graficos(self):
        main_layout = QHBoxLayout()
        main_layout.setSpacing(18)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Painel de Filtros (esquerda)
        opcoes_box = QGroupBox("Filtros")
        opcoes_box.setStyleSheet(
            """
            QGroupBox {
                background: #f7f7f7;
                border: 2px solid #a97b50;
                border-radius: 10px;
                font-size: 16px;
                margin-top: 0px;
                margin-bottom: 0px;
                padding-top: 12px;
            }
            QGroupBox::title {
                color: #1a232e;
                font-weight: bold;
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 8px;
                padding: 0 4px 0 4px;
                background: transparent;
            }
        """
        )
        opcoes_layout = QVBoxLayout()
        opcoes_layout.setSpacing(14)
        opcoes_layout.setContentsMargins(16, 12, 16, 18)

        tipo_design_layout = QHBoxLayout()
        tipo_design_layout.setSpacing(8)

        tipo_col = QVBoxLayout()
        lbl_tipo = QLabel("Tipo de Gráfico:")
        self.combo_tipo_grafico = QComboBox()
        self.combo_tipo_grafico.addItems(["Colunas Horizontais", "Colunas Verticais"])
        self.combo_tipo_grafico.setCurrentIndex(0)
        tipo_col.addWidget(lbl_tipo)
        tipo_col.addWidget(self.combo_tipo_grafico)

        design_col = QVBoxLayout()
        lbl_design = QLabel("Design:")
        self.combo_design_grafico = QComboBox()
        self.combo_design_grafico.addItems(
            [
                "seaborn-v0_8",
                "tableau-colorblind10",
                "ggplot",
                "Solarize_Light2",
                "classic",
                "dark_background",
            ]
        )
        self.combo_design_grafico.setCurrentText("seaborn-v0_8")
        design_col.addWidget(lbl_design)
        design_col.addWidget(self.combo_design_grafico)

        tipo_design_layout.addLayout(tipo_col)
        tipo_design_layout.addSpacing(10)
        tipo_design_layout.addLayout(design_col)
        tipo_design_layout.addStretch()
        opcoes_layout.addLayout(tipo_design_layout)

        self.check_cores_diferentes = QCheckBox("Colunas em cores diferentes")
        self.check_cores_diferentes.setChecked(True)
        opcoes_layout.addWidget(self.check_cores_diferentes)
        opcoes_layout.addSpacing(8)

        # Filtros
        for label, attr in [
            ("Usuário:", "filtro_usuario"),
            ("Mês:", "filtro_mes"),
            ("Tipo:", "filtro_tipo"),
            ("Agendamento:", "filtro_agendamento"),
        ]:
            lbl = QLabel(label)
            lbl.setStyleSheet("font-weight: 500; margin-top: 4px; margin-bottom: 2px;")
            opcoes_layout.addWidget(lbl)
            lb = QListWidget()
            lb.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
            lb.setStyleSheet(
                """
                QListWidget {
                    background: #f5f5f5;
                    border-radius: 6px;
                    font-size: 13px;
                    padding: 4px;
                    border: 1px solid #a97b50;
                    color: #1a232e;
                    selection-background-color: rgba(189, 158, 87, 0.25);
                    selection-color: #1a232e;
                }
            """
            )
            setattr(self, attr, lb)
            opcoes_layout.addWidget(lb)

        opcoes_layout.addStretch()

        btn_graph = QPushButton("Gerar Gráfico")
        btn_graph.setStyleSheet(
            """
            QPushButton {
                background-color: #1a232e;
                color: white;
                border-radius: 8px;
                padding: 10px 24px;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #a97b50;
                color: #fff;
            }
        """
        )
        btn_graph.clicked.connect(self.gerar_grafico_com_filtros)
        opcoes_layout.addWidget(btn_graph, alignment=Qt.AlignmentFlag.AlignBottom)
        opcoes_box.setLayout(opcoes_layout)
        main_layout.addWidget(opcoes_box, 1)

        # Painel do Gráfico (direita) - IGUAL ao de Filtros!
        grafico_box = QGroupBox("Gráfico")
        grafico_box.setStyleSheet(
            """
            QGroupBox {
                background: #f7f7f7;
                border: 2px solid #a97b50;
                border-radius: 10px;
                font-size: 16px;
                margin-top: 0px;
                margin-bottom: 0px;
                padding-top: 12px;
            }
            QGroupBox::title {
                color: #1a232e;
                font-weight: bold;
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 8px;
                padding: 0 4px 0 4px;
                background: transparent;
            }
        """
        )
        grafico_layout = QVBoxLayout()
        grafico_layout.setContentsMargins(16, 12, 16, 18)
        self.grafico_label = QLabel()
        self.grafico_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.grafico_label.setStyleSheet(
            "background: #fff; border: none; border-radius: 10px;"
        )
        grafico_layout.addWidget(self.grafico_label)
        grafico_box.setLayout(grafico_layout)
        main_layout.addWidget(grafico_box, 3)

        self.tab_graficos.setLayout(main_layout)

        # Ligações dos filtros e combos
        self.combo_tipo_grafico.currentIndexChanged.connect(self._debounce_grafico)
        self.combo_design_grafico.currentIndexChanged.connect(self._debounce_grafico)
        self.check_cores_diferentes.stateChanged.connect(self._debounce_grafico)
        self.filtro_usuario.itemSelectionChanged.connect(self._debounce_grafico)
        self.filtro_mes.itemSelectionChanged.connect(self._debounce_grafico)
        self.filtro_tipo.itemSelectionChanged.connect(self._debounce_grafico)
        self.filtro_agendamento.itemSelectionChanged.connect(self._debounce_grafico)
        self._canvas_valido = True

    def _debounce_grafico(self):
        if self._canvas_valido:
            self._debounce_timer.start(200)  # 200 ms

    def setup_tab_comparacao_meses(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Selecione os meses para comparar:"))

        self.listbox_meses = QListWidget()
        self.listbox_meses.setSelectionMode(
            QAbstractItemView.SelectionMode.MultiSelection
        )
        self.listbox_meses.setMaximumHeight(110)
        self.listbox_meses.setStyleSheet(
            """
            QListWidget {
                background: #f7f7f7;
                border-radius: 6px;
                font-size: 13px;
                padding: 4px;
                border: 2px solid #a97b50;
                color: #1a232e;
                selection-background-color: rgba(189, 158, 87, 0.25);
                selection-color: #1a232e;
            }
            """
        )

        layout.addWidget(self.listbox_meses)

        btns = QHBoxLayout()
        btn_comparar = QPushButton("Comparar")
        btn_comparar.setStyleSheet(
            """
            QPushButton {
                background-color: #1a232e;
                color: white;
                border-radius: 8px;
                padding: 8px 24px;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #a97b50;
                color: #fff;
            }
            """
        )
        btn_comparar.clicked.connect(self.comparar_meses)
        btns.addWidget(btn_comparar)
        layout.addLayout(btns)

        self.table_comp = QTableWidget()
        self.table_comp.setStyleSheet(
            """
            QTableWidget {
                background: #fff;
                font-size: 14px;
                border: 2px solid #a97b50;
            }
            QHeaderView::section {
                background: #a97b50;
                color: white;
                font-weight: bold;
                border: none;
                font-size: 15px;
            }
            QTableWidget::item {
                border-bottom: 1px solid #d1d5db;
            }
        """
        )

        self.table_comp.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.table_comp.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        layout.addWidget(self.table_comp)

        self.tab_comparacao_meses.setLayout(layout)
        self.atualizar_listbox_meses()

        btn_comparar_todos = QPushButton("Comparar Todos")
        btn_comparar_todos.setStyleSheet(
            """
            QPushButton {
                background-color: #a97b50;
                color: white;
                border-radius: 8px;
                padding: 8px 24px;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #1a232e;
                color: #fff;
            }
            """
        )
        btn_comparar_todos.clicked.connect(self.comparar_todos_meses)
        btns.addWidget(btn_comparar_todos)

        btn_excluir = QPushButton("Excluir dados dos meses")
        btn_excluir.setStyleSheet(
            """
            QPushButton {
                background-color: #d32f2f;
                color: white;
                border-radius: 8px;
                padding: 8px 24px;
                font-size: 15px;
                font-weight: 600;
            }
            QPushButton:hover {
                background-color: #a97b50;
                color: #fff;
            }
            """
        )
        btn_excluir.clicked.connect(self.limpar_dados_anteriores)
        btns.addWidget(btn_excluir)

    def atualizar_listbox_meses(self):
        pasta = "dados_mensais"
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
        self.listbox_meses.clear()
        for mes in sorted(opcoes_amigaveis, key=self.chave_mes_ano):
            self.listbox_meses.addItem(mes)
        self.mapa_arquivo_meses = mapa_arquivo_meses

    def comparar_meses(self):
        selecionados = [item.text() for item in self.listbox_meses.selectedItems()]
        if len(selecionados) < 2:
            QMessageBox.warning(
                self, "Aviso", "Selecione pelo menos dois meses para comparar!"
            )
            return

        pasta = "dados_mensais"
        dados_meses = []
        for mes in selecionados:
            arquivo = self.mapa_arquivo_meses.get(mes)
            if not arquivo:
                continue
            file_path = os.path.join(pasta, arquivo)
            df = carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip()
            col_usuario = "Usuário"
            col_agendamento = "Agendamento"
            if "peso" not in df.columns:
                if col_agendamento in df.columns:
                    tipo_limpo = df[col_agendamento].apply(
                        lambda x: re.sub(r"\s*\(.*?\)", "", str(x)).strip()
                    )
                    df["peso"] = tipo_limpo.apply(self.gerenciador_pesos.obter_peso)
                else:
                    df["peso"] = 1.0
            if col_usuario in df.columns and "peso" in df.columns:
                prod = df.groupby(col_usuario)["peso"].sum()
            else:
                prod = pd.Series(dtype=float)
            dados_meses.append(prod)

        todos_usuarios = sorted(
            {usuario for prod in dados_meses for usuario in prod.index}
        )
        colunas = ["Usuário"] + selecionados + ["Total"]
        self.table_comp.setColumnCount(len(colunas))
        self.table_comp.setHorizontalHeaderLabels(colunas)
        self.table_comp.setRowCount(len(todos_usuarios))

        for row_idx, usuario in enumerate(todos_usuarios):
            linha = [usuario]
            total = 0.0
            for prod in dados_meses:
                valor = float(prod.get(usuario, 0.0))
                linha.append(f"{valor:.2f}")
                total += valor
            linha.append(f"{total:.2f}")
            for col_idx, valor in enumerate(linha):
                item = QTableWidgetItem(str(valor))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table_comp.setItem(row_idx, col_idx, item)

    def preencher_listbox_meses(self, lista_meses):
        self.listbox_meses.clear()
        for mes in lista_meses:
            self.listbox_meses.addItem(str(mes))

    def preencher_tabela_tjsc(self, dados):
        """
        Preenche a QTableWidget self.table_comp com os dados do TJSC,
        colorindo a coluna 'Status' conforme o significado jurídico.
        """
        # Defina os status e cores relevantes do TJSC:
        STATUS_CORES = {
            "Anexada ao processo": "#388e3c",  # Verde
            "Enviada para jurisprudência": "#bfa046",  # Dourado
            "Para assinar": "#1976d2",  # Azul
            "Para conferir": "#1976d2",  # Azul
            "Conferida": "#388e3c",  # Verde
            "Remetidos os Autos": "#1976d2",  # Azul
            "Cancelada": "#d32f2f",  # Vermelho
            "Extinto o processo por desistência": "#d32f2f",
            "Processo Suspenso": "#bfa046",
            "Gratuidade da justiça não concedida": "#d32f2f",
            "Indeferido o pedido": "#d32f2f",
            "Homologada a Desistência do Recurso": "#bfa046",
            # Adicione outros status relevantes conforme sua base de dados
        }

        self.table_comp.setRowCount(len(dados))
        for row_idx, row_data in enumerate(dados):
            for col_idx, (col, valor) in enumerate(row_data.items()):
                item = QTableWidgetItem(str(valor))
                if col.lower() == "status":
                    cor = None
                    for chave, hexcor in STATUS_CORES.items():
                        if chave.lower() in str(valor).lower():
                            cor = hexcor
                            break
                    if cor:
                        item.setForeground(QColor(cor))
                    else:
                        item.setForeground(QColor("#1a232e"))  # Padrão sóbrio
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table_comp.setItem(row_idx, col_idx, item)

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
            df = carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            df = df.loc[:, ~df.columns.duplicated()]
            if not df.empty:
                dfs.append(df)

        if not dfs:
            for lbl in self.kpiwidgets.values():
                lbl.setText("--")
            # Atualiza QML também!
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

        minutas = len(df_total)
        self.kpiwidgets["minutas"].setText(str(minutas))

        minutas_qml = str(minutas)
        media_qml = "--"
        dia_qml = "--"
        top3_qml = "--"

        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            produtividade = df_total.groupby("Usuário")["peso"].sum()
            media = produtividade.mean() if not produtividade.empty else 0
            self.kpiwidgets["media"].setText(f"{media:.2f}")
            media_qml = f"{media:.2f}"

            top3 = produtividade.sort_values(ascending=False).head(3)
            if not top3.empty:
                top3_txt = "\n".join(
                    [f"{i+1}º {u}: {p:.1f}" for i, (u, p) in enumerate(top3.items())]
                )
                self.kpiwidgets["top3"].setText(top3_txt)
                top3_qml = top3_txt
                if len(top3_txt) > 25:
                    self.kpiwidgets["top3"].setStyleSheet(
                        "font-size: 13px; color: #1a232e; font-weight: bold;"
                    )
                else:
                    self.kpiwidgets["top3"].setStyleSheet(
                        "font-size: 18px; color: #1a232e; font-weight: bold;"
                    )
            else:
                self.kpiwidgets["top3"].setText("--")
                top3_qml = "--"
        else:
            self.kpiwidgets["media"].setText("--")
            self.kpiwidgets["top3"].setText("--")
            media_qml = "--"
            top3_qml = "--"

        if "Data criação" in df_total.columns and "peso" in df_total.columns:
            df_total["Data criação"] = pd.to_datetime(
                df_total["Data criação"], errors="coerce"
            )
            df_total = df_total.dropna(subset=["Data criação"])
            dias = df_total.groupby(df_total["Data criação"].dt.date)["peso"].sum()
            if not dias.empty:
                dia_mais = dias.idxmax()
                valor_mais = dias.max()
                dia_str = f"{dia_mais.strftime('%d/%m/%Y')} ({valor_mais:.2f})"
                self.kpiwidgets["dia"].setText(dia_str)
                dia_qml = dia_str
            else:
                self.kpiwidgets["dia"].setText("--")
                dia_qml = "--"
        else:
            self.kpiwidgets["dia"].setText("--")
            dia_qml = "--"

        # Atualiza KPIs no QML
        if hasattr(self, "kpis_qml") and self.kpis_qml is not None:
            minutas_qml = self.kpiwidgets["minutas"].text()
            media_qml = self.kpiwidgets["media"].text()
            dia_qml = self.kpiwidgets["dia"].text()
            top3_qml = self.kpiwidgets["top3"].text()
            self.kpis_qml.atualizar_kpis(minutas_qml, media_qml, dia_qml, top3_qml)

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

        col_usuario = "Usuário"
        if col_usuario not in self.df.columns:
            col_usuario = next(
                (
                    c
                    for c in self.df.columns
                    if isinstance(c, str) and c.lower().startswith("usu")
                ),
                None,
            )
        if not col_usuario:
            usuarios = []
        else:
            usuarios = sorted(
                self.df[col_usuario].dropna().astype(str).str.upper().unique()
            )

        if "Data criação" in self.df.columns:
            self.df["Data criação"] = pd.to_datetime(
                self.df["Data criação"], errors="coerce", dayfirst=True
            )
            meses = sorted(
                self.df["Data criação"].dropna().dt.strftime("%b/%Y").unique(),
                key=self.chave_mes_ano,
            )
        else:
            meses = []
        tipos = (
            sorted(self.df["Tipo"].dropna().unique())
            if "Tipo" in self.df.columns
            else []
        )
        agendamentos = (
            sorted(self.df["Agendamento"].dropna().unique())
            if "Agendamento" in self.df.columns
            else []
        )

        for lb, vals in [
            (self.filtro_usuario, usuarios),
            (self.filtro_mes, meses),
            (self.filtro_tipo, tipos),
            (self.filtro_agendamento, agendamentos),
        ]:
            lb.clear()
            for v in vals:
                lb.addItem(str(v))

    def gerar_grafico_com_filtros(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(
                self, "Aviso", "Nenhum dado carregado para gerar gráfico."
            )
            return

        df_filtrado = self.df.copy()

        # Filtro Usuário
        usuarios_sel = [
            item.text().upper() for item in self.filtro_usuario.selectedItems()
        ]
        if usuarios_sel:
            df_filtrado = df_filtrado[
                df_filtrado["Usuário"].astype(str).str.upper().isin(usuarios_sel)
            ]

        # Filtro Mês
        meses_sel = [item.text() for item in self.filtro_mes.selectedItems()]
        if meses_sel and "Data criação" in df_filtrado.columns:
            df_filtrado["Data criação"] = pd.to_datetime(
                df_filtrado["Data criação"], errors="coerce", dayfirst=True
            )
            df_filtrado["Mes"] = df_filtrado["Data criação"].dt.strftime("%b/%Y")
            df_filtrado = df_filtrado[df_filtrado["Mes"].isin(meses_sel)]

        # Filtro Tipo
        tipos_sel = [item.text() for item in self.filtro_tipo.selectedItems()]
        if tipos_sel:
            df_filtrado = df_filtrado[df_filtrado["Tipo"].isin(tipos_sel)]

        # Filtro Agendamento
        agendamentos_sel = [
            item.text() for item in self.filtro_agendamento.selectedItems()
        ]
        if agendamentos_sel:
            df_filtrado = df_filtrado[df_filtrado["Agendamento"].isin(agendamentos_sel)]

        if df_filtrado.empty:
            QMessageBox.information(
                self, "Informação", "Nenhum dado corresponde aos filtros selecionados."
            )
            return

        self.mostrar_grafico_usuarios(df_filtrado)

    def mostrar_grafico_usuarios(self, df_filtrado):
        tipo = self.combo_tipo_grafico.currentText()
        estilo = self.combo_design_grafico.currentText()
        usar_cores_diferentes = self.check_cores_diferentes.isChecked()

        if estilo != getattr(self, "current_style", None):
            plt.style.use(estilo)
            self.current_style = estilo

        fig = Figure(figsize=(8, 6), dpi=100)
        self.fig = fig
        ax = fig.add_subplot(111)
        if "Usuário" not in df_filtrado.columns or df_filtrado.empty:
            self.grafico_label.clear()
            return

        counts = df_filtrado["Usuário"].value_counts()
        labels = counts.index.astype(str)
        values = counts.values

        if usar_cores_diferentes:
            cores = plt.get_cmap("tab10").colors
            bar_colors = [cores[i % len(cores)] for i in range(len(labels))]
        else:
            bar_colors = ["#1565c0"] * len(labels)

        if estilo == "dark_background":
            label_color = "white"
            grid_color = "#444"
        else:
            label_color = "black"
            grid_color = "#ccc"

        if tipo == "Colunas Horizontais":
            bars = ax.barh(labels, values, color=bar_colors)
            ax.set_xlabel("Quantidade de Minutas", color=label_color)
            ax.set_ylabel("Usuário", color=label_color)
            ax.tick_params(axis="x", colors=label_color)
            ax.tick_params(axis="y", colors=label_color)
            ax.grid(True, axis="x", color=grid_color, linestyle="--", alpha=0.5)
            for bar in bars:
                width = bar.get_width()
                ax.annotate(
                    f"{float(width):.0f}",
                    xy=(width, bar.get_y() + bar.get_height() / 2),
                    xytext=(8, 0),
                    textcoords="offset points",
                    va="center",
                    ha="left",
                    fontweight="bold",
                    color=label_color,
                    fontsize=11,
                )
        else:
            bars = ax.bar(labels, values, color=bar_colors)
            ax.set_xlabel("Usuário", color=label_color)
            ax.set_ylabel("Quantidade de Minutas", color=label_color)
            ax.tick_params(axis="x", colors=label_color, labelsize=9)
            ax.tick_params(axis="y", colors=label_color)
            plt.setp(
                ax.get_xticklabels(),
                rotation=45,
                ha="right",
                rotation_mode="anchor",
            )
            ax.grid(True, axis="y", color=grid_color, linestyle="--", alpha=0.5)
            for bar in bars:
                height = bar.get_height()
                ax.annotate(
                    f"{float(height):.0f}",
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 14),
                    textcoords="offset points",
                    va="bottom",
                    ha="center",
                    fontweight="bold",
                    color=label_color,
                    fontsize=11,
                )
        ax.set_title(
            "Produtividade por Usuário",
            fontsize=18,
            fontweight="bold",
            color=label_color,
        )
        fig.tight_layout()

        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
            fig.savefig(tmpfile.name, bbox_inches="tight")
            self.grafico_label.setPixmap(QPixmap(tmpfile.name))

    def setup_tab_config(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Pesos para cada tipo de Agendamento:"))

        # Tabela de pesos
        self.table_pesos = QTableWidget()
        self.table_pesos.setColumnCount(2)
        self.table_pesos.setHorizontalHeaderLabels(["Tipo de Agendamento", "Peso"])
        self.table_pesos.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.table_pesos.setStyleSheet(
            """
            QTableWidget {
                background: #fff;
                font-size: 14px;
                border: 2px solid #a97b50;
            }
            QHeaderView::section {
                background: #a97b50;
                color: white;
                font-weight: bold;
                border: none;
                font-size: 15px;
            }
            QTableWidget::item {
                border-bottom: 1px solid #d1d5db;
            }
            QTableWidget::item:selected {
                background: rgba(189, 158, 87, 0.25);
                color: #1a232e;
            }
            """
        )

        layout.addWidget(self.table_pesos)

        # Botões de ação
        btns = QHBoxLayout()
        btn_salvar = QPushButton("Salvar Pesos")
        btn_salvar.setStyleSheet(
            """
            QPushButton {
                background-color: #43a047; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #388e3c; }
        """
        )
        btn_salvar.clicked.connect(self.salvar_pesos_interface)
        btns.addWidget(btn_salvar)

        btn_carregar = QPushButton("Carregar Pesos")
        btn_carregar.setStyleSheet(
            """
            QPushButton {
                background-color: #1565c0; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #0d47a1; }
        """
        )
        btn_carregar.clicked.connect(self.carregar_pesos_interface)
        btns.addWidget(btn_carregar)

        btn_padrao = QPushButton("Restaurar Padrão")
        btn_padrao.setStyleSheet(
            """
            QPushButton {
                background-color: #e53935; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #b71c1c; }
        """
        )
        btn_padrao.clicked.connect(self.restaurar_pesos_padrao)
        btns.addWidget(btn_padrao)

        layout.addLayout(btns)
        self.tab_config.setLayout(layout)
        self.atualizar_tabela_pesos()

    def atualizar_tabela_pesos(self):
        # Garante que todos os tipos de agendamento únicos estejam listados
        tipos = sorted(set(self.gerenciador_pesos.pesos.keys()))
        self.table_pesos.setRowCount(len(tipos))
        for row, tipo in enumerate(tipos):
            item_tipo = QTableWidgetItem(str(tipo))
            item_tipo.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            self.table_pesos.setItem(row, 0, item_tipo)
            item_peso = QTableWidgetItem(str(self.gerenciador_pesos.obter_peso(tipo)))
            self.table_pesos.setItem(row, 1, item_peso)

    def salvar_pesos_interface(self):
        for row in range(self.table_pesos.rowCount()):
            tipo = self.table_pesos.item(row, 0).text()
            peso_str = self.table_pesos.item(row, 1).text()
            try:
                peso = float(peso_str)
                self.gerenciador_pesos.atualizar_peso(tipo, peso)
            except ValueError:
                QMessageBox.warning(self, "Erro", f"Peso inválido para '{tipo}'.")
                return
        self.gerenciador_pesos.salvar_pesos()
        QMessageBox.information(self, "Sucesso", "Pesos salvos com sucesso!")
        self.atualizar_tabela_pesos()

    def carregar_pesos_interface(self):
        self.gerenciador_pesos.carregar_pesos()
        self.atualizar_tabela_pesos()
        QMessageBox.information(self, "Configuração", "Pesos carregados do arquivo.")

    def restaurar_pesos_padrao(self):
        resposta = QMessageBox.question(
            self,
            "Restaurar Padrão",
            "Deseja restaurar todos os pesos para 1.0?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if resposta == QMessageBox.StandardButton.Yes:
            self.gerenciador_pesos.pesos = defaultdict(lambda: 1.0)
            self.gerenciador_pesos.salvar_pesos()
            self.atualizar_tabela_pesos()
            QMessageBox.information(
                self, "Restaurado", "Todos os pesos foram restaurados para 1.0."
            )

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
        # Ordena meses no formato "mmm/AAAA"
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
            df = carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            df = df.loc[:, ~df.columns.duplicated()]
            if not df.empty:
                dfs.append(df)
        self.df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        self.df = self.df.loc[:, ~self.df.columns.duplicated()]
        for col in self.df.columns:
            if (
                isinstance(col, str)
                and col.lower().startswith("usu")
                and col != "Usuário"
            ):
                self.df.rename(columns={col: "Usuário"}, inplace=True)
        if "Usuário" in self.df.columns:
            self.df["Usuário"] = self.df["Usuário"].astype(str).str.strip().str.upper()
        self.atualizar_filtros_grafico()
        self.atualizar_kpis()

    def comparar_todos_meses(self):
        for i in range(self.listbox_meses.count()):
            self.listbox_meses.item(i).setSelected(True)
        self.comparar_meses()

    def limpar_dados_anteriores(self):
        resposta = QMessageBox.question(
            self,
            "Confirmação",
            "Você tem certeza que deseja excluir os dados dos meses?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if resposta != QMessageBox.StandardButton.Yes:
            return
        pasta_dados_mensais = "dados_mensais"
        if os.path.exists(pasta_dados_mensais):
            for arquivo in os.listdir(pasta_dados_mensais):
                file_path_arquivo = os.path.join(pasta_dados_mensais, arquivo)
                if os.path.isfile(file_path_arquivo):
                    os.remove(file_path_arquivo)
        self.df = None
        self.atualizar_listbox_meses()
        self.table_comp.setRowCount(0)

    def exportar_tabela_para_excel(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Salvar como Excel", "", "Excel Files (*.xlsx)"
        )
        if path:
            df_export = self.obter_df_tabela_comparacao()
            if df_export is not None:
                df_export.to_excel(path, index=False)
                QMessageBox.information(
                    self, "Exportação", "Tabela exportada com sucesso!"
                )

    def exportar_tabela_para_pdf(self):
        if FPDF is None:
            QMessageBox.warning(self, "Erro", "FPDF não está instalado.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Salvar como PDF", "", "PDF Files (*.pdf)"
        )
        if path:
            df_export = self.obter_df_tabela_comparacao()
            if df_export is not None:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=10)
                # Cabeçalho
                for col in df_export.columns:
                    pdf.cell(40, 10, col, 1)
                pdf.ln()
                # Dados
                for _, row in df_export.iterrows():
                    for val in row:
                        pdf.cell(40, 10, str(val), 1)
                    pdf.ln()
                pdf.output(path)
                QMessageBox.information(
                    self, "Exportação", "Tabela exportada em PDF com sucesso!"
                )

    def obter_df_tabela_comparacao(self):
        # Converte a tabela de comparação em DataFrame
        colunas = [
            self.table_comp.horizontalHeaderItem(i).text()
            for i in range(self.table_comp.columnCount())
        ]
        dados = []
        for row in range(self.table_comp.rowCount()):
            linha = []
            for col in range(self.table_comp.columnCount()):
                item = self.table_comp.item(row, col)
                linha.append(item.text() if item else "")
            dados.append(linha)
        if dados:
            return pd.DataFrame(dados, columns=colunas)
        return None

    def atualizar_tudo(self):
        self.carregar_dados_mensais()
        self.atualizar_listbox_meses()
        self.atualizar_filtros_grafico()
        self.atualizar_kpis()

    def exportar_grafico(self):
        if not hasattr(self, "fig") or self.fig is None:
            QMessageBox.warning(self, "Aviso", "Nenhum gráfico para exportar.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Salvar Gráfico", "", "PNG Files (*.png);;PDF Files (*.pdf)"
        )
        if path:
            try:
                self.fig.savefig(path)
                QMessageBox.information(
                    self, "Exportação", "Gráfico exportado com sucesso!"
                )
            except RuntimeError:
                QMessageBox.warning(
                    self,
                    "Erro",
                    "O gráfico não está mais disponível. Gere novamente antes de exportar.",
                )

    def closeEvent(self, event):
        self._debounce_timer.stop()
        self._canvas_valido = False
        try:
            self.combo_tipo_grafico.currentIndexChanged.disconnect()
            self.combo_design_grafico.currentIndexChanged.disconnect()
            self.check_cores_diferentes.stateChanged.disconnect()
            self.filtro_usuario.itemSelectionChanged.disconnect()
            self.filtro_mes.itemSelectionChanged.disconnect()
            self.filtro_tipo.itemSelectionChanged.disconnect()
            self.filtro_agendamento.itemSelectionChanged.disconnect()
            self._debounce_timer.stop()
        except Exception:
            pass
        event.accept()

    def _atualizar_grafico_debounced(self):
        if (
            not getattr(self, "_canvas_valido", True)
            or not hasattr(self, "canvas_grafico")
            or self.canvas_grafico is None
        ):
            return
        self.gerar_grafico_com_filtros()

    def tab_changed(self, idx):
        aba = self.tabs.widget(idx)
        if aba == self.tab_graficos:
            self._canvas_valido = True
        else:
            self._canvas_valido = False
            self._debounce_timer.stop()
