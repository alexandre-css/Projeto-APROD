from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QSplashScreen, QApplication
from PyQt5.QtGui import QPixmap

import sys
import os
import pandas as pd
import calendar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from collections import defaultdict
from xls2xlsx import XLS2XLSX
import re
import json
import unicodedata
import numpy as np
from PyQt5.QtGui import QPixmap
import time

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

from PyQt5.QtWidgets import (
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
)
from PyQt5.QtCore import Qt


def carregar_planilha(file_path):
    import pandas as pd

    header = pd.read_excel(file_path, nrows=0)
    if "peso" in header.columns:
        df = pd.read_excel(file_path, usecols="A:H")
    else:
        df = pd.read_excel(file_path, skiprows=1, usecols="A:G")
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
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Erro ao ler {self.arquivo_pesos}: {e}. Reiniciando arquivo.")
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


import time


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Análise e Produtividade")
        self.resize(1200, 800)
        self.df = None
        self.gerenciador_pesos = GerenciadorPesosAgendamento()

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.tab_kpi = QWidget()
        self.tab_graficos = QWidget()
        self.tab_comparacao_meses = QWidget()
        self.tab_config = QWidget()

        self.tabs.addTab(self.tab_kpi, "KPIs")
        self.tabs.addTab(self.tab_graficos, "Análise e Gráficos")
        self.tabs.addTab(self.tab_comparacao_meses, "Comparação de Meses")
        self.tabs.addTab(self.tab_config, "Atribuir Pesos")

        self.setup_tab_kpi()
        self.setup_tab_graficos()
        self.setup_tab_comparacao_meses()
        self.setup_tab_config()
        self.carregar_dados_mensais()
        self.atualizar_listbox_meses()
        self.atualizar_filtros_grafico()
        self.atualizar_kpis()

    def setup_tab_kpi(self):
        # Layout principal da aba KPIs
        vbox = QVBoxLayout()
        hbox_kpis = QHBoxLayout()
        self.kpi_labels = {}

        kpis = [
            ("Minutas", "minutas"),
            ("Média Produtividade", "media"),
            ("Dia Mais Produtivo", "dia"),
            ("Top 3 Usuários", "top3"),
        ]
        for title, key in kpis:
            card = QGroupBox(title)
            vbox_card = QVBoxLayout()
            lbl = QLabel("--")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("font-size: 26px; font-weight: bold; color: #1565c0;")
            vbox_card.addWidget(lbl)
            card.setLayout(vbox_card)
            card.setStyleSheet(
                """
                QGroupBox { 
                    background: #e3f2fd; 
                    border-radius: 10px; 
                    font-weight: bold; 
                    border: 2px solid #1565c0;
                }
                """
            )
            hbox_kpis.addWidget(card)
            self.kpi_labels[key] = lbl
        vbox.addLayout(hbox_kpis)

        # Botão Importar Arquivo Excel
        btn_importar = QPushButton("Importar Arquivo Excel")
        btn_importar.setStyleSheet(
            "QPushButton { background-color: #1565c0; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px; } QPushButton:hover { background-color: #0d47a1; }"
        )
        btn_importar.clicked.connect(self.importar_arquivo_excel)
        vbox.addWidget(btn_importar, alignment=Qt.AlignLeft)

        self.tab_kpi.setLayout(vbox)

    def importar_arquivo_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Importar arquivo Excel",
            "",
            "Todos Arquivos (*);;Excel Files (*.xlsx *.xls)",
        )
        if path:
            nome_destino = os.path.basename(path)
            pasta = "dados_mensais"
            os.makedirs(pasta, exist_ok=True)
            destino = os.path.join(pasta, nome_destino)
            if not os.path.exists(destino):
                import shutil

                shutil.copy2(path, destino)
                QMessageBox.information(
                    self, "Importação", "Arquivo importado com sucesso!"
                )
                self.carregar_dados_mensais()
                self.atualizar_listbox_meses()
                self.atualizar_filtros_grafico()
                self.atualizar_kpis()
            else:
                QMessageBox.warning(
                    self, "Aviso", "Arquivo já existe na pasta de dados."
                )

    def setup_tab_graficos(self):
        layout = QHBoxLayout()

        # Filtros laterais
        filtro_box = QGroupBox("Filtros para Gráfico")
        filtro_box.setStyleSheet(
            "QGroupBox { background: #e3f2fd; border-radius: 8px; border: 2px solid #1565c0; }"
        )
        filtro_layout = QVBoxLayout()
        filtro_layout.addWidget(QLabel("Usuário:"))
        self.filtro_usuario = QListWidget()
        self.filtro_usuario.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_usuario.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_usuario)

        filtro_layout.addWidget(QLabel("Mês:"))
        self.filtro_mes = QListWidget()
        self.filtro_mes.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_mes.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_mes)

        filtro_layout.addWidget(QLabel("Tipo:"))
        self.filtro_tipo = QListWidget()
        self.filtro_tipo.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_tipo.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_tipo)

        filtro_layout.addWidget(QLabel("Agendamento:"))
        self.filtro_agendamento = QListWidget()
        self.filtro_agendamento.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_agendamento.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_agendamento)

        btn_graph = QPushButton("Gerar Gráfico com Filtros")
        btn_graph.setStyleSheet(
            """
            QPushButton {
                background-color: #43a047; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #388e3c; }
            """
        )
        btn_graph.clicked.connect(self.gerar_grafico_com_filtros)
        filtro_layout.addWidget(btn_graph)
        filtro_box.setLayout(filtro_layout)
        layout.addWidget(filtro_box, 1)

        # Painel de gráfico (ordem correta: crie Figure/canvas antes do layout!)
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.canvas_grafico = FigureCanvas(self.fig)
        self.grafico_layout = QVBoxLayout()
        self.grafico_layout.addWidget(self.canvas_grafico)

        self.painel_grafico = QGroupBox("Gráfico")
        self.painel_grafico.setStyleSheet(
            "QGroupBox { background: #fafafa; border-radius: 8px; }"
        )
        self.painel_grafico.setMinimumWidth(600)
        self.painel_grafico.setMinimumHeight(400)
        self.painel_grafico.setLayout(self.grafico_layout)
        layout.addWidget(self.painel_grafico, 3)

        self.tab_graficos.setLayout(layout)

    def setup_tab_comparacao_meses(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Selecione os meses para comparar:"))

        self.listbox_meses = QListWidget()
        self.listbox_meses.setSelectionMode(QListWidget.MultiSelection)
        self.listbox_meses.setMaximumHeight(110)
        self.listbox_meses.setStyleSheet(
            """
            QListWidget { background: #e3f2fd; font-size: 12pt; border-radius: 6px; }
        """
        )
        layout.addWidget(self.listbox_meses)

        btns = QHBoxLayout()
        btn_comparar = QPushButton("Comparar")
        btn_comparar.setStyleSheet(
            "QPushButton { background-color: #43a047; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px; } QPushButton:hover { background-color: #388e3c; }"
        )
        btn_comparar.clicked.connect(self.comparar_meses)
        btns.addWidget(btn_comparar)
        layout.addLayout(btns)

        self.table_comp = QTableWidget()
        self.table_comp.setStyleSheet(
            "QTableWidget { background: #f5f5f5; font-size: 11pt; } QHeaderView::section { background: #1565c0; color: white; font-weight: bold; }"
        )
        self.table_comp.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_comp.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table_comp)

        self.tab_comparacao_meses.setLayout(layout)
        self.atualizar_listbox_meses()

        btn_comparar_todos = QPushButton("Comparar Todos")
        btn_comparar_todos.setStyleSheet(
            "QPushButton { background-color: #1565c0; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px; } QPushButton:hover { background-color: #0d47a1; }"
        )
        btn_comparar_todos.clicked.connect(self.comparar_todos_meses)
        btns.addWidget(btn_comparar_todos)

        btn_excluir = QPushButton("Excluir dados dos meses")
        btn_excluir.setStyleSheet(
            "QPushButton { background-color: #e53935; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px; } QPushButton:hover { background-color: #b71c1c; }"
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
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
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
                item.setTextAlignment(Qt.AlignCenter)
                self.table_comp.setItem(row_idx, col_idx, item)

    def preencher_listbox_meses(self, lista_meses):
        self.listbox_meses.clear()
        for mes in lista_meses:
            self.listbox_meses.addItem(str(mes))

    def atualizar_kpis(self):
        # Carregue e consolide todos os arquivos .xlsx da pasta dados_mensais
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
            dfs.append(df)
        if not dfs:
            for lbl in self.kpi_labels.values():
                lbl.setText("--")
            return
        df_total = pd.concat(dfs, ignore_index=True)

        # Minutas (total de linhas)
        minutas = len(df_total)
        self.kpi_labels["minutas"].setText(str(minutas))

        # Média Produtividade (média de peso por usuário)
        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            media = df_total.groupby("Usuário")["peso"].sum().mean()
            self.kpi_labels["media"].setText(f"{media:.2f}")
        else:
            self.kpi_labels["media"].setText("--")

        # Dia Mais Produtivo (dia com maior soma de peso)
        if "Data criação" in df_total.columns and "peso" in df_total.columns:
            df_total["data"] = pd.to_datetime(
                df_total["Data criação"], errors="coerce"
            ).dt.date
            dias = df_total.groupby("data")["peso"].sum()
            if not dias.empty:
                dia_mais = dias.idxmax()
                valor_mais = dias.max()
                self.kpi_labels["dia"].setText(
                    f"{dia_mais.strftime('%d/%m/%Y')} ({valor_mais:.2f})"
                )
            else:
                self.kpi_labels["dia"].setText("--")
        else:
            self.kpi_labels["dia"].setText("--")

        # Top 3 Usuários (maior soma de peso)
        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            top3 = (
                df_total.groupby("Usuário")["peso"]
                .sum()
                .sort_values(ascending=False)
                .head(3)
            )
            top3_txt = "\n".join(
                [f"{i+1}º {u}: {v:.2f}" for i, (u, v) in enumerate(top3.items())]
            )
            self.kpi_labels["top3"].setText(top3_txt)
        else:
            self.kpi_labels["top3"].setText("--")

    def atualizar_filtros_grafico(self):
        # Preenche os filtros de acordo com os dados reais do DataFrame consolidado
        if self.df is None or self.df.empty:
            for lb in [
                self.filtro_usuario,
                self.filtro_mes,
                self.filtro_tipo,
                self.filtro_agendamento,
            ]:
                lb.clear()
            return

        usuarios = sorted(self.df["Usuário"].dropna().unique())
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
        usuarios_sel = [item.text() for item in self.filtro_usuario.selectedItems()]
        if usuarios_sel:
            df_filtrado = df_filtrado[df_filtrado["Usuário"].isin(usuarios_sel)]

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
        if not hasattr(self, "fig") or self.fig is None:
            QMessageBox.warning(self, "Erro", "Canvas do gráfico não inicializado.")
            return
        self.fig.clear()
        ax = self.fig.add_subplot(111)
        if "Usuário" not in df_filtrado.columns or df_filtrado.empty:
            self.canvas_grafico.draw()
            return
        counts = df_filtrado["Usuário"].value_counts()
        bars = ax.bar(counts.index.astype(str), counts.values, color="#1565c0")
        ax.set_xlabel("Usuário")
        ax.set_ylabel("Quantidade de Minutas")
        ax.set_title("Produtividade por Usuário", fontsize=18, fontweight="bold")
        for bar in bars:
            height = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                height,
                f"{float(height):.0f}",
                ha="center",
                va="bottom",
                fontweight="bold",
            )
        self.fig.tight_layout()
        self.canvas_grafico.draw()

    def setup_tab_config(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Pesos para cada tipo de Agendamento:"))

        # Tabela de pesos
        self.table_pesos = QTableWidget()
        self.table_pesos.setColumnCount(2)
        self.table_pesos.setHorizontalHeaderLabels(["Tipo de Agendamento", "Peso"])
        self.table_pesos.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_pesos.setStyleSheet(
            """
            QTableWidget { background: #f5f5f5; font-size: 12pt; }
            QHeaderView::section { background: #1565c0; color: white; font-weight: bold; }
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
            item_tipo.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
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
        from PyQt5.QtWidgets import QMessageBox

        resposta = QMessageBox.question(
            self,
            "Restaurar Padrão",
            "Deseja restaurar todos os pesos para 1.0?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if resposta == QMessageBox.Yes:
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
        except Exception:
            pass
        return None

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
            dfs.append(df)
        self.df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
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
            QMessageBox.Yes | QMessageBox.No,
        )
        if resposta != QMessageBox.Yes:
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
        canvas = getattr(self, "canvas_grafico", None)
        if canvas is None or not canvas.isVisible():
            QMessageBox.warning(self, "Aviso", "Nenhum gráfico para exportar.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Salvar Gráfico", "", "PNG Files (*.png);;PDF Files (*.pdf)"
        )
        if path:
            try:
                canvas.figure.savefig(path)
                QMessageBox.information(
                    self, "Exportação", "Gráfico exportado com sucesso!"
                )
            except RuntimeError:
                QMessageBox.warning(
                    self,
                    "Erro",
                    "O gráfico não está mais disponível. Gere novamente antes de exportar.",
                )


if __name__ == "__main__":
    app = QApplication(sys.argv)

    splash_pix = QPixmap("splash_loading.png")  # Use o caminho da sua imagem
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()

    def start_main():
        win = MainWindow()
        win.show()
        splash.finish(win)

    # Garante que a splash fique visível por pelo menos 2 segundos (2000 ms)
    QTimer.singleShot(2000, start_main)

    sys.exit(app.exec_())
