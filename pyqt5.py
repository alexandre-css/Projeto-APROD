import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem, QListWidget, QFileDialog,
    QGroupBox, QHeaderView, QFrame, QAbstractItemView
)
from PyQt5.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Análise e Produtividade")
        self.resize(1200, 800)
        self.df = None

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Abas
        self.tab_kpi = QWidget()
        self.tab_graficos = QWidget()
        self.tab_comparacao_meses = QWidget()
        self.tab_config = QWidget()

        self.tabs.addTab(self.tab_kpi, "KPIs")
        self.tabs.addTab(self.tab_graficos, "Análise e Gráficos")
        self.tabs.addTab(self.tab_comparacao_meses, "Comparação de Meses")
        self.tabs.addTab(self.tab_config, "Configurações")

        self.setup_tab_kpi()
        self.setup_tab_graficos()
        self.setup_tab_comparacao_meses()
        self.setup_tab_config()

    def setup_tab_kpi(self):
        layout = QHBoxLayout()
        kpis = [
            ("Minutas", "--"),
            ("Média Produtividade", "--"),
            ("Dia Mais Produtivo", "--"),
            ("Top 3 Usuários", "--")
        ]
        for title, value in kpis:
            card = QGroupBox(title)
            vbox = QVBoxLayout()
            lbl = QLabel(value)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("font-size: 26px; font-weight: bold; color: #1565c0;")
            vbox.addWidget(lbl)
            card.setLayout(vbox)
            card.setStyleSheet("""
                QGroupBox { 
                    background: #e3f2fd; 
                    border-radius: 10px; 
                    font-weight: bold; 
                    border: 2px solid #1565c0;
                }
            """)
            layout.addWidget(card)
        self.tab_kpi.setLayout(layout)

    def setup_tab_graficos(self):
        layout = QHBoxLayout()
        # Filtros laterais
        filtro_box = QGroupBox("Filtros para Gráfico")
        filtro_box.setStyleSheet("QGroupBox { background: #e3f2fd; border-radius: 8px; border: 2px solid #1565c0; }")
        filtro_layout = QVBoxLayout()

        # Filtro Usuário
        filtro_layout.addWidget(QLabel("Usuário:"))
        self.filtro_usuario = QListWidget()
        self.filtro_usuario.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_usuario.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_usuario)

        # Filtro Mês
        filtro_layout.addWidget(QLabel("Mês:"))
        self.filtro_mes = QListWidget()
        self.filtro_mes.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_mes.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_mes)

        # Filtro Tipo
        filtro_layout.addWidget(QLabel("Tipo:"))
        self.filtro_tipo = QListWidget()
        self.filtro_tipo.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_tipo.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_tipo)

        # Filtro Agendamento
        filtro_layout.addWidget(QLabel("Agendamento:"))
        self.filtro_agendamento = QListWidget()
        self.filtro_agendamento.setSelectionMode(QListWidget.MultiSelection)
        self.filtro_agendamento.setStyleSheet("background: #e3f2fd;")
        filtro_layout.addWidget(self.filtro_agendamento)

        # Botão Gerar Gráfico com cor
        btn_graph = QPushButton("Gerar Gráfico com Filtros")
        btn_graph.setStyleSheet("""
            QPushButton {
                background-color: #43a047; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #388e3c; }
        """)
        filtro_layout.addWidget(btn_graph)
        filtro_box.setLayout(filtro_layout)
        layout.addWidget(filtro_box, 1)

        # Painel de gráfico (placeholder)
        painel_grafico = QGroupBox("Gráfico")
        painel_grafico.setStyleSheet("QGroupBox { background: #fafafa; border-radius: 8px; }")
        painel_grafico.setMinimumWidth(600)
        painel_grafico.setMinimumHeight(400)
        layout.addWidget(painel_grafico, 3)

        self.tab_graficos.setLayout(layout)

    def setup_tab_comparacao_meses(self):
        layout = QVBoxLayout()
        # Label
        layout.addWidget(QLabel("Selecione os meses para comparar:"))
        # Listbox de meses (compacta e azul clara)
        self.listbox_meses = QListWidget()
        self.listbox_meses.setSelectionMode(QListWidget.MultiSelection)
        self.listbox_meses.setMaximumHeight(110)
        self.listbox_meses.setStyleSheet("""
            QListWidget { background: #e3f2fd; font-size: 12pt; border-radius: 6px; }
        """)
        layout.addWidget(self.listbox_meses)
        # Botões lado a lado com cores
        btns = QHBoxLayout()
        btn_comparar = QPushButton("Comparar")
        btn_comparar.setStyleSheet("""
            QPushButton {
                background-color: #43a047; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #388e3c; }
        """)
        btns.addWidget(btn_comparar)
        btn_comparar_todos = QPushButton("Comparar Todos")
        btn_comparar_todos.setStyleSheet("""
            QPushButton {
                background-color: #1565c0; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #0d47a1; }
        """)
        btns.addWidget(btn_comparar_todos)
        btn_excluir = QPushButton("Excluir dados dos meses")
        btn_excluir.setStyleSheet("""
            QPushButton {
                background-color: #e53935; color: white; font-weight: bold; border-radius: 8px; padding: 6px 18px;
            }
            QPushButton:hover { background-color: #b71c1c; }
        """)
        btns.addWidget(btn_excluir)
        layout.addLayout(btns)
        # Tabela de comparação
        self.table_comp = QTableWidget(10, 5)
        self.table_comp.setHorizontalHeaderLabels(["Usuário", "Mês 1", "Mês 2", "Mês 3", "Total"])
        self.table_comp.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_comp.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_comp.setStyleSheet("""
            QTableWidget { background: #f5f5f5; font-size: 11pt; }
            QHeaderView::section { background: #1565c0; color: white; font-weight: bold; }
        """)
        layout.addWidget(self.table_comp)
        self.tab_comparacao_meses.setLayout(layout)

    def setup_tab_config(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Configurações gerais aqui"))
        self.tab_config.setLayout(layout)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())