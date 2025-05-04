import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import MessageDialog
from tkinter import filedialog, messagebox
from tksheet import Sheet
from checklistcombobox import ChecklistCombobox
import pandas as pd
import calendar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import defaultdict
from xls2xlsx import XLS2XLSX
import re
import json
import os
import tkinter as tk
import unicodedata
import numpy as np
try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

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
    "seaborn": "seaborn-v0_8"
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

class SplashScreen(tk.Toplevel):
    def __init__(self, parent, theme="flatly"):
        super().__init__(parent)
        self.transient(parent)
        self.title("Carregando...")
        self.geometry("420x180+{}+{}".format(
            self.winfo_screenwidth()//2 - 210,
            self.winfo_screenheight()//2 - 90
        ))
        self.overrideredirect(True)
        self.configure(bg="#2b2b2b")
        self.style = ttk.Style(theme=theme)
        frame = ttk.Frame(self, bootstyle=PRIMARY)
        frame.pack(expand=True, fill="both", padx=20, pady=20)
        label = ttk.Label(frame, text="Carregando Sistema de Produtividade...",
                          font=("Segoe UI", 16, "bold"),
                          anchor="center", bootstyle=INVERSE)
        label.pack(pady=(10, 20))
        self.progress = ttk.Progressbar(frame, mode="indeterminate", bootstyle="info-striped", length=340)
        self.progress.pack(pady=(0, 10))
        self.progress.start(12)

    def close(self):
        if self.progress:
            self.progress.stop()
        self.destroy()

class ExcelAnalyzerApp:
    def __init__(self, root):
        self.dados_carregados = False
        self.inicializando = True
        self.root = root
        self.root.title("Sistema de Análise e Produtividade")
        self.root.geometry("1200x800")
        self.df = None
        self.tipos_agendamento_unicos = []
        self.coluna_agendamento = "Data criação"
        self.style = ttk.Style()
        self.style.configure('TNotebook.Tab', font=('Segoe UI', 14, 'bold'), padding=[24, 8])
        self.style.configure('Pesos.Treeview', font=('Segoe UI', 14))
        self.style.configure('Pesos.Treeview.Heading', font=('Segoe UI', 15, 'bold'))
        self.style.configure('Relatorio.Treeview', font=('Segoe UI', 14))
        self.style.configure('Relatorio.Treeview.Heading', font=('Segoe UI', 15, 'bold'))
        self.style.configure('Semana.Treeview', font=('Segoe UI', 13))
        self.style.configure('Semana.Treeview.Heading', font=('Segoe UI', 14, 'bold'))

        self.frame_kpis = ttk.Frame(self.root)
        self.frame_kpis.pack(fill="x", padx=10, pady=8)
        self.kpi_labels = {}
        for kpi in ["Minutas", "Média Produtividade", "Dia Mais Produtivo", "Top 3 Usuários"]:
            frame = ttk.LabelFrame(self.frame_kpis, text=kpi, padding=8, bootstyle=PRIMARY)
            frame.pack(side="left", padx=8, fill="y")
            label = ttk.Label(
                frame,
                text="--",
                font=("Segoe UI", 16, "bold"),
                anchor="center",
                justify="center"
            )
            label.pack(expand=True, fill="both")
            self.kpi_labels[kpi] = label

        self.notebook = ttk.Notebook(root, bootstyle=PRIMARY)
        self.notebook.pack(fill="both", expand=True)
        self.criar_aba_analise_graficos()
        self.criar_aba_config_pesos()
        self.criar_aba_produtividade_semana()
        self.criar_aba_comparacao()
        self.criar_aba_comparacao_meses()
        self.criar_aba_configuracoes()
        self.carregar_configuracoes()
        self.gerenciador_pesos = GerenciadorPesosAgendamento()
        self.carregar_pesos_automaticamente()
        self.carregar_dados_mensais_automaticamente()
        self.inicializando = False
        self.atualizar_kpis()
        self.atualizar_checkboxes_usuarios()
        self.atualizar_tabela_pesos()
        self.gerar_tabela_semana()
        if hasattr(self, "atualizar_comboboxes_comparacao"):
            self.atualizar_comboboxes_comparacao()
        if hasattr(self, "comparar_meses"):
            self.comparar_meses(mostrar_popup=False)
        self.gerenciador_pesos.carregar_pesos()
        self.root.protocol("WM_DELETE_WINDOW", self.fechar_programa)
        self.primeira_execucao = True
        if hasattr(self, "comparar_todos_meses"):
            self.comparar_todos_meses()

    @staticmethod
    def chave_mes_ano(mes_ano):
        meses_pt = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
        try:
            if not mes_ano or "/" not in mes_ano:
                raise ValueError("Formato inválido")
            mes_abrev, ano = mes_ano.split('/')
            mes_num = meses_pt.index(mes_abrev.lower()) + 1
            return (int(ano), mes_num)
        except Exception as e:
            print(f"Erro ao converter '{mes_ano}': {str(e)}")
            return (0, 0)

    def atualizar_tipos_agendamento_unicos(self):
        """
        Atualiza a lista de tipos únicos de agendamento a partir do DataFrame consolidado (self.df),
        garantindo que todos os tipos presentes em todos os meses carregados apareçam na tabela de pesos.
        """
        if self.df is not None and "Agendamento" in self.df.columns:
            tipos_unicos = self.df["Agendamento"].dropna().unique()
            tipos_limpos = [re.sub(r'\s*\(.*?\)', '', str(x)).strip() for x in tipos_unicos]
            self.tipos_agendamento_unicos = sorted(set(tipos_limpos))
        else:
            self.tipos_agendamento_unicos = []

    def testar_extracao_meses(self):
        pasta = "dados_mensais"
        if not os.path.exists(pasta):
            print("Pasta dados_mensais não existe!")
            return
        for arq in os.listdir(pasta):
            file_path = os.path.join(pasta, arq)
            print(arq, self.extrair_mes_ano_do_arquivo(file_path))

    def limpar_dados_anteriores(self):
        from tkinter import messagebox

        resposta = messagebox.askyesno(
            "Confirmação",
            "Você tem certeza que deseja excluir os dados dos meses?"
        )
        if not resposta:
            return

        pasta_dados_mensais = "dados_mensais"
        if os.path.exists(pasta_dados_mensais):
            for arquivo in os.listdir(pasta_dados_mensais):
                file_path_arquivo = os.path.join(pasta_dados_mensais, arquivo)
                if os.path.isfile(file_path_arquivo):
                    os.remove(file_path_arquivo)

        # Limpa o DataFrame principal
        self.df = None

        # Atualiza KPIs e interfaces dependentes
        self.atualizar_kpis()
        self.atualizar_comboboxes_comparacao()
        self.atualizar_checkboxes_usuarios()
        self.atualizar_tabela_pesos()
        self.gerar_tabela_semana()
        if hasattr(self, "comparar_meses"):
            self.comparar_meses(mostrar_popup=False)

    def carregar_dados_mensais_automaticamente(self):
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta) if arq.endswith(('.xlsx'))] if os.path.exists(pasta) else []
        if not arquivos:
            return
        dfs = []
        for arquivo in arquivos:
            file_path = os.path.join(pasta, arquivo)
            try:
                df = carregar_planilha(file_path)
                if "Usuário" in df.columns:
                    df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()


                col_agendamento = next((col for col in df.columns if col == "Agendamento"), None)
                if "peso" not in df.columns:
                    if col_agendamento:
                        tipo_limpo = df[col_agendamento].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
                        df["peso"] = tipo_limpo.apply(self.gerenciador_pesos.obter_peso)
                    else:
                        df["peso"] = 1.0

                df = df.loc[:, ~df.columns.duplicated()]
                dfs.append(df)
            except Exception as e:
                print(f"Erro ao carregar {arquivo}: {e}")

        if dfs:
            self.df = pd.concat(dfs, ignore_index=True)
            self.atualizar_tipos_agendamento_unicos()
            self.atualizar_kpis()
            self.atualizar_checkboxes_usuarios()
            self.atualizar_tabela_pesos()
            self.atualizar_filtros_grafico()
            self.gerar_tabela_semana()
            if hasattr(self, "atualizar_comboboxes_comparacao"):
                self.atualizar_comboboxes_comparacao()
            if hasattr(self, "comparar_meses"):
                self.comparar_meses(mostrar_popup=False)

    def salvar_pesos_interface(self):
        try:
            for item in self.tree_pesos.get_children():
                tipo, peso_str = self.tree_pesos.item(item, 'values')
                novo_peso = float(peso_str)
                self.gerenciador_pesos.atualizar_peso(tipo, novo_peso)
            self.gerenciador_pesos.salvar_pesos()
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Sucesso", "Pesos salvos com sucesso!")
        except ValueError:
            if getattr(self, 'inicializando', False):
                return
            messagebox.showerror("Erro", "Formato de peso inválido!")

    def salvar_pesos_automaticamente(self):
        with open("pesos.json", "w", encoding="utf-8") as f:
            json.dump(dict(self.gerenciador_pesos.pesos), f, ensure_ascii=False, indent=2)
        self.carregar_pesos_automaticamente()

    def salvar_dados_mes(self):
        if self.df is None:
            return

        df_save = self.df.copy()

        # Garante coluna "peso" atualizada
        col_agendamento = "Agendamento"
        if col_agendamento in df_save.columns:
            tipo_limpo = df_save[col_agendamento].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
            df_save["peso"] = tipo_limpo.apply(self.gerenciador_pesos.obter_peso)
        else:
            df_save["peso"] = 1.0

        # Identifica mês/ano para salvar
        col_data = "Data criação"
        datas_validas = pd.to_datetime(df_save[col_data], errors="coerce").dropna() if col_data in df_save.columns else []
        if not datas_validas.empty:
            mes = datas_validas.dt.month.mode()[0]
            ano = datas_validas.dt.year.mode()[0]
        else:
            messagebox.showerror("Erro", "A coluna 'Data criação' está vazia ou inválida. Não é possível salvar o mês corretamente.")
            return

        pasta = "dados_mensais"
        if not os.path.exists(pasta):
            os.makedirs(pasta)
        nome_arquivo = os.path.basename(self.file_path.get())
        file_path = os.path.join(pasta, nome_arquivo)

        colunas_relevantes = [
            "Tipo", "Código", "Nro. processo", "Usuário", "Data criação", "Status", "Agendamento", "peso"
        ]
        for col in colunas_relevantes:
            if col not in df_save.columns:
                df_save[col] = ""
        df_save = df_save[colunas_relevantes]
        df_save.to_excel(file_path, index=False)

        if hasattr(self, 'atualizar_comboboxes_comparacao'):
            self.atualizar_comboboxes_comparacao()

        return file_path

    def carregar_pesos_automaticamente(self):
        """Carrega pesos do arquivo JSON e força atualização."""
        if os.path.exists("pesos.json"):
            with open("pesos.json", "r", encoding="utf-8") as f:
                pesos = json.load(f)
                self.gerenciador_pesos.pesos = defaultdict(lambda: 1.0, {k: float(v) for k, v in pesos.items()})

        # Atualiza tabelas SEM mostrar popups
        if hasattr(self, 'tree_comp_meses'):
            self.comparar_meses(mostrar_popup=False)

    def salvar_configuracoes(self):
        config = {
            "tema": self.tema_var.get(),
            "fonte": self.fonte_var.get(),
            "export_path": self.export_path_var.get(),
            "auto_save": self.auto_save_var.get()
        }
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def carregar_configuracoes(self):
        if os.path.exists("config.json"):
            with open("config.json", "r", encoding="utf-8") as f:
                config = json.load(f)
                self.tema_var.set(config.get("tema", "flatly"))
                self.fonte_var.set(config.get("fonte", "Pequeno"))
                self.export_path_var.set(config.get("export_path", ""))
                self.auto_save_var.set(config.get("auto_save", True))
            self.atualizar_checkboxes_usuarios()

    def fechar_programa(self):
        self.inicializando = True
        if getattr(self, 'dados_carregados', False):
            self.salvar_pesos_automaticamente()
            self.salvar_configuracoes()
        self.root.destroy()

    def criar_aba_analise_graficos(self):
        frame_analise = ttk.Frame(self.notebook)
        self.notebook.add(frame_analise, text="Análise e Gráficos")
        frame_arquivo = ttk.LabelFrame(frame_analise, text="Seleção de Arquivo", padding=10, bootstyle=INFO)
        frame_arquivo.pack(fill="x", padx=10, pady=5)
        self.file_path = ttk.StringVar()
        ttk.Entry(frame_arquivo, textvariable=self.file_path, width=60).pack(side="left", padx=5)
        ttk.Button(frame_arquivo, text="Procurar", command=self.browse_file, bootstyle=PRIMARY).pack(side="left", padx=5)
        self.label_meses_carregados = ttk.Label(frame_arquivo, text="Meses carregados: --", font=("Segoe UI", 11, "bold"))
        self.label_meses_carregados.pack(side="left", padx=15)

        painel_h = ttk.PanedWindow(frame_analise, orient="horizontal")
        painel_h.pack(fill="both", expand=True, padx=10, pady=10)

        frame_esquerdo = ttk.Frame(painel_h)
        painel_h.add(frame_esquerdo, weight=1)

        frame_filtros = ttk.LabelFrame(painel_h, text="Filtros para Gráfico", padding=10, bootstyle=INFO)
        painel_h.add(frame_filtros, weight=1)

        # Filtro por Usuário
        ttk.Label(frame_filtros, text="Usuário:").grid(row=0, column=0, sticky="w", pady=2)
        frame_usuario = ttk.LabelFrame(frame_filtros, text="Usuário", bootstyle=INFO, padding=2)
        frame_usuario.grid(row=0, column=1, sticky="ew", pady=2)
        self.listbox_usuario = tk.Listbox(
            frame_usuario, selectmode="multiple", exportselection=False, height=8,
            bg="#e3f2fd", highlightthickness=0, relief="flat", borderwidth=0
        )
        self.listbox_usuario.pack(side="left", fill="both", expand=True)
        scroll_usuario = ttk.Scrollbar(frame_usuario, orient="vertical", command=self.listbox_usuario.yview)
        scroll_usuario.pack(side="right", fill="y")
        self.listbox_usuario.config(yscrollcommand=scroll_usuario.set)

        # Filtro por Mês
        ttk.Label(frame_filtros, text="Mês:").grid(row=1, column=0, sticky="w", pady=2)
        frame_mes = ttk.LabelFrame(frame_filtros, text="Mês", bootstyle=INFO, padding=2)
        frame_mes.grid(row=1, column=1, sticky="ew", pady=2)
        self.listbox_mes = tk.Listbox(
            frame_mes, selectmode="multiple", exportselection=False, height=8,
            bg="#e3f2fd", highlightthickness=0, relief="flat", borderwidth=0
        )
        self.listbox_mes.pack(side="left", fill="both", expand=True)
        scroll_mes = ttk.Scrollbar(frame_mes, orient="vertical", command=self.listbox_mes.yview)
        scroll_mes.pack(side="right", fill="y")
        self.listbox_mes.config(yscrollcommand=scroll_mes.set)

        # Filtro por Tipo
        ttk.Label(frame_filtros, text="Tipo:").grid(row=2, column=0, sticky="w", pady=2)
        frame_tipo = ttk.LabelFrame(frame_filtros, text="Tipo", bootstyle=INFO, padding=2)
        frame_tipo.grid(row=2, column=1, sticky="ew", pady=2)
        self.listbox_tipo = tk.Listbox(
            frame_tipo, selectmode="multiple", exportselection=False, height=8,
            bg="#e3f2fd", highlightthickness=0, relief="flat", borderwidth=0
        )
        self.listbox_tipo.pack(side="left", fill="both", expand=True)
        scroll_tipo = ttk.Scrollbar(frame_tipo, orient="vertical", command=self.listbox_tipo.yview)
        scroll_tipo.pack(side="right", fill="y")
        self.listbox_tipo.config(yscrollcommand=scroll_tipo.set)

        # Filtro por Agendamento
        ttk.Label(frame_filtros, text="Agendamento:").grid(row=3, column=0, sticky="w", pady=2)
        frame_agendamento = ttk.LabelFrame(frame_filtros, text="Agendamento", bootstyle=INFO, padding=2)
        frame_agendamento.grid(row=3, column=1, sticky="ew", pady=2)
        self.listbox_agendamento = tk.Listbox(
            frame_agendamento, selectmode="multiple", exportselection=False, height=8,
            bg="#e3f2fd", highlightthickness=0, relief="flat", borderwidth=0
        )
        self.listbox_agendamento.pack(side="left", fill="both", expand=True)
        scroll_agendamento = ttk.Scrollbar(frame_agendamento, orient="vertical", command=self.listbox_agendamento.yview)
        scroll_agendamento.pack(side="right", fill="y")
        self.listbox_agendamento.config(yscrollcommand=scroll_agendamento.set)


        # Botão para aplicar filtros e gerar gráfico
        ttk.Button(frame_filtros, text="Gerar Gráfico com Filtros", command=self.generate_graph_com_filtros, bootstyle=SUCCESS).grid(row=4, column=0, columnspan=2, pady=10)

        # Configure grid para expansão
        frame_filtros.columnconfigure(1, weight=1)

        frame_controles = ttk.LabelFrame(frame_esquerdo, text="Configurações do Gráfico", padding=10, bootstyle=INFO)
        frame_controles.pack(fill="x", padx=5, pady=5)
        ttk.Label(frame_controles, text="Tipo de gráfico:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.graph_type = ttk.StringVar(value="barras horizontais")
        self.graph_type_map = {
            "barras verticais": "bar",
            "barras horizontais": "barh",
            "pizza": "pie",
            "em linha": "line"
        }
        graph_combo = ttk.Combobox(
            frame_controles, textvariable=self.graph_type,
            values=list(self.graph_type_map.keys()), state="readonly", width=15, bootstyle=INFO
        )
        graph_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(frame_controles, text="Estilo visual:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.style_var = ttk.StringVar(value="Solarize_Light2")
        self.style_map = STYLE_OPTIONS
        style_combo = ttk.Combobox(
            frame_controles, textvariable=self.style_var,
            values=list(self.style_map.keys()), state="readonly", width=20, bootstyle=INFO
        )
        style_combo.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(frame_controles, text="Título:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.title_var = ttk.StringVar(value="Produtividade por Usuário")
        ttk.Entry(frame_controles, textvariable=self.title_var, width=30).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.multi_color_var = ttk.BooleanVar(value=True)
        ttk.Checkbutton(frame_controles, text="Barras coloridas", variable=self.multi_color_var, bootstyle=SUCCESS).grid(row=4, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(frame_controles, text="Gerar Gráfico", command=self.generate_graph, bootstyle=SUCCESS).grid(row=5, column=0, columnspan=2, padx=5, pady=5)
        ttk.Button(frame_controles, text="Salvar Gráfico", command=self.save_graph, bootstyle=SECONDARY).grid(row=6, column=0, columnspan=2, padx=5, pady=5)
        self.frame_grafico = ttk.LabelFrame(painel_h, text="Gráfico", bootstyle=INFO)
        painel_h.add(self.frame_grafico, weight=2)
        self.fig = None
        self.canvas = None

    def atualizar_label_meses_carregados(self):
        pasta = "dados_mensais"
        meses = []
        if os.path.exists(pasta):
            for arq in os.listdir(pasta):
                file_path = os.path.join(pasta, arq)
                mes_ano = self.extrair_mes_ano_do_arquivo(file_path)
                if mes_ano != "Mês desconhecido":
                    meses.append(mes_ano)
        meses_validos = sorted(set(meses), key=self.chave_mes_ano)
        texto = "Meses carregados: " + (", ".join(meses_validos) if meses_validos else "--")
        self.label_meses_carregados.config(text=texto, font=("Segoe UI", 11, "bold"), foreground="#1565c0", padding=5)

    def criar_aba_config_pesos(self):
        frame_pesos = ttk.Frame(self.notebook)
        self.notebook.add(frame_pesos, text="Configuração de Pesos")
        ttk.Label(
            frame_pesos,
            text="Pesos para cada tipo de Agendamento:",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", padx=10, pady=(10,0))

        # Frame para tabela e scrollbar
        frame_tabela_pesos = ttk.Frame(frame_pesos)
        frame_tabela_pesos.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree_pesos = ttk.Treeview(
            frame_tabela_pesos,
            columns=("Tipo", "Peso"),
            show="headings",
            selectmode="extended",
            bootstyle=INFO
        )
        self.tree_pesos.heading("Tipo", text="Tipo de Agendamento", command=lambda: self.ordenar_treeview_pesos("Tipo", False))
        self.tree_pesos.heading("Peso", text="Peso Atribuído", command=lambda: self.ordenar_treeview_pesos("Peso", False))
        self.tree_pesos.column("Tipo", width=400, anchor="w", stretch=True)
        self.tree_pesos.column("Peso", width=100, anchor="center", stretch=True)

        # Barra de rolagem vertical
        scrollbar = ttk.Scrollbar(frame_tabela_pesos, orient="vertical", command=self.tree_pesos.yview)
        self.tree_pesos.configure(yscrollcommand=scrollbar.set)

        # Layout com grid para alinhamento correto
        self.tree_pesos.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        frame_tabela_pesos.rowconfigure(0, weight=1)
        frame_tabela_pesos.columnconfigure(0, weight=1)

        # Controles de edição de peso
        frame_controles = ttk.Frame(frame_pesos)
        frame_controles.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame_controles, text="Novo Peso:").pack(side="left")
        self.entry_novo_peso = ttk.Entry(frame_controles, width=10)
        self.entry_novo_peso.pack(side="left", padx=5)
        ttk.Button(
            frame_controles,
            text="Atualizar Peso Selecionado",
            command=self.atualizar_peso_selecionado,
            bootstyle=SUCCESS
        ).pack(side="left")

        frame_config = ttk.Frame(frame_pesos)
        frame_config.pack(fill="x", padx=10, pady=5)
        ttk.Button(
            frame_config,
            text="Salvar Configuração",
            command=self.salvar_pesos_interface,
            bootstyle=SECONDARY
        ).pack(side="left", padx=5)
        ttk.Button(
            frame_config,
            text="Carregar Configuração",
            command=self.carregar_configuracao_pesos,
            bootstyle=SECONDARY
        ).pack(side="left", padx=5)
        ttk.Button(
            frame_config,
            text="Restaurar Padrão",
            command=self.restaurar_pesos_padrao,
            bootstyle=WARNING
        ).pack(side="left", padx=5)

    def atualizar_filtros_grafico(self):
        if self.df is None or self.df.empty:
            for lb in [self.listbox_usuario, self.listbox_mes, self.listbox_tipo, self.listbox_agendamento]:
                lb.delete(0, tk.END)
            return

        usuarios = sorted(self.df['Usuário'].dropna().unique())
        if 'Data criação' in self.df.columns:
            self.df['Data criação'] = pd.to_datetime(self.df['Data criação'], errors="coerce", dayfirst=True)
            meses = sorted(self.df['Data criação'].dropna().dt.strftime('%b/%Y').unique(), key=self.chave_mes_ano)
        else:
            meses = []
        tipos = sorted(self.df['Tipo'].dropna().unique()) if 'Tipo' in self.df.columns else []
        agendamentos = sorted(self.df['Agendamento'].dropna().unique()) if 'Agendamento' in self.df.columns else []

        for lb, vals in [
            (self.listbox_usuario, usuarios),
            (self.listbox_mes, meses),
            (self.listbox_tipo, tipos),
            (self.listbox_agendamento, agendamentos)
        ]:
            lb.delete(0, tk.END)
            for v in vals:
                lb.insert(tk.END, v)

    def criar_aba_configuracoes(self):
        # Frame principal da aba
        frame_config = ttk.Frame(self.notebook)
        self.notebook.add(frame_config, text="Configurações")

        # Canvas + Scrollbar para rolagem vertical
        canvas = tk.Canvas(frame_config, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_config, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Variáveis de configuração
        self.tema_var = ttk.StringVar(value="flatly")
        self.fonte_var = ttk.StringVar(value="Pequeno")
        self.export_path_var = ttk.StringVar(value="")
        self.auto_save_var = ttk.BooleanVar(value=True)

        # Tema
        ttk.Label(scrollable_frame, text="Tema do programa:", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(16, 4))
        temas = ["morph", "flatly", "darkly", "cyborg", "journal", "solar", "vapor", "superhero"]
        combo_tema = ttk.Combobox(
            scrollable_frame, textvariable=self.tema_var,
            values=temas, state="readonly", width=20, bootstyle=INFO
        )
        combo_tema.pack(anchor="w", padx=16, pady=(0, 16))

        def mudar_tema(event=None):
            self.style.theme_use(self.tema_var.get())
        combo_tema.bind("<<ComboboxSelected>>", mudar_tema)

        # Pasta de exportação
        ttk.Label(scrollable_frame, text="Pasta padrão para exportação:").pack(anchor="w", padx=16, pady=(8, 2))
        ttk.Entry(scrollable_frame, textvariable=self.export_path_var, width=40).pack(anchor="w", padx=16)
        ttk.Button(
            scrollable_frame, text="Procurar",
            command=self.selecionar_pasta_exportacao, bootstyle=SECONDARY
        ).pack(anchor="w", padx=16, pady=(2, 8))

        # Fonte
        ttk.Label(scrollable_frame, text="Tamanho da fonte:").pack(anchor="w", padx=16, pady=(8, 2))
        combo_fonte = ttk.Combobox(
            scrollable_frame, textvariable=self.fonte_var,
            values=["Pequeno", "Médio", "Grande"], state="readonly"
        )
        combo_fonte.pack(anchor="w", padx=16)
        font_sizes = {"Pequeno": 10, "Médio": 14, "Grande": 18}

        def mudar_fonte(event=None):
            tamanho = font_sizes.get(self.fonte_var.get(), 10)
            self.style.configure('TNotebook.Tab', font=('Segoe UI', tamanho, 'bold'), padding=[24, 8])
            self.style.configure('Pesos.Treeview', font=('Segoe UI', tamanho))
            self.style.configure('Pesos.Treeview.Heading', font=('Segoe UI', tamanho+1, 'bold'))
            self.style.configure('Relatorio.Treeview', font=('Segoe UI', tamanho))
            self.style.configure('Relatorio.Treeview.Heading', font=('Segoe UI', tamanho+1, 'bold'))
            self.style.configure('Semana.Treeview', font=('Segoe UI', tamanho-1))
            self.style.configure('Semana.Treeview.Heading', font=('Segoe UI', tamanho, 'bold'))
        combo_fonte.bind("<<ComboboxSelected>>", mudar_fonte)
        mudar_fonte()

        # Auto save
        ttk.Checkbutton(
            scrollable_frame,
            text="Salvar configurações automaticamente ao sair",
            variable=self.auto_save_var
        ).pack(anchor="w", padx=16, pady=(0, 2))

        # Restaurar padrão
        ttk.Button(
            scrollable_frame,
            text="Restaurar configurações padrão",
            command=self.restaurar_configuracoes_padrao
        ).pack(anchor="w", padx=16, pady=(16, 8))


    def restaurar_configuracoes_padrao(self):
        if getattr(self, 'inicializando', False):
            return
        resposta = messagebox.askyesno(
            title="Confirmação",
            message="Você tem certeza que deseja restaurar as configurações para o padrão?"
        )
        if not resposta:
            return
        self.tema_var.set("flatly")
        self.fonte_var.set("Pequeno")
        self.export_path_var.set("")
        self.auto_save_var.set(True)
        self.style.theme_use("flatly")
        self.salvar_configuracoes()
        messagebox.showinfo("Configurações", "Configurações restauradas para o padrão!")

    def selecionar_pasta_exportacao(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.export_path_var.set(pasta)

    def criar_aba_produtividade_semana(self):
        frame_semana = ttk.Frame(self.notebook)
        self.notebook.add(frame_semana, text="Produtividade por Dia da Semana")
        frame_top = ttk.Frame(frame_semana)
        frame_top.pack(fill="x", padx=10, pady=10)
        ttk.Button(frame_top, text="Atualizar Tabela", command=self.gerar_tabela_semana, bootstyle=SUCCESS).pack(side="left")
        self.tree_semana = ttk.Treeview(frame_semana, show="headings", bootstyle=INFO)
        self.tree_semana.pack(fill="both", expand=True, padx=10, pady=10)
        self.labelframe_ranking = ttk.LabelFrame(frame_semana, text="Produtividade geral por dia da semana", padding=16, bootstyle=INFO)
        self.labelframe_ranking.pack(pady=18, padx=40, fill="x")
        self.label_ranking = ttk.Label(self.labelframe_ranking, text="", font=("Segoe UI", 14, "bold"), anchor="center", justify="center")
        self.label_ranking.pack(fill="x")

    def criar_aba_comparacao(self):
        frame_comp = ttk.Frame(self.notebook)
        self.notebook.add(frame_comp, text="Comparação")
        self.frame_check_users = ttk.LabelFrame(frame_comp, text="Selecione usuários para comparar", padding=10, bootstyle=ttk.INFO)
        self.frame_check_users.pack(fill="x", padx=10, pady=10)
        ttk.Button(self.frame_check_users, text="Atualizar Lista de Usuários", command=self.atualizar_checkboxes_usuarios, bootstyle=ttk.SECONDARY).pack(anchor="w", pady=(0,5))
        frame_btn = ttk.Frame(frame_comp)
        frame_btn.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame_btn, text="Comparar Selecionados", command=self.comparar_produtividade_usuarios, bootstyle=SUCCESS).pack(side="left", padx=5)
        ttk.Button(frame_btn, text="Comparar Todos", command=self.comparar_todos_meses, bootstyle=SUCCESS).pack(side="left", padx=5)
        frame_tabela = ttk.LabelFrame(frame_comp, text="Comparação de Usuários", padding=10, bootstyle=ttk.INFO)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)
        self.sheet_comp = Sheet(
            frame_tabela,
            data=[[""]],
            headers=["Usuário"],
            theme="light blue",
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            height=400
        )
        self.sheet_comp.pack(fill="both", expand=True)
        self.check_vars_usuarios = dict()
        self.checkbuttons_usuarios = dict()
        self.atualizar_checkboxes_usuarios()

    def comparar_produtividade_usuarios(self):
        # Coleta usuários selecionados
        usuarios = [u for u, var in self.check_vars_usuarios.items() if var.get()]
        if not usuarios:
            messagebox.showwarning("Aviso", "Selecione pelo menos um usuário para comparar!")
            return
        if self.df is None or self.df.empty:
            messagebox.showwarning("Aviso", "Nenhum dado carregado!")
            return
        col_usuario = "Usuário"
        col_peso = "peso"
        col_mes = "Data criação"
        df = self.df.copy()
        if col_mes in df.columns:
            df[col_mes] = pd.to_datetime(df[col_mes], errors="coerce")
            df["Mes"] = df[col_mes].dt.strftime("%b/%Y")
        else:
            df["Mes"] = ""
        # CORRIGIDO:
        meses_validos = [
            m for m in pd.Series(df["Mes"].unique()).dropna()
            if isinstance(m, str) and "/" in m
        ]
        meses = sorted(meses_validos, key=ExcelAnalyzerApp.chave_mes_ano)
        dados = []
        for usuario in usuarios:
            linha = [usuario]
            for mes in meses:
                prod = df[(df[col_usuario] == usuario) & (df["Mes"] == mes)][col_peso].sum() if col_peso in df.columns else 0.0
                linha.append(f"{float(prod):.2f}")
            linha.append(f"{float(df[df[col_usuario] == usuario][col_peso].sum()):.2f}")
            dados.append(linha)
        headers = ["Usuário"] + meses + ["Total"]
        self.sheet_comp.headers(headers)
        self.sheet_comp.set_sheet_data(dados)
        self.sheet_comp.set_all_column_widths(120)
        self.sheet_comp.column_width(0, 200)

    def atualizar_comboboxes_comparacao(self):
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta)] if os.path.exists(pasta) else []
        arquivos = [arq for arq in arquivos if arq.endswith(".xlsx")]
        opcoes_amigaveis = []
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            mes_ano = self.extrair_mes_ano_do_arquivo(file_path)
            opcoes_amigaveis.append(mes_ano)
        opcoes_amigaveis = [m for m in opcoes_amigaveis if m and "/" in str(m)]
        self.mapa_arquivo_meses = dict(zip(opcoes_amigaveis, arquivos))
        if hasattr(self, 'listbox_meses'):
            self.listbox_meses.delete(0, 'end')
            for opcao in opcoes_amigaveis:
                self.listbox_meses.insert('end', opcao)

    def criar_aba_comparacao_meses(self):
        frame_principal = ttk.Frame(self.notebook)
        self.notebook.add(frame_principal, text="Comparação de Meses")
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta)] if os.path.exists(pasta) else []
        arquivos = [arq for arq in arquivos if arq.endswith(".xlsx")]
        opcoes_amigaveis = []
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            mes_ano = self.extrair_mes_ano_do_arquivo(file_path)
            opcoes_amigaveis.append(mes_ano)
        opcoes_amigaveis = [m for m in opcoes_amigaveis if m and "/" in str(m)]
        self.mapa_arquivo_meses = dict(zip(opcoes_amigaveis, arquivos))
        frame_superior = ttk.Frame(frame_principal)
        frame_superior.pack(fill="x", padx=10, pady=10)

        # Label superior
        ttk.Label(frame_superior, text="Selecione os meses para comparar:").pack(anchor="w", padx=2, pady=(6, 2))

        # Frame estilizado para a listbox (bordas e integração visual)
        frame_listbox_meses = ttk.LabelFrame(
            frame_superior, text="Meses", bootstyle=INFO, padding=2
        )
        frame_listbox_meses.pack(fill="x", padx=0, pady=(0, 8))

        # Listbox compacta, azul claro, igual à dos filtros de gráfico
        self.listbox_meses = tk.Listbox(
            frame_listbox_meses,
            selectmode="multiple",
            height=5,           # altura compacta
            width=18,           # largura compacta
            font=("Segoe UI", 11),
            bg="#e3f2fd",
            highlightthickness=0,
            relief="flat",
            borderwidth=0
        )
        self.listbox_meses.pack(side="left", fill="x", expand=True)

        # Scrollbar integrada
        scroll = ttk.Scrollbar(frame_listbox_meses, orient="vertical", command=self.listbox_meses.yview)
        scroll.pack(side="right", fill="y")
        self.listbox_meses.config(yscrollcommand=scroll.set)

        # Preencher a listbox (exemplo)
        self.listbox_meses.delete(0, tk.END)
        for mes in opcoes_amigaveis:
            self.listbox_meses.insert("end", mes)

        frame_botoes = ttk.Frame(frame_superior)
        frame_botoes.pack(fill="x", pady=10)
        ttk.Button(frame_botoes, text="Comparar", command=self.comparar_meses, bootstyle=SUCCESS).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="Comparar Todos", command=self.comparar_todos_meses, bootstyle=SUCCESS).pack(side="left", padx=5)
        ttk.Button(frame_botoes, text="Excluir dados dos meses", command=self.limpar_dados_anteriores, bootstyle=DANGER).pack(side="left", padx=5)
        self.kpi_comp_frame = ttk.Frame(frame_principal)
        self.kpi_comp_frame.pack(fill="x", padx=10, pady=5)
        self.kpi_comp_labels = {}
        frame_tabela = ttk.LabelFrame(frame_principal, text="Comparação Detalhada", padding=0)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)
        frame_tabela.pack_propagate(False)
        frame_tabela.columnconfigure(0, weight=1)
        frame_tabela.rowconfigure(0, weight=1)
        self.sheet_comp_meses = Sheet(
            frame_tabela,
            data=[[""]],
            headers=["Usuário"],
            theme="light blue",
            show_x_scrollbar=True,
            show_y_scrollbar=True
        )
        self.sheet_comp_meses.grid(row=0, column=0, sticky="nsew")
        self.sheet_comp_meses.enable_bindings((
            "single_select", "row_select", "column_width_resize", "double_click_column_resize",
            "arrowkeys", "right_click_popup_menu", "rc_select", "copy", "cut", "paste", "delete", "undo", "edit_cell",
            "column_select", "column_select_drag", "column_select_toggle", "drag_select"
        ))
        self.sheet_comp_meses.extra_bindings([
            ("header_left_click", self.ordenar_coluna_comparacao_meses)
        ])

    def ordenar_coluna_comparacao_meses(self, event):
        col = event[1]
        if not hasattr(self, '_ordem_crescente'):
            self._ordem_crescente = {}
        crescente = self._ordem_crescente.get(col, True)
        self.sheet_comp_meses.sort_table(col, reverse=not crescente)
        self._ordem_crescente[col] = not crescente


    def verificar_arquivos_mensais(self):
        pasta = "dados_mensais"
        if os.path.exists(pasta):
            print("Arquivos encontrados em dados_mensais:")
            for arq in os.listdir(pasta):
                file_path = os.path.join(pasta, arq)
                print(f"- {arq} -> {self.extrair_mes_ano_do_arquivo(file_path)}")
        else:
            print("Pasta dados_mensais não existe!")

    def excluir_dados_comparacao(self):
        # Limpa tabelas/visualizações de comparação de meses se existirem
        if hasattr(self, 'sheet_comp_meses'):
            self.sheet_comp_meses.set_sheet_data([[""]])
            self.sheet_comp_meses.headers(["Usuário"])
        if hasattr(self, 'kpi_comp_labels'):
            for lbl in self.kpi_comp_labels.values():
                lbl.destroy()
            self.kpi_comp_labels = {}

    def sort_treeview_comp(self, col, reverse):
        items = [(self.tree_comp_meses.set(k, col), k) for k in self.tree_comp_meses.get_children('')]
        try:
            items = [(float(val.replace(',', '.')), k) for val, k in items]
        except ValueError:
            items = [(val.lower(), k) for val, k in items]
        items.sort(reverse=reverse)
        for index, (_, k) in enumerate(items):
            self.tree_comp_meses.move(k, '', index)
        self.tree_comp_meses.heading(col, command=lambda: self.sort_treeview_comp(col, not reverse))

    def comparar_meses(self, mostrar_popup=True):
        if not hasattr(self, 'kpi_comp_frame'):
            return
        if not hasattr(self, 'kpi_comp_labels'):
            self.kpi_comp_labels = {}

        selecoes = self.listbox_meses.curselection()
        if not hasattr(self, 'mapa_arquivo_meses') or not self.mapa_arquivo_meses:
            return

        meses_selecionados = [self.listbox_meses.get(i) for i in selecoes]
        if len(meses_selecionados) < 2:
            if mostrar_popup and not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "Selecione pelo menos dois meses para comparar!")
            return

        self.gerenciador_pesos.carregar_pesos()
        meses_selecionados.sort(key=ExcelAnalyzerApp.chave_mes_ano)
        arquivos_selecionados = [self.mapa_arquivo_meses[mes] for mes in meses_selecionados]

        dados_meses = []
        for mes, arquivo in zip(meses_selecionados, arquivos_selecionados):
            file_path = os.path.join("dados_mensais", arquivo)
            df = carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()

            col_usuario = "Usuário"
            if col_usuario in df.columns:
                # Padroniza nomes de usuários
                df[col_usuario] = df[col_usuario].astype(str).str.strip().str.upper()

            col_agendamento = "Agendamento"
            if "peso" not in df.columns:
                if col_agendamento in df.columns:
                    tipo_limpo = df[col_agendamento].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
                    df["peso"] = tipo_limpo.apply(self.gerenciador_pesos.obter_peso)
                else:
                    df["peso"] = 1.0

            if col_usuario in df.columns and "peso" in df.columns:
                prod = df.groupby(col_usuario)["peso"].sum()
            else:
                prod = pd.Series(dtype=float)
            dados_meses.append(prod)


        todos_usuarios = sorted({usuario for prod in dados_meses for usuario in prod.index})
        colunas = ["Usuário"] + meses_selecionados + ["Total"]
        dados_tabela = []
        celulas_verdes = []
        celulas_vermelhas = []
        celulas_top3 = []
        celulas_total = []
        celulas_zeradas = []

        # Monta a tabela de dados e destaca top 3 com medalhas
        top3_por_mes = {}
        for idx, usuario in enumerate(todos_usuarios):
            linha = [usuario]
            valores = []
            total = 0.0
            for i, prod in enumerate(dados_meses):
                try:
                    valor = float(prod.at[usuario])
                except Exception:
                    valor = 0.0
                valores.append(valor)
                total += valor

            for i, v in enumerate(valores):
                medalha = ""
                # Calcula top 3 de cada mês (coluna)
                if i not in top3_por_mes:
                    valores_col = [(row, float(dados_meses[i].get(u, 0.0))) for row, u in enumerate(todos_usuarios)]
                    top3 = sorted(valores_col, key=lambda x: x[1], reverse=True)[:3]
                    top3_por_mes[i] = [row for row, _ in top3]
                if idx in top3_por_mes[i]:
                    pos = top3_por_mes[i].index(idx)
                    if pos == 0:
                        medalha = "🥇 "
                    elif pos == 1:
                        medalha = "🥈 "
                    elif pos == 2:
                        medalha = "🥉 "
                if i == 0:
                    linha.append(f"{medalha}{v:.2f}")
                    if v == 0:
                        celulas_zeradas.append((idx, i+1))
                else:
                    if v > valores[i-1]:
                        linha.append(f"{medalha}{v:.2f} ▲")
                        celulas_verdes.append((idx, i+1))
                    elif v < valores[i-1]:
                        linha.append(f"{medalha}{v:.2f} ▼")
                        celulas_vermelhas.append((idx, i+1))
                    else:
                        linha.append(f"{medalha}{v:.2f}")
                    if v == 0:
                        celulas_zeradas.append((idx, i+1))
            linha.append(f"{total:.2f}")
            dados_tabela.append(linha)
            celulas_total.append((idx, len(colunas)-1))

        self.sheet_comp_meses.headers(colunas)
        self.sheet_comp_meses.set_sheet_data(dados_tabela)
        self.sheet_comp_meses.align_columns(columns=tuple(range(len(colunas))), align="center")
        self.sheet_comp_meses.align_header(columns=tuple(range(len(colunas))), align="center")

        # Coluna de total (azul claro)
        for row, col in celulas_total:
            self.sheet_comp_meses.highlight_cells(row=row, column=col, bg="#e3f2fd", fg="#1565c0")
        # Flecha para cima (verde forte)
        for row, col in celulas_verdes:
            self.sheet_comp_meses.highlight_cells(row=row, column=col, bg="#d0f5e8", fg="green")
        # Flecha para baixo (vermelho forte)
        for row, col in celulas_vermelhas:
            self.sheet_comp_meses.highlight_cells(row=row, column=col, bg="#ffebee", fg="red")

        self.sheet_comp_meses.set_all_column_widths(120)
        self.sheet_comp_meses.column_width(0, 200)

        # Limpa labels antigos
        for lbl in self.kpi_comp_labels.values():
            lbl.destroy()
        self.kpi_comp_labels = {}

        # Pega os dados EXATOS que estão sendo exibidos na tabela
        sheet_data = self.sheet_comp_meses.get_sheet_data()

        for col_idx, mes in enumerate(meses_selecionados, 1):
            total_mes = 0.0
            for row in sheet_data:
                if col_idx < len(row):  # Garante que a coluna existe
                    valor_str = str(row[col_idx])
                    # Remove medalhas, setas, espaços extras
                    for simbolo in ["▲", "▼", "🥇", "🥈", "🥉"]:
                        valor_str = valor_str.replace(simbolo, "")
                    valor_str = valor_str.replace(",", ".").strip()
                    try:
                        valor = float(valor_str)
                    except Exception:
                        valor = 0.0
                    total_mes += valor
            lbl_texto = f"Total {mes}: {total_mes:.2f}"
            lbl = ttk.Label(
                self.kpi_comp_frame,
                text=lbl_texto,
                font=("Segoe UI", 13, "bold"),
                foreground="#1565c0",
                background="#e3f2fd",
                padding=8
            )
            lbl.pack(side="left", padx=12, pady=6)
            self.kpi_comp_labels[mes] = lbl




    def comparar_todos_meses(self):
        # Seleciona todos os meses na listbox e executa a comparação
        self.listbox_meses.selection_set(0, "end")
        self.comparar_meses()

    def atualizar_checkboxes_usuarios(self):
        # Limpa checkboxes antigos
        for widget in self.frame_check_users.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                widget.destroy()
        self.check_vars_usuarios.clear()
        self.checkbuttons_usuarios.clear()

        if self.df is None or self.df.empty:
            if not getattr(self, 'inicializando', False):
                # Verifica se há arquivos na pasta antes de alertar
                import os
                if os.path.exists("dados_mensais") and os.listdir("dados_mensais"):
                    messagebox.showwarning("Aviso", "O arquivo não possui dados válidos!")
            return
        col_usuario = "Usuário"
        if col_usuario not in self.df.columns:
            return
        usuarios = sorted(self.df[col_usuario].dropna().unique())
        for usuario in usuarios:
            var = ttk.BooleanVar(value=False)
            chk = ttk.Checkbutton(self.frame_check_users, text=usuario, variable=var, bootstyle=INFO)
            chk.pack(anchor="w")
            self.check_vars_usuarios[usuario] = var
            self.checkbuttons_usuarios[usuario] = chk

    def gerar_tabela_semana(self):
        df_consolidado = self.consolidar_dados_meses()
        if df_consolidado is None or df_consolidado.empty:
            if not getattr(self, 'inicializando', False):
                # Só mostra o aviso se realmente há arquivos na pasta de dados
                pasta = "dados_mensais"
                arquivos = [arq for arq in os.listdir(pasta)] if os.path.exists(pasta) else []
                arquivos = [arq for arq in arquivos if arq.endswith(".xlsx")]
                if arquivos:
                    messagebox.showwarning("Aviso", "Não há dados mensais carregados!")
            return

        col_usuario = "Usuário"
        col_data = "Data criação"
        col_agendamento = "Agendamento"
        if not all(col in df_consolidado.columns for col in [col_usuario, col_data, col_agendamento]):
            messagebox.showwarning("Aviso", "A planilha precisa ter as colunas 'Usuário', 'Data criação' e 'Agendamento'.")
            return

        df = df_consolidado.copy()
        df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")
        df = df.dropna(subset=[col_usuario, col_data])

        dias_semana_pt = [
            'Segunda-feira', 'Terça-feira', 'Quarta-feira',
            'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo'
        ]
        dias_trad = {
            'Monday': 'Segunda-feira',
            'Tuesday': 'Terça-feira',
            'Wednesday': 'Quarta-feira',
            'Thursday': 'Quinta-feira',
            'Friday': 'Sexta-feira',
            'Saturday': 'Sábado',
            'Sunday': 'Domingo'
        }
        df["DiaSemana"] = df[col_data].dt.day_name().map(dias_trad)

        if "DiaSemana" not in df.columns or df["DiaSemana"].isnull().all():
            messagebox.showerror("Erro", "Não foi possível criar a coluna 'DiaSemana'. Verifique se a coluna de data está correta.")
            return

        tabela = df.pivot_table(
            index=col_usuario,
            columns="DiaSemana",
            values="peso",
            aggfunc="sum",
            fill_value=0
        )
        tabela = tabela.reindex(columns=dias_semana_pt, fill_value=0)

        self.tree_semana.delete(*self.tree_semana.get_children())
        self.tree_semana["columns"] = ["Usuário"] + dias_semana_pt

        for col in self.tree_semana["columns"]:
            self.tree_semana.heading(col, text=col, command=lambda c=col: self.ordenar_treeview_semana(c, False))
            self.tree_semana.column(col, width=120, anchor="center")

        for usuario, row in tabela.iterrows():
            valores = [usuario] + [float(row.get(d, 0)) for d in dias_semana_pt]
            self.tree_semana.insert("", "end", values=valores)

        soma_por_dia = tabela.sum(axis=0)
        ranking = soma_por_dia.sort_values(ascending=False)
        ranking_str = "\n".join(
            [f"{i+1}º - {dia}: {float(ranking[dia]):.2f}" for i, dia in enumerate(ranking.index)]
        )

        self.label_ranking.config(
            text=ranking_str,
            font=("Segoe UI", 14, "bold"),
            anchor="center",
            justify="center",
            foreground="#3a3a3a"
        )

    def ordenar_treeview_semana(self, col, reverse):
        l = [(self.tree_semana.set(k, col), k) for k in self.tree_semana.get_children('')]
        try:
            l.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
        except ValueError:
            l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree_semana.move(k, '', index)
        self.tree_semana.heading(col, command=lambda: self.ordenar_treeview_semana(col, not reverse))

    def ordenar_treeview_pesos(self, col, reverse):
        l = [(self.tree_pesos.set(k, col), k) for k in self.tree_pesos.get_children('')]
        try:
            l.sort(key=lambda t: float(str(t[0]).replace(',', '.')), reverse=reverse)
        except ValueError:
            l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree_pesos.move(k, '', index)
        self.tree_pesos.heading(col, command=lambda: self.ordenar_treeview_pesos(col, not reverse))

    def ordenar_treeview_relatorio(self, col, reverse):
        l = [(self.tree_relatorio.set(k, col), k) for k in self.tree_relatorio.get_children('')]
        try:
            l.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
        except ValueError:
            l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree_relatorio.move(k, '', index)
        self.tree_relatorio.heading(col, command=lambda: self.ordenar_treeview_relatorio(col, not reverse))

    def browse_file(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_paths:
            self.file_path.set(";".join(file_paths))
            self.carregar_dados_multiplos(list(file_paths))


    def carregar_dados(self):
        try:
            if getattr(self, 'inicializando', False):
                return
            if not self.file_path.get():
                if not getattr(self, 'inicializando', False):
                    messagebox.showwarning("Aviso", "Selecione um arquivo primeiro!")
                return
            file_path = self.file_path.get()
            df = carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()

            col_agendamento = "Agendamento"
            if "peso" not in df.columns:
                if col_agendamento in df.columns:
                    tipo_limpo = df[col_agendamento].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
                    df["peso"] = tipo_limpo.apply(self.gerenciador_pesos.obter_peso)
                else:
                    df["peso"] = 1.0

            self.df = df
            if "Agendamento" in self.df.columns:
                tipos_unicos = self.df["Agendamento"].dropna().unique()
                tipos_limpos = [re.sub(r'\s*\(.*?\)', '', str(x)).strip() for x in tipos_unicos]
                self.tipos_agendamento_unicos = sorted(set(tipos_limpos))
            else:
                self.tipos_agendamento_unicos = []
            self.atualizar_kpis()
            self.atualizar_checkboxes_usuarios()
            self.atualizar_tabela_pesos()
            self.gerar_tabela_semana()
            if hasattr(self, "atualizar_comboboxes_comparacao"):
                self.atualizar_comboboxes_comparacao()
            if hasattr(self, "comparar_meses"):
                self.comparar_meses(mostrar_popup=False)
            self.salvar_dados_mes()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{str(e)}")
            return
        messagebox.showinfo("Sucesso", "Dados carregados com sucesso!")
        self.testar_extracao_meses()
        self.atualizar_label_meses_carregados()

    def extrair_mes_ano_do_arquivo(self, file_path):
        try:
            df = carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            col_data = "Data criação"
            if col_data in df.columns and not df.empty:
                # Converte para datetime (padrão brasileiro: dd/mm/aaaa hh:mm:ss)
                datas = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce").dropna()
                if not datas.empty:
                    mes = int(datas.dt.month.mode()[0])
                    ano = int(datas.dt.year.mode()[0])
                    meses_pt = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                                'jul', 'ago', 'set', 'out', 'nov', 'dez']
                    return f"{meses_pt[mes-1]}/{ano}"
        except Exception as e:
            print(f"Erro ao extrair mês/ano do arquivo {file_path}: {e}")
        return "Mês desconhecido"

    def atualizar_tabela_pesos(self):
        print("Atualizando tabela de pesos...")
        ordem_atual = [self.tree_pesos.item(item)['values'][0] for item in self.tree_pesos.get_children()]
        if not ordem_atual:
            ordem_atual = sorted(self.tipos_agendamento_unicos)
        self.tree_pesos.delete(*self.tree_pesos.get_children())
        for tipo in ordem_atual:
            peso = self.gerenciador_pesos.obter_peso(tipo)
            self.tree_pesos.insert("", "end", values=(tipo, peso))

    def atualizar_peso_selecionado(self):
        selecionados = self.tree_pesos.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione um ou mais tipos de Agendamento primeiro!")
            return
        tipos_selecionados = [self.tree_pesos.item(item)['values'][0] for item in selecionados]
        novo_peso = self.entry_novo_peso.get()
        try:
            peso_float = float(novo_peso)
        except ValueError:
            messagebox.showerror("Erro", "Peso inválido! Use um número válido.")
            return
        for tipo in tipos_selecionados:
            self.gerenciador_pesos.atualizar_peso(tipo, peso_float)
        self.gerenciador_pesos.salvar_pesos()
        self.gerenciador_pesos.carregar_pesos()
        # Recalcule pesos no DataFrame
        if self.df is not None and "Agendamento" in self.df.columns:
            tipos_unicos = self.df["Agendamento"].dropna().unique()
            tipos_limpos = [re.sub(r'\s*\(.*?\)', '', str(x)).strip() for x in tipos_unicos]
            self.tipos_agendamento_unicos = sorted(set(tipos_limpos))
        else:
            self.tipos_agendamento_unicos = []

        self.atualizar_tabela_pesos()
        self.atualizar_kpis()
        self.gerar_tabela_semana()
        self.comparar_meses(mostrar_popup=False)
        novos_ids = []
        for item in self.tree_pesos.get_children():
            tipo_atual = self.tree_pesos.item(item)['values'][0]
            if tipo_atual in tipos_selecionados:
                novos_ids.append(item)
                self.tree_pesos.item(item, tags=("destaque",))
                self.tree_pesos.tag_configure("destaque", background="#fff59d")
        if novos_ids:
            self.root.after(2000, lambda: [self.tree_pesos.item(item, tags=()) for item in novos_ids])

    def generate_graph(self):
        if self.inicializando:
            return
        if self.df is None or self.df.empty:
            if not getattr(self, 'inicializando', False):
                # Só mostra o aviso se realmente há arquivos na pasta de dados
                if os.path.exists("dados_mensais") and os.listdir("dados_mensais"):
                    messagebox.showwarning("Aviso", "O arquivo não possui dados válidos!")
            return
        coluna = "Usuário"
        for widget in self.frame_grafico.winfo_children():
            widget.destroy()
        plt.style.use(self.style_map[self.style_var.get()])
        self.fig = plt.Figure(figsize=(8, 6), dpi=100)
        ax = self.fig.add_subplot(111)
        counts = self.df[coluna].value_counts()
        colors = plt.cm.tab10.colors if self.multi_color_var.get() else '#1f77b4'
        graph_type = self.graph_type_map[self.graph_type.get()]
        title = self.title_var.get()
        ax.set_title(title, fontsize=22, fontweight='bold', fontname='DejaVu Sans', color='#1a237e', pad=25)
        if graph_type == "bar":
            bars = ax.bar(counts.index.astype(str), counts.values, color=colors)
            ax.set_xlabel(coluna)
            ax.set_ylabel('Quantidade')
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2, height, f'{float(height):.2f}', ha='center', va='bottom', fontweight='bold')
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
        elif graph_type == "barh":
            bars = ax.barh(counts.index.astype(str), counts.values, color=colors)
            ax.set_ylabel(coluna)
            ax.set_xlabel('Quantidade')
            for bar in bars:
                width = bar.get_width()
                ax.text(width, bar.get_y() + bar.get_height()/2, f'{width}', ha='left', va='center', fontweight='bold')
        elif graph_type == "pie":
            wedges, texts, autotexts = ax.pie(
                counts.values, labels=counts.index.astype(str),
                autopct='%1.1f%%', startangle=90, colors=colors
            )
            ax.axis('equal')
            for autotext in autotexts:
                autotext.set_fontweight('bold')
        elif graph_type == "line":
            ax.plot(counts.index.astype(str), counts.values, marker='o', color=colors)
            ax.set_xlabel(coluna)
            ax.set_ylabel('Quantidade')
            for i, v in enumerate(counts.values):
                ax.text(i, v, f"{v}", ha='center', va='bottom', fontweight='bold')
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
        self.fig.tight_layout()
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.frame_grafico)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def generate_graph_com_filtros(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Aviso", "Nenhum dado carregado para gerar gráfico.")
            return

        df_filtrado = self.df.copy()

        # Múltiplos Usuários
        usuarios_sel = [self.listbox_usuario.get(i) for i in self.listbox_usuario.curselection()]
        if usuarios_sel:
            df_filtrado = df_filtrado[df_filtrado['Usuário'].isin(usuarios_sel)]

        # Múltiplos Meses
        meses_sel = [self.listbox_mes.get(i) for i in self.listbox_mes.curselection()]
        if meses_sel and 'Data criação' in df_filtrado.columns:
            df_filtrado['Data criação'] = pd.to_datetime(df_filtrado['Data criação'], errors="coerce", dayfirst=True)
            df_filtrado['Mes'] = df_filtrado['Data criação'].dt.strftime('%b/%Y')
            df_filtrado = df_filtrado[df_filtrado['Mes'].isin(meses_sel)]

        # Múltiplos Tipos
        tipos_sel = [self.listbox_tipo.get(i) for i in self.listbox_tipo.curselection()]
        if tipos_sel:
            df_filtrado = df_filtrado[df_filtrado['Tipo'].isin(tipos_sel)]

        # Múltiplos Agendamentos
        agendamentos_sel = [self.listbox_agendamento.get(i) for i in self.listbox_agendamento.curselection()]
        if agendamentos_sel:
            df_filtrado = df_filtrado[df_filtrado['Agendamento'].isin(agendamentos_sel)]

        if df_filtrado.empty:
            messagebox.showinfo("Informação", "Nenhum dado corresponde aos filtros selecionados.")
            return

        self.gerar_grafico_filtrado(df_filtrado)

    def carregar_dados_multiplos(self, file_paths):
        dfs = []
        for file_path in file_paths:
            try:
                df = carregar_planilha(file_path)
                if "Usuário" in df.columns:
                    df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
                col_agendamento = "Agendamento"
                if "peso" not in df.columns:
                    if col_agendamento in df.columns:
                        tipo_limpo = df[col_agendamento].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
                        df["peso"] = tipo_limpo.apply(self.gerenciador_pesos.obter_peso)
                    else:
                        df["peso"] = 1.0
                # Salva o arquivo padronizado na pasta dados_mensais
                pasta = "dados_mensais"
                if not os.path.exists(pasta):
                    os.makedirs(pasta)
                nome_arquivo = os.path.basename(file_path)
                df.to_excel(os.path.join(pasta, nome_arquivo), index=False)
                dfs.append(df)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar {file_path}:\n{str(e)}")

        if dfs:
            self.df = pd.concat(dfs, ignore_index=True)
            self.atualizar_tipos_agendamento_unicos()
            self.atualizar_kpis()
            self.atualizar_checkboxes_usuarios()
            self.atualizar_tabela_pesos()
            self.gerar_tabela_semana()
            if hasattr(self, "atualizar_comboboxes_comparacao"):
                self.atualizar_comboboxes_comparacao()
            self.atualizar_filtros_grafico()
            # Seleciona todos os meses na listbox antes de comparar
            if hasattr(self, "listbox_meses"):
                self.listbox_meses.selection_clear(0, "end")
                self.listbox_meses.selection_set(0, "end")
            if hasattr(self, "comparar_meses"):
                self.comparar_meses(mostrar_popup=False)
            self.atualizar_label_meses_carregados()
            messagebox.showinfo("Sucesso", "Todos os arquivos foram carregados com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado válido foi carregado.")

    def save_graph(self):
        if self.fig is None:
            if getattr(self, 'inicializando', False):
                return
            messagebox.showwarning("Aviso", "Gere um gráfico primeiro!")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile="grafico.png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("PDF", "*.pdf")]
        )
        if file_path:
            self.fig.savefig(file_path, bbox_inches='tight', dpi=300)
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Sucesso", f"Gráfico salvo em:\n{file_path}")

    def gerar_relatorio_produtividade(self):
        if self.df is None or self.df.empty:
            return
        col_usuario = "Usuário"
        col_agendamento = "Agendamento"
        col_tipo = "Tipo"
        if col_tipo in self.df.columns and col_agendamento in self.df.columns:
            self.df.loc[self.df[col_tipo].str.upper().str.strip() == "ACÓRDÃO", col_agendamento] = "Juntada de Relatório/Voto/Acórdão (581)"
        df_temp = self.df.copy()
        df_temp["Data criação"] = pd.to_datetime(df_temp["Data criação"], errors="coerce")
        df_temp["Mes"] = df_temp["Data criação"].apply(lambda x: x.strftime("%b/%Y") if not pd.isnull(x) else "")
        usuarios = sorted(df_temp[col_usuario].unique())
        meses_validos = [m for m in df_temp["Mes"].unique() if m and "/" in str(m)]
        meses = sorted(meses_validos, key=ExcelAnalyzerApp.chave_mes_ano)
        dados_usuarios = {Usuário: {} for Usuário in usuarios}
        for Usuário in usuarios:
            grupo = df_temp[df_temp[col_usuario] == Usuário]
            for mes in meses:
                prod_mes = grupo.loc[grupo["Mes"] == mes, "peso"].sum() if "peso" in grupo.columns else 0.0
                dados_usuarios[Usuário][mes] = float(prod_mes)
        colunas = ["Usuário"] + meses + ["Total"]
        dados_tabela = []
        for Usuário in usuarios:
            linha = [Usuário]
            total = 0.0
            for mes in meses:
                valor = dados_usuarios[Usuário].get(mes, 0.0)
                linha.append(formatar_valor(valor))
                total += valor
            linha.append(f"{float(total):.2f}")
            dados_tabela.append(linha)
        self.tree_relatorio.delete(*self.tree_relatorio.get_children())
        self.tree_relatorio["columns"] = colunas
        for i, col in enumerate(colunas):
            self.tree_relatorio.heading(col, text="Usuário" if col == "Usuário" else col)
            self.tree_relatorio.column(col, width=120 if i > 0 else 180, anchor="center")
        for linha in dados_tabela:
            self.tree_relatorio.insert("", "end", values=linha)

    def mostrar_mensagem(tipo, titulo, mensagem):
        if tipo == "erro":
            messagebox.showerror(titulo, mensagem)
        elif tipo == "aviso":
            messagebox.showwarning(titulo, mensagem)
        elif tipo == "info":
            messagebox.showinfo(titulo, mensagem)

    def consolidar_dados_meses(self):
        if self.df is not None and not self.df.empty:
            return self.df.copy()
        return None

    def visualizar_graficos_produtividade(self):
        if self.df is None or self.df.empty or not self.tree_relatorio.get_children():
            if not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "Gere o relatório de produtividade primeiro!")
            return

        colunas = self.tree_relatorio["columns"]
        meses = [col for col in colunas if col not in ("Usuário", "Total")]
        if not meses:
            messagebox.showwarning("Aviso", "Gere o relatório de produtividade primeiro!")
            return

        def abrir_popup():
            popup = ttk.Toplevel(self.root)
            popup.title("Selecione o mês")
            popup.geometry("350x300")
            ttk.Label(popup, text="Escolha o mês para o gráfico:").pack(pady=10)
            lb = tk.Listbox(popup, selectmode="single", font=("Segoe UI", 12))
            lb.pack(fill="both", expand=True, padx=15, pady=10)
            lb.insert("end", "Todos os meses")
            for mes in meses:
                lb.insert("end", mes)
            def confirmar():
                idx = lb.curselection()
                if not idx:
                    messagebox.showwarning("Aviso", "Selecione um mês!")
                    return
                popup.destroy()
                self._mostrar_grafico_produtividade(meses, lb.get(idx[0]))
            ttk.Button(popup, text="OK", command=confirmar, bootstyle=SUCCESS).pack(pady=10)
        abrir_popup()

    def _mostrar_grafico_produtividade(self, meses, mes_escolhido):
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

        usuarios = []
        pontuacoes = []

        if mes_escolhido == "Todos os meses":
            idxs = [self.tree_relatorio["columns"].index(m) for m in meses]
            for item in self.tree_relatorio.get_children():
                valores = self.tree_relatorio.item(item, "values")
                usuarios.append(valores[0])
                total = sum(float(valores[i]) for i in idxs)
                pontuacoes.append(total)
            titulo = "Produtividade Total por Usuário"
        else:
            idx = self.tree_relatorio["columns"].index(mes_escolhido)
            for item in self.tree_relatorio.get_children():
                valores = self.tree_relatorio.item(item, "values")
                usuarios.append(valores[0])
                pontuacoes.append(float(valores[idx]))
            titulo = f"Produtividade por Usuário - {mes_escolhido}"

        grafico_window = ttk.Toplevel(self.root)
        grafico_window.title(titulo)
        grafico_window.geometry("1050x700")

        frame_controles = ttk.Frame(grafico_window, padding=10)
        frame_controles.pack(fill="x")
        ttk.Label(frame_controles, text="Tipo de gráfico:").pack(side="left", padx=5)
        tipo_grafico = ttk.StringVar(value="barras horizontais")
        grafico_combo = ttk.Combobox(
            frame_controles,
            textvariable=tipo_grafico,
            values=["barras verticais", "barras horizontais", "pizza", "em linha"],
            state="readonly",
            width=15,
            bootstyle=ttk.INFO
        )
        grafico_combo.pack(side="left", padx=5)

        ttk.Label(frame_controles, text="Estilo visual:").pack(side="left", padx=5)
        style_var_prod = ttk.StringVar(value="Solarize_Light2")
        style_combo_prod = ttk.Combobox(
            frame_controles,
            textvariable=style_var_prod,
            values=list(STYLE_OPTIONS.keys()),
            state="readonly",
            width=20,
            bootstyle=ttk.INFO
        )
        style_combo_prod.pack(side="left", padx=5)
        multi_color_var_prod = ttk.BooleanVar(value=True)
        ttk.Checkbutton(frame_controles, text="Barras coloridas", variable=multi_color_var_prod, bootstyle=ttk.SUCCESS).pack(side="left", padx=5)
        salvar_btn = ttk.Button(frame_controles, text="Salvar Gráfico", bootstyle=ttk.SECONDARY)
        salvar_btn.pack(side="left", padx=20)
        gerar_btn = ttk.Button(frame_controles, text="Gerar Gráfico", bootstyle=ttk.SUCCESS)
        gerar_btn.pack(side="left", padx=5)

        frame_grafico = ttk.Frame(grafico_window)
        frame_grafico.pack(fill="both", expand=True, padx=10, pady=10)
        self.fig_produtividade = None

        def gerar_grafico_produtividade():
            for widget in frame_grafico.winfo_children():
                widget.destroy()
            plt.style.use(STYLE_OPTIONS[style_var_prod.get()])
            fig = plt.Figure(figsize=(8, 6))
            ax = fig.add_subplot(111)
            colors = plt.cm.tab10.colors if multi_color_var_prod.get() else '#1f77b4'
            tipo = tipo_grafico.get()
            if tipo == "barras verticais":
                bars = ax.bar(usuarios, pontuacoes, color=colors)
                ax.set_xlabel('Usuários')
                ax.set_ylabel('Pontuação de Produtividade')
                plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
                for bar in bars:
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height, f'{height:.2f}', ha='center', va='bottom', fontweight='bold')
            elif tipo == "barras horizontais":
                bars = ax.barh(usuarios, pontuacoes, color=colors)
                ax.set_ylabel('Usuários')
                ax.set_xlabel('Pontuação de Produtividade')
                for bar in bars:
                    width = bar.get_width()
                    ax.text(width, bar.get_y() + bar.get_height()/2., f'{width:.2f}', ha='left', va='center', fontweight='bold')
            elif tipo == "pizza":
                wedges, texts, autotexts = ax.pie(
                    pontuacoes,
                    labels=usuarios,
                    autopct='%1.1f%%',
                    startangle=90,
                    colors=colors if multi_color_var_prod.get() else None
                )
                ax.axis('equal')
                for autotext in autotexts:
                    autotext.set_fontweight('bold')
            elif tipo == "em linha":
                ax.plot(usuarios, pontuacoes, marker='o', color=colors if not multi_color_var_prod.get() else plt.cm.tab10(0))
                ax.set_xlabel('Usuários')
                ax.set_ylabel('Pontuação de Produtividade')
                plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
                for i, v in enumerate(pontuacoes):
                    ax.text(i, v + 0.1, f"{v:.2f}", ha='center', fontweight='bold')
            ax.set_title(titulo, fontsize=22, fontweight='bold', fontname='DejaVu Sans', color='#1a237e', pad=25)
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            self.fig_produtividade = fig

        def salvar_grafico_produtividade():
            if self.fig_produtividade is None:
                messagebox.showwarning("Aviso", "Gere o gráfico primeiro!")
                return
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                initialfile="grafico_produtividade.png",
                filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("PDF", "*.pdf")]
            )
            if file_path:
                self.fig_produtividade.savefig(file_path, bbox_inches='tight', dpi=300)
                messagebox.showinfo("Sucesso", f"Gráfico salvo em:\n{file_path}")

        salvar_btn.config(command=salvar_grafico_produtividade)
        gerar_btn.config(command=gerar_grafico_produtividade)
        gerar_grafico_produtividade()

    def gerar_grafico_filtrado(self, df_filtrado):
        # Exemplo: gráfico de barras por usuário
        for widget in self.frame_grafico.winfo_children():
            widget.destroy()
        plt.style.use(self.style_map[self.style_var.get()])
        fig = plt.Figure(figsize=(8, 6), dpi=100)
        ax = fig.add_subplot(111)
        coluna = "Usuário"
        counts = df_filtrado[coluna].value_counts()
        colors = plt.cm.tab10.colors if self.multi_color_var.get() else '#1f77b4'
        graph_type = self.graph_type_map[self.graph_type.get()]
        title = self.title_var.get()
        ax.set_title(title, fontsize=22, fontweight='bold', fontname='DejaVu Sans', color='#1a237e', pad=25)
        if graph_type == "bar":
            bars = ax.bar(counts.index.astype(str), counts.values, color=colors)
            ax.set_xlabel(coluna)
            ax.set_ylabel('Quantidade')
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2, height, f'{float(height):.2f}', ha='center', va='bottom', fontweight='bold')
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
        elif graph_type == "barh":
            bars = ax.barh(counts.index.astype(str), counts.values, color=colors)
            ax.set_ylabel(coluna)
            ax.set_xlabel('Quantidade')
            for bar in bars:
                width = bar.get_width()
                ax.text(width, bar.get_y() + bar.get_height()/2, f'{width}', ha='left', va='center', fontweight='bold')
        elif graph_type == "pie":
            wedges, texts, autotexts = ax.pie(
                counts.values, labels=counts.index.astype(str),
                autopct='%1.1f%%', startangle=90, colors=colors
            )
            ax.axis('equal')
            for autotext in autotexts:
                autotext.set_fontweight('bold')
        elif graph_type == "line":
            ax.plot(counts.index.astype(str), counts.values, marker='o', color=colors)
            ax.set_xlabel(coluna)
            ax.set_ylabel('Quantidade')
            for i, v in enumerate(counts.values):
                ax.text(i, v, f"{v}", ha='center', va='bottom', fontweight='bold')
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self.fig = fig
        self.canvas = canvas


    def exportar_relatorio_excel(self):
        if self.df is None or self.df.empty or not self.tree_relatorio.get_children():
            if not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "Gere o relatório primeiro!")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])
        if file_path:
            dados = [self.tree_relatorio.item(item)['values'] for item in self.tree_relatorio.get_children()]
            df_export = pd.DataFrame(dados, columns=self.tree_relatorio["columns"])
            df_export.to_excel(file_path, index=False)
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Sucesso", f"Relatório exportado para:\n{file_path}")

    def exportar_relatorio_pdf(self):
        if FPDF is None:
            if not getattr(self, 'inicializando', False):
                messagebox.showerror("Erro", "Instale a biblioteca fpdf: pip install fpdf")
            return
        if self.df is None or self.df.empty or not self.tree_relatorio.get_children():
            if not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "Gere o relatório primeiro!")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if file_path:
            dados = [self.tree_relatorio.item(item)['values'] for item in self.tree_relatorio.get_children()]
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(0, 10, "Relatório de Produtividade", ln=True, align="C")
            pdf.ln(10)
            for col in self.tree_relatorio["columns"]:
                pdf.cell(60, 10, str(col), border=1)
            pdf.ln()
            for linha in dados:
                for valor in linha:
                    pdf.cell(60, 10, str(valor), border=1)
                pdf.ln()
            pdf.output(file_path)
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Sucesso", f"Relatório exportado para:\n{file_path}")

    def salvar_configuracao_pesos(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Configuração de Pesos", "*.json")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(dict(self.gerenciador_pesos.pesos), f, ensure_ascii=False, indent=2)
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Sucesso", "Configuração salva!")

    def carregar_configuracao_pesos(self):
        file_path = filedialog.askopenfilename(filetypes=[("Configuração de Pesos", "*.json")])
        if file_path:
            with open(file_path, "r", encoding="utf-8") as f:
                pesos = json.load(f)
            self.gerenciador_pesos.pesos = defaultdict(lambda: 1.0, {k: float(v) for k, v in pesos.items()})
            self.atualizar_tabela_pesos()
            self.atualizar_kpis()
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Sucesso", "Configuração carregada!")

    def restaurar_pesos_padrao(self):
        if getattr(self, 'inicializando', False):
            return
        dlg = MessageDialog(
            "Você tem certeza que deseja restaurar os valores dos pesos?",
            title="Restaurar Pesos Padrão",
            buttons=["Cancelar:secondary", "Restaurar:danger"],
            default="Cancelar"
        )
        result = dlg.show()
        if result == "Restaurar":
            self.gerenciador_pesos.pesos = defaultdict(lambda: 1.0)
            self.atualizar_tabela_pesos()
            self.atualizar_kpis()
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Restaurado", "Os pesos foram restaurados para o padrão (1.0).")

    def atualizar_kpis(self):
        if getattr(self, 'inicializando', False):
            return

        df_consolidado = self.consolidar_dados_meses()
        if df_consolidado is None or df_consolidado.empty:
            for label in self.kpi_labels.values():
                label.config(text="--")
            return

        col_usuario = "Usuário"
        col_agendamento = "Agendamento"
        col_data = "Data criação"
        col_nroproc = "Nro. processo"
        col_tipo = "Tipo"
        col_status = "Status"
        col_cod = "Código"

        # Total de minutas
        if col_usuario in df_consolidado.columns:
            total_minutas = df_consolidado[col_usuario].count()
            self.kpi_labels["Minutas"].config(text=str(total_minutas))
        else:
            self.kpi_labels["Minutas"].config(text="--")

        # Média produtividade
        if col_usuario in df_consolidado.columns and "peso" in df_consolidado.columns:
            produtividade = df_consolidado.groupby(col_usuario)["peso"].sum()
            media_prod = produtividade.mean() if not produtividade.empty else 0
            self.kpi_labels["Média Produtividade"].config(text=f"{media_prod:.2f}")
            top3 = produtividade.sort_values(ascending=False).head(3)
            if not top3.empty:
                top3_str = "\n".join([f"{i+1}º {u}: {p:.1f}" for i, (u, p) in enumerate(top3.items())])
                self.kpi_labels["Top 3 Usuários"].config(text=top3_str)
            else:
                self.kpi_labels["Top 3 Usuários"].config(text="--")
        else:
            self.kpi_labels["Média Produtividade"].config(text="--")
            self.kpi_labels["Top 3 Usuários"].config(text="--")

        # Dia mais produtivo
        if col_data in df_consolidado.columns and "peso" in df_consolidado.columns:
            df_consolidado[col_data] = pd.to_datetime(df_consolidado[col_data], errors="coerce")
            df_consolidado = df_consolidado.dropna(subset=[col_data])
            dias_trad = {
                'Monday': 'Segunda-feira',
                'Tuesday': 'Terça-feira',
                'Wednesday': 'Quarta-feira',
                'Thursday': 'Quinta-feira',
                'Friday': 'Sexta-feira',
                'Saturday': 'Sábado',
                'Sunday': 'Domingo'
            }
            df_consolidado["DiaSemana"] = df_consolidado[col_data].dt.day_name().map(dias_trad)
            tabela = df_consolidado.pivot_table(index=col_usuario, columns="DiaSemana", values="peso", aggfunc="sum", fill_value=0)
            soma_por_dia = tabela.sum(axis=0)
            if not soma_por_dia.empty:
                dia_mais_prod = soma_por_dia.idxmax()
                self.kpi_labels["Dia Mais Produtivo"].config(text=dia_mais_prod)
            else:
                self.kpi_labels["Dia Mais Produtivo"].config(text="--")
        else:
            self.kpi_labels["Dia Mais Produtivo"].config(text="--")

if __name__ == "__main__":
    import sys
    import os

    root = ttk.Window(themename="flatly")

    # Define o ícone, se não estiver em ambiente frozen (ex: pyinstaller)
    if not getattr(sys, 'frozen', False):
        try:
            root.iconbitmap(os.path.join(os.path.dirname(__file__), "icon.ico"))
        except Exception:
            pass

    root.withdraw()
    splash = SplashScreen(root, theme="flatly")
    splash.update()
    root.update()

    def iniciar_app():
        try:
            start_time = time.time()
            app = ExcelAnalyzerApp(root)
            tema_salvo = app.tema_var.get()
            app.style.theme_use(tema_salvo)
            elapsed = time.time() - start_time
            remaining = max(1.0 - elapsed, 0)
            root.after(int(remaining * 1000), lambda: [
                splash.close(),
                root.deiconify()
            ])
        except Exception as e:
            messagebox.showerror("Erro Fatal", f"Erro na inicialização:\n{str(e)}")
            root.destroy()

    root.after(100, iniciar_app)
    root.mainloop()