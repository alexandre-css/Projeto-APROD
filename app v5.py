import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import MessageDialog
from tkinter import filedialog, messagebox
from tksheet import Sheet
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
try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

def encontrar_linha_cabecalho_excel(filepath, colunas_esperadas):
    import pandas as pd
    for i in range(10):
        try:
            df_temp = pd.read_excel(filepath, header=i)
            cols = [normalizar_coluna(str(c)) for c in df_temp.columns]
            if all(col in cols for col in colunas_esperadas):
                return i
        except Exception:
            continue
    return 0

def encontrar_linha_cabecalho_csv(filepath, colunas_esperadas):
    import pandas as pd
    for i in range(10):
        try:
            df_temp = pd.read_csv(filepath, header=i, sep=None, engine='python')
            cols = [normalizar_coluna(str(c)) for c in df_temp.columns]
            if all(col in cols for col in colunas_esperadas):
                return i
        except Exception:
            continue
    return 0

def normalizar_coluna(nome):
    return ''.join(
        c for c in unicodedata.normalize('NFD', nome)
        if unicodedata.category(c) != 'Mn'
    ).lower().replace(' ', '')

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
        """Carrega pesos do arquivo JSON com tratamento de erros."""
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
        """Salva pesos no arquivo JSON."""
        with open(self.arquivo_pesos, "w", encoding="utf-8") as f:
            json.dump(dict(self.pesos), f, indent=4)

    def atualizar_peso(self, tipo_agendamento, novo_peso):
        try:
            self.pesos[tipo_agendamento] = float(novo_peso)
            self.salvar_pesos()  # Salva automaticamente após atualizar
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
        self.progress.start(12)  # velocidade da animação

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
        self.colunas_norm = {
            "usuario": None,
            "agendamento": None,
            "datacriacao": None
        }
        self.tipos_agendamento_unicos = []
        self.coluna_agendamento = "Agendamento"
        self.style = ttk.Style()
        style = ttk.Style()
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
                anchor="center",  # centraliza vertical e horizontalmente
                justify="center"  # centraliza múltiplas linhas, se houver
            )
            label.pack(expand=True, fill="both")  # faz o label ocupar todo o espaço do frame
            self.kpi_labels[kpi] = label
        self.inicializando = True  # Flag para controle de inicialização
        self.notebook = ttk.Notebook(root, bootstyle=PRIMARY)
        self.notebook.pack(fill="both", expand=True)
        self.criar_aba_analise_graficos()
        self.criar_aba_config_pesos()
        self.criar_aba_relatorios()
        self.criar_aba_produtividade_semana()
        self.criar_aba_comparacao()
        self.criar_aba_comparacao_meses()
        self.criar_aba_configuracoes()
        self.carregar_configuracoes()
        self.gerenciador_pesos = GerenciadorPesosAgendamento()
        self.carregar_pesos_automaticamente()
        self.gerenciador_pesos.carregar_pesos()
        self.root.protocol("WM_DELETE_WINDOW", self.fechar_programa)
        self.inicializando = False  # Finaliza inicialização
        self.primeira_execucao = True  # Nova flag para controle pós-inicialização

    @staticmethod
    def chave_mes_ano(mes_ano):
        meses_pt = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 
                    'jul', 'ago', 'set', 'out', 'nov', 'dez']
        if mes_ano.lower() == "atual":
            # Garante que "Atual" sempre aparece por último
            return (9999, 13)
        try:
            mes_abrev, ano = mes_ano.split('/')
            mes_num = meses_pt.index(mes_abrev.lower()) + 1
            return (int(ano), mes_num)
        except Exception as e:
            print(f"Erro ao converter '{mes_ano}': {str(e)}")
            return (0, 0)  # Coloca itens inválidos no início

    def testar_extracao_meses(self):
        pasta = "dados_mensais"
        if not os.path.exists(pasta):
            print("Pasta dados_mensais não existe!")
            return
        for arq in os.listdir(pasta):
            caminho = os.path.join(pasta, arq)
            print(arq, self.extrair_mes_ano_do_arquivo(caminho))

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
                caminho_arquivo = os.path.join(pasta_dados_mensais, arquivo)
                if os.path.isfile(caminho_arquivo):
                    os.remove(caminho_arquivo)
        self.atualizar_comboboxes_comparacao()

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
        import os
        import re
        import pandas as pd
        from datetime import datetime

        if self.df is None:
            return

        df_save = self.df.copy()
        colunas_disponiveis = list(df_save.columns)
        col_agendamento = next((col for col in colunas_disponiveis if normalizar_coluna(col) == "agendamento"), None)
        if col_agendamento:
            pontuacoes = []
            for _, row in df_save.iterrows():
                if pd.notna(row[col_agendamento]):
                    tipo = re.sub(r'[\r\n]', '', str(row[col_agendamento]).strip())
                    tipo_limpo = re.sub(r'\s*\(.*?\)', '', tipo).strip()
                    peso = self.gerenciador_pesos.obter_peso(tipo_limpo)
                    pontuacoes.append(peso)
                else:
                    pontuacoes.append(0)
            df_save["Peso"] = pontuacoes

        # Remove a coluna Peso antes de salvar para evitar dados obsoletos
        if "Peso" in df_save.columns:
            df_save = df_save.drop(columns=["Peso"])

        # Identifica mês/ano
        colunas_disponiveis = list(df_save.columns)
        col_data = next((col for col in colunas_disponiveis if normalizar_coluna(col) == "datacriacao"), None)
        if col_data is not None:
            datas_validas = pd.to_datetime(df_save[col_data], errors="coerce").dropna()
            if not datas_validas.empty:
                mes = datas_validas.dt.month.mode()[0]
                ano = datas_validas.dt.year.mode()[0]
            else:
                messagebox.showerror("Erro", "A coluna 'Data criação' está vazia ou inválida. Não é possível salvar o mês corretamente.")
                return
        else:
            messagebox.showerror("Erro", "A coluna 'Data criação' não foi encontrada. Não é possível salvar o mês corretamente.")
            return

        pasta = "dados_mensais"
        if not os.path.exists(pasta):
            os.makedirs(pasta)
        nome_arquivo = f"produtividade_{ano}-{mes:02d}.xlsx"
        caminho = os.path.join(pasta, nome_arquivo)
        df_save.to_excel(caminho, index=False)

        # Atualiza comboboxes de comparação, se necessário
        if hasattr(self, 'atualizar_comboboxes_comparacao'):
            self.atualizar_comboboxes_comparacao()

        return caminho

    def carregar_pesos_automaticamente(self):
        """Carrega pesos do arquivo JSON e força atualização."""
        if os.path.exists("pesos.json"):
            with open("pesos.json", "r", encoding="utf-8") as f:
                pesos = json.load(f)
                self.gerenciador_pesos.pesos = defaultdict(lambda: 1.0, {k: float(v) for k, v in pesos.items()})

        # Atualiza tabelas SEM mostrar popups
        if hasattr(self, 'tree_comp_meses'):
            self.comparar_meses(mostrar_popup=False)  # <--- Adicione mostrar_popup=False
        if hasattr(self, 'tree_relatorio'):
            self.gerar_relatorio_produtividade()

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
                self.atualizar_checkboxes_usuarios(),
                self.gerar_relatorio_produtividade()

    def fechar_programa(self):
        # Bloqueia popups durante o fechamento
        self.inicializando = True  # <--- Adicione esta linha

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
        self.treeview = ttk.Treeview(frame_esquerdo, bootstyle=INFO)
        self.treeview.pack(fill="both", expand=True, padx=5, pady=5)
        frame_controles = ttk.LabelFrame(frame_esquerdo, text="Configurações do Gráfico", padding=10, bootstyle=INFO)
        frame_controles.pack(fill="x", padx=5, pady=5)
        ttk.Label(frame_controles, text="Coluna para análise:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combo_colunas = ttk.Combobox(frame_controles, state="readonly", width=30, bootstyle=INFO)
        self.combo_colunas.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(frame_controles, text="Tipo de gráfico:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.graph_type = ttk.StringVar(value="barras horizontais")
        self.graph_type_map = {
            "barras verticais": "bar",
            "barras horizontais": "barh",
            "pizza": "pie",
            "em linha": "line"
        }
        graph_combo = ttk.Combobox(
            frame_controles,
            textvariable=self.graph_type,
            values=list(self.graph_type_map.keys()),
            state="readonly",
            width=15,
            bootstyle=INFO
        )
        graph_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(frame_controles, text="Estilo visual:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.style_var = ttk.StringVar(value="Solarize_Light2")
        self.style_map = STYLE_OPTIONS
        style_combo = ttk.Combobox(
            frame_controles,
            textvariable=self.style_var,
            values=list(self.style_map.keys()),
            state="readonly",
            width=20,
            bootstyle=INFO
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
        if os.path.exists(pasta):
            meses = []
            for arq in os.listdir(pasta):
                caminho = os.path.join(pasta, arq)
                mes_ano = self.extrair_mes_ano_do_arquivo(caminho)
                if mes_ano != "Mês desconhecido":
                    meses.append(mes_ano)
            if meses:
                self.label_meses_carregados.config(text="Meses carregados: " + ", ".join(sorted(meses, key=self.chave_mes_ano)))
            else:
                self.label_meses_carregados.config(text="Meses carregados: --")
        else:
            self.label_meses_carregados.config(text="Meses carregados: --")

    def criar_aba_config_pesos(self):
        frame_pesos = ttk.Frame(self.notebook)
        self.notebook.add(frame_pesos, text="Configuração de Pesos")

        ttk.Label(
            frame_pesos,
            text="Pesos para cada tipo de Agendamento:",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", padx=10, pady=(10,0))

        # Frame para centralizar e expandir a tabela
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
        self.tree_pesos.pack(fill="both", expand=True)

        # Controles para atualizar pesos
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

        # Botões para salvar/carregar/restaurar configuração
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

    def criar_aba_configuracoes(self):
        frame_config = ttk.Frame(self.notebook)
        self.notebook.add(frame_config, text="Configurações")

        # Variáveis de controle
        self.tema_var = ttk.StringVar(value="flatly")
        self.fonte_var = ttk.StringVar(value="Pequeno")
        self.export_path_var = ttk.StringVar(value="")
        self.auto_save_var = ttk.BooleanVar(value=True)

        # Tema do programa
        temas = ["morph", "flatly", "darkly", "cyborg", "journal", "solar", "vapor", "superhero"]
        ttk.Label(frame_config, text="Tema do programa:", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=16, pady=(16, 4))
        combo_tema = ttk.Combobox(
            frame_config,
            textvariable=self.tema_var,
            values=temas,
            state="readonly",
            width=20,
            bootstyle=INFO
        )
        combo_tema.pack(anchor="w", padx=16, pady=(0, 16))

        def mudar_tema(event=None):
            self.style.theme_use(self.tema_var.get())
        combo_tema.bind("<<ComboboxSelected>>", mudar_tema)

        # Pasta padrão para exportação
        ttk.Label(frame_config, text="Pasta padrão para exportação:").pack(anchor="w", padx=16, pady=(8, 2))
        ttk.Entry(frame_config, textvariable=self.export_path_var, width=40).pack(anchor="w", padx=16)
        ttk.Button(
            frame_config,
            text="Procurar",
            command=self.selecionar_pasta_exportacao,
            bootstyle=SECONDARY
        ).pack(anchor="w", padx=16, pady=(2, 8))

        # Tamanho da fonte
        ttk.Label(frame_config, text="Tamanho da fonte:").pack(anchor="w", padx=16, pady=(8, 2))
        self.fonte_var = ttk.StringVar(value="Pequeno")
        combo_fonte = ttk.Combobox(
            frame_config,
            textvariable=self.fonte_var,
            values=["Pequeno", "Médio", "Grande"],
            state="readonly"
        )
        combo_fonte.pack(anchor="w", padx=16)

        font_sizes = {
            "Pequeno": 10,
            "Médio": 14,
            "Grande": 18
        }

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

        # Salvar configurações automaticamente
        ttk.Checkbutton(
            frame_config,
            text="Salvar configurações automaticamente ao sair",
            variable=self.auto_save_var
        ).pack(anchor="w", padx=16, pady=(0, 2))

        # Restaurar padrão
        ttk.Button(
            frame_config,
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
        # ... resto do método ...
        if getattr(self, 'inicializando', False):
            return
        messagebox.showinfo("Configurações", "Configurações restauradas para o padrão!")

    def selecionar_pasta_exportacao(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.export_path_var.set(pasta)

    def criar_aba_relatorios(self):
        frame_relatorios = ttk.Frame(self.notebook)
        self.notebook.add(frame_relatorios, text="Relatórios de Produtividade")
        frame_filtros = ttk.LabelFrame(frame_relatorios, text="Filtros", padding=10, bootstyle=INFO)
        frame_filtros.pack(fill="x", padx=10, pady=5)
        ttk.Label(frame_filtros, text="Coluna de usuário:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Label(frame_filtros, text="Usuário").grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Button(frame_filtros, text="Gerar Relatório de Produtividade", command=self.gerar_relatorio_produtividade, bootstyle=SUCCESS).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(frame_filtros, text="Visualizar Gráficos de Produtividade", command=self.visualizar_graficos_produtividade, bootstyle=PRIMARY).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame_filtros, text="Exportar para Excel", command=self.exportar_relatorio_excel, bootstyle=SECONDARY).grid(row=1, column=2, padx=5, pady=5)
        ttk.Button(frame_filtros, text="Exportar para PDF", command=self.exportar_relatorio_pdf, bootstyle=SECONDARY).grid(row=1, column=3, padx=5, pady=5)
        self.tree_relatorio = ttk.Treeview(frame_relatorios, columns=("Usuário", "Pontuação"), show="headings", bootstyle=INFO)
        self.tree_relatorio.heading("Usuário", text="Usuário", command=lambda: self.ordenar_treeview_relatorio("Usuário", False))
        self.tree_relatorio.heading("Pontuação", text="Pontuação Total", command=lambda: self.ordenar_treeview_relatorio("Pontuação", False))
        self.tree_relatorio.column("Usuário", width=200, anchor="center", stretch=False)
        self.tree_relatorio.column("Pontuação", width=150, anchor="center", stretch=False)
        self.tree_relatorio.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree_relatorio.column("Usuário", width=200, anchor="w")
        for col in self.tree_relatorio["columns"][1:]:
            self.tree_relatorio.column(col, width=120, anchor="center")

    def criar_aba_produtividade_semana(self):
        frame_semana = ttk.Frame(self.notebook)
        self.notebook.add(frame_semana, text="Produtividade por Dia da Semana")
        frame_top = ttk.Frame(frame_semana)
        frame_top.pack(fill="x", padx=10, pady=10)
        ttk.Label(frame_top, text="Coluna de usuário: Usuário").pack(side="left", padx=5)
        ttk.Label(frame_top, text="Coluna de data: Data criação").pack(side="left", padx=20)
        ttk.Button(frame_top, text="Gerar Tabela", command=self.gerar_tabela_semana, bootstyle=SUCCESS).pack(side="left", padx=20)
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

        ttk.Button(
            frame_btn,
            text="Comparar",
            command=self.comparar_produtividade_usuarios,
            bootstyle=SUCCESS
        ).pack(side="left", padx=5)

        ttk.Button(
            frame_btn,
            text="Comparar Todos",
            command=self.comparar_todos_meses,
            bootstyle=SUCCESS
        ).pack(side="left", padx=5)

        frame_tabela = ttk.LabelFrame(frame_comp, text="Comparação de Usuários", padding=10, bootstyle=ttk.INFO)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)

        # Inicializa o Sheet (tksheet) para comparação de usuários
        self.sheet_comp = Sheet(
            frame_tabela,
            data=[[""]],  # Inicialmente vazio
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


    def atualizar_comboboxes_comparacao(self):
        import os
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta)] if os.path.exists(pasta) else []
        arquivos = [arq for arq in arquivos if arq.endswith(".xlsx") or arq.endswith(".csv")]
        opcoes_amigaveis = []
        for arq in arquivos:
            caminho = os.path.join(pasta, arq)
            mes_ano = self.extrair_mes_ano_do_arquivo(caminho)
            opcoes_amigaveis.append(mes_ano)
        self.mapa_arquivo_meses = dict(zip(opcoes_amigaveis, arquivos))

        # Atualiza o Treeview
        if hasattr(self, 'listbox_meses'):
            self.listbox_meses.delete(0, 'end')
            for opcao in opcoes_amigaveis:
                self.listbox_meses.insert('end', opcao)

    def criar_aba_comparacao_meses(self):
        import os
        from tksheet import Sheet
        
    # Frame principal da aba
        frame_principal = ttk.Frame(self.notebook)
        self.notebook.add(frame_principal, text="Comparação de Meses")
        frame_principal.pack(fill="both", expand=True)

        # Frame principal da aba
        frame_principal = ttk.Frame(self.notebook)
        self.notebook.add(frame_principal, text="Comparação de Meses")
        frame_principal.pack(fill="both", expand=True)

        # Listar arquivos disponíveis
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta)] if os.path.exists(pasta) else []
        arquivos = [arq for arq in arquivos if arq.endswith((".xlsx", ".csv"))]

        opcoes_amigaveis = []
        for arq in arquivos:
            caminho = os.path.join(pasta, arq)
            mes_ano = self.extrair_mes_ano_do_arquivo(caminho)
            opcoes_amigaveis.append(mes_ano)
        self.mapa_arquivo_meses = dict(zip(opcoes_amigaveis, arquivos))

        # Container principal
        frame_superior = ttk.Frame(frame_principal)
        frame_superior.pack(fill="x", padx=10, pady=10)

        # Listbox para seleção múltipla
        ttk.Label(frame_superior, text="Selecione os meses para comparar:").pack(anchor="w")
        self.listbox_meses = tk.Listbox(
            frame_superior, 
            selectmode="multiple", 
            height=8, 
            exportselection=False,
            font=("Segoe UI", 11)
        )
        # Scrollbar
        scroll = ttk.Scrollbar(frame_superior, orient="vertical", command=self.listbox_meses.yview)
        scroll.pack(side="right", fill="y")
        self.listbox_meses.configure(yscrollcommand=scroll.set)
        for opcao in opcoes_amigaveis:
            self.listbox_meses.insert("end", opcao)
        self.listbox_meses.pack(fill="x", expand=True)

        # Botões
        frame_botoes = ttk.Frame(frame_superior)
        frame_botoes.pack(fill="x", pady=10)

        ttk.Button(
            frame_botoes,
            text="Comparar",
            command=self.comparar_meses,
            bootstyle=SUCCESS
        ).pack(side="left", padx=5)

        ttk.Button(
            frame_botoes,
            text="Comparar Todos",
            command=self.comparar_todos_meses,
            bootstyle=SUCCESS
        ).pack(side="left", padx=5)

        ttk.Button(
            frame_botoes,
            text="Excluir dados dos meses",
            command=self.limpar_dados_anteriores,
            bootstyle=DANGER
        ).pack(side="right", padx=5)

        # KPIs
        self.kpi_comp_frame = ttk.Frame(frame_principal)
        self.kpi_comp_frame.pack(fill="x", padx=10, pady=5)
        self.kpi_comp_labels = {}

        # Frame da tabela 
        frame_tabela = ttk.LabelFrame(frame_principal, text="Comparação Detalhada", padding=0)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=10)
        frame_tabela.pack_propagate(False)  # Impede ajuste automático de tamanho

        # IMPORTANTE: Usar o método grid em vez de place ou pack
        frame_tabela.columnconfigure(0, weight=1)  # Coluna expandível
        frame_tabela.rowconfigure(0, weight=1)     # Linha expandível

        self.sheet_comp_meses = Sheet(
            frame_tabela,
            data=[[""]],
            headers=["Usuário"],
            theme="light blue",
            show_x_scrollbar=True,
            show_y_scrollbar=True
        )

        # Use grid com sticky="nsew" para ocupar todo o espaço
        self.sheet_comp_meses.grid(row=0, column=0, sticky="nsew")

        # Outros bindings permanecem iguais
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

    def comparar_todos_meses(self):
        self.listbox_meses.selection_set(0, "end")
        self.comparar_meses()

    def verificar_arquivos_mensais(self):
        pasta = "dados_mensais"
        if os.path.exists(pasta):
            print("Arquivos encontrados em dados_mensais:")
            for arq in os.listdir(pasta):
                caminho = os.path.join(pasta, arq)
                print(f"- {arq} -> {self.extrair_mes_ano_do_arquivo(caminho)}")
        else:
            print("Pasta dados_mensais não existe!")

    def excluir_dados_comparacao(self):
        # Aqui você deve apagar os dados das tabelas da aba "Comparação de meses"
        # Supondo que você tenha Treeviews ou widgets onde os dados são mostrados, limpe-os
        
        # Exemplo genérico para limpar Treeviews:
        if hasattr(self, 'treeview_comparacao_mes'):
            for item in self.treeview_comparacao_mes.get_children():
                self.treeview_comparacao_mes.delete(item)
        
        if hasattr(self, 'treeview_comparacao_mes2'):
            for item in self.treeview_comparacao_mes2.get_children():
                self.treeview_comparacao_mes2.delete(item)
        
        # Se tiver outras tabelas ou widgets, limpe aqui também

    def sort_treeview_comp(self, col, reverse):
        items = [(self.tree_comp_meses.set(k, col), k) for k in self.tree_comp_meses.get_children('')]
        try:
            items = [(float(val.replace(',', '')), k) for val, k in items]
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
            df = pd.read_excel(os.path.join("dados_mensais", arquivo))
            col_usuario_mes = next((col for col in df.columns if normalizar_coluna(col) == "usuario"), None)
            col_agendamento_mes = next((col for col in df.columns if normalizar_coluna(col) == "agendamento"), None)
            if col_usuario_mes and col_agendamento_mes:
                df = df.dropna(subset=[col_usuario_mes, col_agendamento_mes])
                df[col_usuario_mes] = df[col_usuario_mes].astype(str).str.strip()
                df["TipoAgendamentoLimpo"] = df[col_agendamento_mes].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
                df["Peso"] = df["TipoAgendamentoLimpo"].apply(self.gerenciador_pesos.obter_peso)
                prod = df.groupby(col_usuario_mes)[["Peso"]].sum()
            else:
                prod = pd.DataFrame(columns=["Peso"])
            dados_meses.append(prod)

        todos_usuarios = sorted({usuario for prod in dados_meses for usuario in prod.index})
        colunas = ["Usuário"] + meses_selecionados + ["Total"]

        # Preparar dados para o tksheet
        dados_tabela = []
        celulas_verdes = []    # flecha para cima
        celulas_vermelhas = [] # flecha para baixo
        celulas_top3 = []      # top 3 valores por mês
        celulas_total = []     # coluna de total
        celulas_zeradas = []   # valores zerados
        linhas_alternadas = [] # linhas alternadas

        # Monta tabela e marca flechas
        for idx, usuario in enumerate(todos_usuarios):
            linha = [usuario]
            valores = []
            total = 0
            for i, prod in enumerate(dados_meses):
                valor = prod.loc[usuario]["Peso"] if usuario in prod.index else 0
                valores.append(valor)
                total += valor

            for i, v in enumerate(valores):
                if i == 0:
                    linha.append(f"{v:.2f}")
                    if v == 0:
                        celulas_zeradas.append((idx, i+1))
                else:
                    if v > valores[i-1]:
                        linha.append(f"{v:.2f} ▲")
                        celulas_verdes.append((idx, i+1))
                    elif v < valores[i-1]:
                        linha.append(f"{v:.2f} ▼")
                        celulas_vermelhas.append((idx, i+1))
                    else:
                        linha.append(f"{v:.2f}")
                    if v == 0:
                        celulas_zeradas.append((idx, i+1))
            linha.append(f"{total:.2f}")
            dados_tabela.append(linha)
            celulas_total.append((idx, len(colunas)-1))
            if idx % 2 == 0:
                linhas_alternadas.append(idx)

        # Top 3 valores de cada mês (coluna)
        for col in range(1, len(colunas)-1):  # ignora usuário e total
            valores_col = [(row, float(dados_tabela[row][col].split()[0])) for row in range(len(dados_tabela))]
            top3 = sorted(valores_col, key=lambda x: x[1], reverse=True)[:3]
            for row, _ in top3:
                celulas_top3.append((row, col))

        # Atualiza o tksheet
        self.sheet_comp_meses.headers(colunas)
        self.sheet_comp_meses.set_sheet_data(dados_tabela)

        # Coluna de total (azul claro)
        for row, col in celulas_total:
            self.sheet_comp_meses.highlight_cells(row=row, column=col, bg="#e3f2fd", fg="#1565c0")

        # Top 3 valores por mês (amarelo claro)
        for row, col in celulas_top3:
            self.sheet_comp_meses.highlight_cells(row=row, column=col, bg="#fff9c4", fg="#795548")

        # Flecha para cima (verde forte)
        for row, col in celulas_verdes:
            self.sheet_comp_meses.highlight_cells(row=row, column=col, bg="#d0f5e8", fg="green")

        # Flecha para baixo (vermelho forte)
        for row, col in celulas_vermelhas:
            self.sheet_comp_meses.highlight_cells(row=row, column=col, bg="#ffebee", fg="red")

        # Ajustar largura das colunas
        self.sheet_comp_meses.set_all_column_widths(120)
        self.sheet_comp_meses.column_width(0, 200)

        # Atualiza KPIs
        totais = [prod["Peso"].sum() for prod in dados_meses]
        for lbl in self.kpi_comp_labels.values():
            lbl.destroy()
        self.kpi_comp_labels = {}
        for i, (mes, total) in enumerate(zip(meses_selecionados, totais)):
            lbl_texto = f"Total {mes}: {total:.2f}"
            lbl = ttk.Label(self.kpi_comp_frame, text=lbl_texto, font=("Segoe UI", 12, "bold"))
            lbl.pack(side="left", padx=16)
            self.kpi_comp_labels[mes] = lbl

        if len(totais) >= 2:
            primeiro = totais[0]
            ultimo = totais[-1]
            variacao = ((ultimo - primeiro) / primeiro * 100) if primeiro != 0 else 0
            lbl_variacao = ttk.Label(
                self.kpi_comp_frame,
                text=f"Variação Total (%): {variacao:+.1f}%",
                font=("Segoe UI", 12, "bold"),
                foreground="green" if variacao > 0 else "red" if variacao < 0 else "black"
            )
            lbl_variacao.pack(side="left", padx=16)
            self.kpi_comp_labels["Variação (%)"] = lbl_variacao



    def comparar_todos_meses(self):
        # Seleciona todos os meses na listbox
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
                messagebox.showwarning("Aviso", "O arquivo não possui dados válidos!")
            return
        col_usuario = next((col for col in self.df.columns if normalizar_coluna(col) == "usuario"), None)
        if not col_usuario:
            return
        usuarios = sorted(self.df[col_usuario].dropna().unique())
        for usuario in usuarios:
            var = ttk.BooleanVar(value=False)
            chk = ttk.Checkbutton(self.frame_check_users, text=usuario, variable=var, bootstyle=INFO)
            chk.pack(anchor="w")
            self.check_vars_usuarios[usuario] = var
            self.checkbuttons_usuarios[usuario] = chk
            self.checkbuttons_usuarios[usuario] = chk

    def comparar_produtividade_usuarios(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Aviso", "Carregue os dados primeiro!")
            return
        col_usuario = next((col for col in self.df.columns if normalizar_coluna(col) == "usuario"), None)
        col_agendamento = next((col for col in self.df.columns if normalizar_coluna(col) == "agendamento"), None)
        if not col_usuario or not col_agendamento:
            messagebox.showwarning("Aviso", "Carregue os dados e garanta que há as colunas 'Usuário' e 'Agendamento'.")
            return

        selecoes_meses = self.listbox_meses.curselection()
        if not selecoes_meses:
            messagebox.showwarning("Aviso", "Selecione meses na aba 'Comparação de Meses' primeiro!")
            return

        meses_selecionados = [self.listbox_meses.get(i) for i in selecoes_meses]
        arquivos_selecionados = [self.mapa_arquivo_meses[mes] for mes in meses_selecionados]

        usuarios_selecionados = [u for u, var in self.check_vars_usuarios.items() if var.get()]
        if not usuarios_selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um usuário para comparar.")
            return

        # Processar dados de cada mês
        from collections import defaultdict
        dados = defaultdict(lambda: defaultdict(float))
        for mes, arquivo in zip(meses_selecionados, arquivos_selecionados):
            try:
                df_mes = pd.read_excel(os.path.join("dados_mensais", arquivo))
                col_usuario_mes = next((col for col in df_mes.columns if normalizar_coluna(col) == "usuario"), None)
                col_agendamento_mes = next((col for col in df_mes.columns if normalizar_coluna(col) == "agendamento"), None)
                if col_usuario_mes and col_agendamento_mes:
                    df_mes = df_mes.dropna(subset=[col_usuario_mes, col_agendamento_mes])
                    df_mes[col_usuario_mes] = df_mes[col_usuario_mes].astype(str).str.strip()
                    df_mes["TipoAgendamentoLimpo"] = df_mes[col_agendamento_mes].apply(
                        lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip()
                    )
                    df_mes["Peso"] = df_mes["TipoAgendamentoLimpo"].apply(self.gerenciador_pesos.obter_peso)
                    for usuario in usuarios_selecionados:
                        if usuario in df_mes[col_usuario_mes].values:
                            dados[usuario][mes] = df_mes[df_mes[col_usuario_mes] == usuario]["Peso"].sum()
                        else:
                            dados[usuario][mes] = 0.0
                else:
                    for usuario in usuarios_selecionados:
                        dados[usuario][mes] = 0.0
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {str(e)}")
                continue

        # Monta dados e aplica flechas e cores
        colunas = ["Usuário"] + meses_selecionados + ["Total"]
        dados_tabela = []
        celulas_verdes = []
        celulas_vermelhas = []

        for idx, usuario in enumerate(usuarios_selecionados):
            linha = [usuario]
            valores = []
            total = 0.0
            for mes in meses_selecionados:
                valor = dados[usuario][mes]
                valores.append(valor)
                total += valor

            valores_fmt = []
            for i, v in enumerate(valores):
                if i == 0:
                    valores_fmt.append(f"{v:.2f}")
                else:
                    if v > valores[i-1]:
                        valores_fmt.append(f"{v:.2f} ▲")
                        celulas_verdes.append((idx, i+1))  # +1 porque primeira coluna é usuário
                    elif v < valores[i-1]:
                        valores_fmt.append(f"{v:.2f} ▼")
                        celulas_vermelhas.append((idx, i+1))
                    else:
                        valores_fmt.append(f"{v:.2f}")
            linha.extend(valores_fmt)
            linha.append(f"{total:.2f}")
            dados_tabela.append(linha)

        # Preenche e colore o Sheet
        self.sheet_comp.headers(colunas)
        self.sheet_comp.set_sheet_data(dados_tabela)

        for linha, coluna in celulas_verdes:
            self.sheet_comp.highlight_cells(row=linha, column=coluna, bg="lightgreen", fg="black")
        for linha, coluna in celulas_vermelhas:
            self.sheet_comp.highlight_cells(row=linha, column=coluna, bg="lightcoral", fg="black")

        self.sheet_comp.set_all_column_widths(120)
        self.sheet_comp.set_column_width(0, 200)



    def gerar_tabela_semana(self):
        df_consolidado = self.consolidar_dados_meses()
        if df_consolidado is None or df_consolidado.empty:
            if not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "Não há dados mensais carregados!")
            return
        colunas_disponiveis = list(df_consolidado.columns)
        col_usuario = next((col for col in colunas_disponiveis if normalizar_coluna(col) == "usuario"), None)
        col_data = next((col for col in colunas_disponiveis if normalizar_coluna(col) == "datacriacao"), None)
        col_agendamento = next((col for col in colunas_disponiveis if normalizar_coluna(col) == "agendamento"), None)
        if not col_usuario or not col_data or not col_agendamento:
            messagebox.showwarning("Aviso", "A planilha precisa ter as colunas 'Usuário', 'Data criação' e 'Agendamento'.")
            return

        df = df_consolidado.copy()

        df[col_data] = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")
        df = df.dropna(subset=[col_usuario, col_data, col_agendamento])

        # Tradução dos dias da semana para português
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

        df["TipoAgendamentoLimpo"] = df[col_agendamento].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
        df["Peso"] = df["TipoAgendamentoLimpo"].apply(self.gerenciador_pesos.obter_peso)

        tabela = df.pivot_table(
            index=col_usuario,
            columns="DiaSemana",
            values="Peso",
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
            valores = [usuario] + [row.get(d, 0) for d in dias_semana_pt]
            self.tree_semana.insert("", "end", values=valores)

        soma_por_dia = tabela.sum(axis=0)
        ranking = soma_por_dia.sort_values(ascending=False)
        ranking_str = "\n".join(
            [f"{i+1}º - {dia}: {ranking[dia]:.2f}" for i, dia in enumerate(ranking.index)]
        )
        self.label_ranking.config(
            text=ranking_str,
            font=("Segoe UI", 14, "bold"),
            anchor="center",
            justify="center",
            foreground="#3a3a3a"
        )

        self.atualizar_kpis()

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
           l.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
        except ValueError:
           l.sort(key=lambda t: t[0], reverse=reverse)

        for index, (val, k) in enumerate(l):
           self.tree_pesos.move(k, '', index)

        # Atualiza cabeçalhos com sorting reverso na próxima chamada
        self.tree_pesos


    def ordenar_treeview_relatorio(self, col, reverse):
        l = [(self.tree_relatorio.set(k, col), k) for k in self.tree_relatorio.get_children('')]
        try:
            # Tenta ordenar como número (float)
            l.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
        except ValueError:
            # Se não for número, ordena como texto
            l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree_relatorio.move(k, '', index)
        self.tree_relatorio.heading(col, command=lambda: self.ordenar_treeview_relatorio(col, not reverse))

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if file_path:
            self.file_path.set(file_path)
            self.carregar_dados()

    def carregar_dados(self):
        try:
            if getattr(self, 'inicializando', False):
                return
            if not self.file_path.get():
                if not getattr(self, 'inicializando', False):
                    messagebox.showwarning("Aviso", "Selecione um arquivo primeiro!")
                return
            file_path = self.file_path.get()

            # Conversão automática de .xls para .xlsx
            if file_path.lower().endswith('.xls') and not file_path.lower().endswith('.xlsx'):
                novo_path = os.path.splitext(file_path)[0] + ".xlsx"
                if not os.path.exists(novo_path):
                    try:
                        x2x = XLS2XLSX(file_path)
                        x2x.to_xlsx(novo_path)
                    except Exception as conv_err:
                        messagebox.showerror("Erro", f"Erro ao converter arquivo .xls para .xlsx:\n{conv_err}")
                        self.dados_carregados = False
                        return
                file_path = novo_path

            colunas_esperadas = ["usuario", "agendamento", "datacriacao"]

            if file_path.endswith('.csv'):
                linha_cabecalho = encontrar_linha_cabecalho_csv(file_path, colunas_esperadas)
                self.df = pd.read_csv(file_path, header=linha_cabecalho, sep=None, engine='python')
            else:
                linha_cabecalho = encontrar_linha_cabecalho_excel(file_path, colunas_esperadas)
                self.df = pd.read_excel(file_path, header=linha_cabecalho)

            self.df = self.df.loc[:, ~self.df.columns.str.contains('^Unnamed')]

            colunas_disponiveis = list(self.df.columns)
            self.combo_colunas['values'] = colunas_disponiveis

            colunas_necessarias_norm = {
                "usuario": None,
                "agendamento": None,
                "datacriacao": None
            }

            for col in colunas_disponiveis:
                nome_norm = normalizar_coluna(col)
                if nome_norm in colunas_necessarias_norm:
                    colunas_necessarias_norm[nome_norm] = col

            faltando = [orig for orig, real in colunas_necessarias_norm.items() if real is None]

            if faltando:
                nomes_legiveis = {
                    "usuario": "Usuário",
                    "agendamento": "Agendamento",
                    "datacriacao": "Data criação"
                }
                faltando_legiveis = [nomes_legiveis[f] for f in faltando]
                messagebox.showerror("Erro", f"As colunas obrigatórias estão faltando: {', '.join(faltando_legiveis)}")
                self.dados_carregados = False
                return

            if not colunas_disponiveis:
                if not getattr(self, 'inicializando', False):
                    messagebox.showwarning("Aviso", "O arquivo não possui colunas válidas!")
                return

            usuario_col = next((col for col in colunas_disponiveis if normalizar_coluna(col) in ["usuario"]), None)

            if usuario_col:
                self.combo_colunas.set(usuario_col)

            # Garante que a coluna 'Peso' existe
            if "Peso" not in self.df.columns:
                self.df["Peso"] = 0.0

            col_agendamento = colunas_necessarias_norm["agendamento"]

            if col_agendamento in self.df.columns:
                tipos = self.df[col_agendamento].dropna().unique().tolist()
                tipos_limpos = [re.sub(r'\s*\(.*?\)', '', str(t)).strip() for t in tipos]
                self.tipos_agendamento_unicos = sorted(set(tipos_limpos))

                pontuacoes = []
                for _, row in self.df.iterrows():
                    if pd.notna(row[col_agendamento]):
                        tipo = re.sub(r'\s*\(.*?\)', '', str(row[col_agendamento])).strip()
                        peso = self.gerenciador_pesos.obter_peso(tipo)
                        pontuacoes.append(peso)
                    else:
                        pontuacoes.append(0)
                self.df["Peso"] = pontuacoes

            # Validação da coluna 'Data criação'
            col_data = next((col for col in self.df.columns if normalizar_coluna(col) == "datacriacao"), None)

            if col_data:
                self.df[col_data] = pd.to_datetime(self.df[col_data], dayfirst=True, errors="coerce")
                if self.df[col_data].isna().all():
                    if not getattr(self, 'inicializando', False):
                        messagebox.showerror("Erro", "Não foi possível converter a coluna 'Data criação' para data.")
                    self.dados_carregados = False
                    return

            # Atualizações após carregar dados
            self.atualizar_treeview()
            self.atualizar_kpis()
            self.atualizar_checkboxes_usuarios()
            self.atualizar_tabela_pesos()
            self.gerar_relatorio_produtividade()
            self.gerar_tabela_semana()
            if hasattr(self, "atualizar_comboboxes_comparacao"):
                self.atualizar_comboboxes_comparacao()
            if hasattr(self, "comparar_meses"):
                self.comparar_meses(mostrar_popup=False)
            self.salvar_dados_mes()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{str(e)}")
            return
        messagebox.showinfo("Sucesso", "Dados carregados com sucesso!")
        self.testar_extracao_meses()
        self.atualizar_label_meses_carregados()

    def extrair_mes_ano_do_arquivo(self, caminho):
        try:
            colunas_esperadas = ["datacriacao"]
            if caminho.endswith('.csv'):
                linha_cabecalho = encontrar_linha_cabecalho_csv(caminho, colunas_esperadas)
                df = pd.read_csv(caminho, header=linha_cabecalho, sep=None, engine='python')
            else:
                linha_cabecalho = encontrar_linha_cabecalho_excel(caminho, colunas_esperadas)
                df = pd.read_excel(caminho, header=linha_cabecalho)
            col_data = next((col for col in df.columns if normalizar_coluna(col) == "datacriacao"), None)
            if col_data is not None:
                datas = pd.to_datetime(df[col_data], errors="coerce").dropna()
                if not datas.empty:
                    mes = datas.dt.month.mode()[0]
                    ano = datas.dt.year.mode()[0]
                    meses_pt = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                                'jul', 'ago', 'set', 'out', 'nov', 'dez']
                    return f"{meses_pt[mes-1]}/{ano}"
        except Exception as e:
            print(f"Erro ao extrair mês/ano do arquivo {caminho}: {e}")
        return "Mês desconhecido"


    def atualizar_treeview(self):
        if self.df is None:
            return
        for i in self.treeview.get_children():
            self.treeview.delete(i)
        cols = list(self.df.columns)
        self.treeview["columns"] = cols
        self.treeview["show"] = "headings"
        for col in cols:
            self.treeview.heading(col, text=col)
            self.treeview.column(col, width=130, anchor="center", stretch=False)
        for _, row in self.df.head(100).iterrows():
            self.treeview.insert("", "end", values=[row[col] for col in cols])

    def atualizar_tabela_pesos(self):
        print("Atualizando tabela de pesos...")
        ordem_atual = [self.tree_pesos.item(item)['values'][0] for item in self.tree_pesos.get_children()]
        if not ordem_atual:
            ordem_atual = sorted(self.tipos_agendamento_unicos)
        for i in self.tree_pesos.get_children():
            self.tree_pesos.delete(i)
        for tipo in ordem_atual:
            peso = self.gerenciador_pesos.obter_peso(tipo)
            self.tree_pesos.insert("", "end", values=(tipo, peso))

    def atualizar_peso_selecionado(self):
        selecionados = self.tree_pesos.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione um ou mais tipos de agendamento primeiro!")
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
        self.atualizar_tabela_pesos()
        self.atualizar_kpis()
        self.gerar_relatorio_produtividade()
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
        messagebox.showinfo("Sucesso", "Pesos atualizados e tabelas recarregadas!")

    def generate_graph(self):
        if self.inicializando:
            return

        if self.df is None or self.df.empty:
            if not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "O arquivo não possui dados válidos!")
            return
        coluna = self.combo_colunas.get()
        if not coluna:
            if not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "Selecione uma coluna para análise!")
            return
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
                ax.text(bar.get_x() + bar.get_width()/2, height, f'{height}', ha='center', va='bottom', fontweight='bold')
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
            ax.set_xlabel('Quantidade')
            for i, v in enumerate(counts.values):
                ax.text(i, v, f"{v}", ha='center', va='bottom', fontweight='bold')
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')

        self.fig.tight_layout()
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.frame_grafico)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

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
        if getattr(self, 'inicializando', False):
            return
        self.tree_relatorio.delete(*self.tree_relatorio.get_children())
        for label in self.kpi_labels.values():
            label.config(text="--")
        dados_usuarios = defaultdict(lambda: defaultdict(float))

        # Inclui o arquivo carregado
        if self.df is not None and not self.df.empty:
            col_usuario = next((col for col in self.df.columns if normalizar_coluna(col) == "usuario"), None)
            col_agendamento = next((col for col in self.df.columns if normalizar_coluna(col) == "agendamento"), None)
            if col_usuario and col_agendamento:
                self.df[col_usuario] = self.df[col_usuario].astype(str).str.strip()
                self.df["TipoAgendamentoLimpo"] = self.df[col_agendamento].apply(
                    lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip()
                )
                self.df["Peso"] = self.df["TipoAgendamentoLimpo"].apply(self.gerenciador_pesos.obter_peso)
                for usuario, grupo in self.df.groupby(col_usuario):
                    dados_usuarios[usuario]["Atual"] = grupo["Peso"].sum()

        # Inclui arquivos mensais salvos
        pasta_mensal = "dados_mensais"
        arquivos_meses = [f for f in os.listdir(pasta_mensal) if f.endswith(('.xlsx', '.csv'))] if os.path.exists(pasta_mensal) else []
        for arquivo in arquivos_meses:
            try:
                caminho = os.path.join(pasta_mensal, arquivo)
                mes = self.extrair_mes_ano_do_arquivo(caminho)
                df_mes = pd.read_excel(caminho)
                col_usuario_mes = next((col for col in df_mes.columns if normalizar_coluna(col) == "usuario"), None)
                col_agendamento_mes = next((col for col in df_mes.columns if normalizar_coluna(col) == "agendamento"), None)
                if col_usuario_mes and col_agendamento_mes:
                    df_mes = df_mes.dropna(subset=[col_usuario_mes, col_agendamento_mes])
                    df_mes[col_usuario_mes] = df_mes[col_usuario_mes].astype(str).str.strip()
                    df_mes["TipoAgendamentoLimpo"] = df_mes[col_agendamento_mes].apply(
                        lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip()
                    )
                    df_mes["Peso"] = df_mes["TipoAgendamentoLimpo"].apply(self.gerenciador_pesos.obter_peso)
                    for usuario, grupo in df_mes.groupby(col_usuario_mes):
                        dados_usuarios[usuario][mes] = grupo["Peso"].sum()
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {str(e)}")
                continue

        # Monta as colunas do relatório
        meses = sorted({mes for user in dados_usuarios.values() for mes in user.keys()}, key=self.chave_mes_ano)
        if "Atual" in meses:
            meses.remove("Atual")
            meses.append("Atual")
        colunas = ["Usuário"] + meses + ["Total"]

        self.tree_relatorio["columns"] = colunas
        for col in colunas:
            self.tree_relatorio.heading(
                col,
                text=col,
                command=lambda c=col: self.ordenar_treeview_relatorio(c, False)
            )
            self.tree_relatorio.column(col, width=120, anchor="center")

        # Insere dados
        for usuario, meses_data in sorted(dados_usuarios.items(),
                                         key=lambda x: sum(x[1].values()),
                                         reverse=True):
            total = sum(meses_data.values())
            linha = [usuario]
            for mes in meses:
                linha.append(f"{meses_data.get(mes, 0):.2f}")
            linha.append(f"{total:.2f}")
            self.tree_relatorio.insert("", "end", values=linha)

        self.atualizar_kpis()


    def mostrar_mensagem(tipo, titulo, mensagem):
        if tipo == "erro":
            messagebox.showerror(titulo, mensagem)
        elif tipo == "aviso":
            messagebox.showwarning(titulo, mensagem)
        elif tipo == "info":
            messagebox.showinfo(titulo, mensagem)

    def consolidar_dados_meses(self):
        import os
        import pandas as pd
        import re

        dfs = []

        # Inclui o arquivo atualmente carregado na interface
        if self.df is not None and not self.df.empty:
            df_atual = self.df.copy()
            col_usuario = next((col for col in df_atual.columns if normalizar_coluna(col) == "usuario"), None)
            col_agendamento = next((col for col in df_atual.columns if normalizar_coluna(col) == "agendamento"), None)
            if col_usuario and col_agendamento:
                df_atual = df_atual.dropna(subset=[col_usuario, col_agendamento])
                df_atual[col_usuario] = df_atual[col_usuario].astype(str).str.strip()
                df_atual["TipoAgendamentoLimpo"] = df_atual[col_agendamento].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
                df_atual["Peso"] = df_atual["TipoAgendamentoLimpo"].apply(self.gerenciador_pesos.obter_peso)
                dfs.append(df_atual)

        # Inclui os arquivos salvos em dados_mensais
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta) if arq.endswith(('.xlsx', '.csv'))] if os.path.exists(pasta) else []
        for arquivo in arquivos:
            caminho = os.path.join(pasta, arquivo)
            try:
                df_mes = pd.read_excel(caminho)
                col_usuario_mes = next((col for col in df_mes.columns if normalizar_coluna(col) == "usuario"), None)
                col_agendamento_mes = next((col for col in df_mes.columns if normalizar_coluna(col) == "agendamento"), None)
                if col_usuario_mes and col_agendamento_mes:
                    df_mes = df_mes.dropna(subset=[col_usuario_mes, col_agendamento_mes])
                    df_mes[col_usuario_mes] = df_mes[col_usuario_mes].astype(str).str.strip()
                    df_mes["TipoAgendamentoLimpo"] = df_mes[col_agendamento_mes].apply(lambda x: re.sub(r'\s*\(.*?\)', '', str(x)).strip())
                    df_mes["Peso"] = df_mes["TipoAgendamentoLimpo"].apply(self.gerenciador_pesos.obter_peso)
                    dfs.append(df_mes)
            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")
        if dfs:
            return pd.concat(dfs, ignore_index=True)
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
            import numpy as np

            # Coleta dados do Treeview de relatório
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

            # Cria a janela do gráfico
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
                    if getattr(self, 'inicializando', False):
                        return
                    messagebox.showwarning("Aviso", "Gere um gráfico primeiro!")
                    return
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".png",
                    initialfile="produtividade.png",
                    filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("PDF", "*.pdf")]
                )
                if file_path:
                    self.fig_produtividade.savefig(file_path, bbox_inches='tight', dpi=300)
                    if getattr(self, 'inicializando', False):
                        return
                    messagebox.showinfo("Sucesso", f"Gráfico salvo em:\n{file_path}")

            salvar_btn.config(command=salvar_grafico_produtividade)
            gerar_btn.config(command=gerar_grafico_produtividade)
            gerar_grafico_produtividade()


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
            ax.set_title("Produtividade por Usuário", fontsize=22, fontweight='bold', fontname='DejaVu Sans', color='#1a237e', pad=25)
            fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            self.fig_produtividade = fig

        def salvar_grafico_produtividade():
            if self.fig_produtividade is None:
                if getattr(self, 'inicializando', False):
                    return
                messagebox.showwarning("Aviso", "Gere um gráfico primeiro!")
                return
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                initialfile="produtividade.png",
                filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg"), ("PDF", "*.pdf")]
            )
            if file_path:
                self.fig_produtividade.savefig(file_path, bbox_inches='tight', dpi=300)
                if getattr(self, 'inicializando', False):
                    return
                messagebox.showinfo("Sucesso", f"Gráfico salvo em:\n{file_path}")

        salvar_btn.config(command=salvar_grafico_produtividade)
        gerar_btn.config(command=gerar_grafico_produtividade)
        gerar_grafico_produtividade()

    def exportar_relatorio_excel(self):
        if self.df is None or self.df.empty or not self.tree_relatorio.get_children():
            if not getattr(self, 'inicializando', False):
                messagebox.showwarning("Aviso", "Gere o relatório primeiro!")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])
        if file_path:
            dados = [self.tree_relatorio.item(item)['values'] for item in self.tree_relatorio.get_children()]
            df_export = pd.DataFrame(dados, columns=["Usuário", "Pontuação"])
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
            pdf.cell(60, 10, "Usuário", border=1)
            pdf.cell(40, 10, "Pontuação", border=1)
            pdf.ln()
            for usuario, pontuacao in dados:
                pdf.cell(60, 10, str(usuario), border=1)
                pdf.cell(40, 10, str(pontuacao), border=1)
                pdf.ln()
            pdf.output(file_path)
            if getattr(self, 'inicializando', False):
                return
            messagebox.showinfo("Sucesso", f"Relatório exportado para:\n{file_path}")

    def salvar_configuracao_pesos(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Configuração de Pesos", "*.json")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(self.gerenciador_pesos.pesos, f, ensure_ascii=False, indent=2)
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
            from collections import defaultdict
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

        col_usuario = next((col for col in df_consolidado.columns if normalizar_coluna(col) == "usuario"), "Usuário")
        col_data = next((col for col in df_consolidado.columns if normalizar_coluna(col) == "datacriacao"), None)

        # Minutas: total de processos (linhas)
        minutas = len(df_consolidado)
        self.kpi_labels["Minutas"].config(text=str(minutas))

        # Produtividade total por usuário (soma dos pesos)
        produtividade = df_consolidado.groupby(col_usuario)["Peso"].sum()
        media_prod = produtividade.mean() if not produtividade.empty else 0
        self.kpi_labels["Média Produtividade"].config(text=f"{media_prod:.2f}")

        # Dia mais produtivo: usa exatamente o mesmo cálculo da aba "Produtividade por Dia da Semana"
        if col_data:
            # Converte a coluna de data para o formato datetime
            df_consolidado[col_data] = pd.to_datetime(df_consolidado[col_data], errors="coerce")

            # Remove linhas com valores nulos na coluna de data
            df_consolidado = df_consolidado.dropna(subset=[col_data])

            # Mapeia os dias da semana para português
            dias_trad = {
                'Monday': 'Segunda-feira',
                'Tuesday': 'Terça-feira',
                'Wednesday': 'Quarta-feira',
                'Thursday': 'Quinta-feira',
                'Friday': 'Sexta-feira',
                'Saturday': 'Sábado',
                'Sunday': 'Domingo'
            }

            # Cria a coluna "DiaSemana" com os dias da semana em português
            df_consolidado["DiaSemana"] = df_consolidado[col_data].dt.day_name().map(dias_trad)

            # Cria a tabela pivot com a soma dos pesos por dia da semana
            tabela = df_consolidado.pivot_table(index=col_usuario, columns="DiaSemana", values="Peso", aggfunc="sum", fill_value=0)

            # Calcula a soma dos pesos por dia da semana
            soma_por_dia = tabela.sum(axis=0)

            # Determina o dia mais produtivo
        if not soma_por_dia.empty:
            dia_mais_prod = soma_por_dia.idxmax()
            self.kpi_labels["Dia Mais Produtivo"].config(text=dia_mais_prod)
        else:
            self.kpi_labels["Dia Mais Produtivo"].config(text="--")

        # Top 3 Usuários (soma da produtividade)
        top3 = produtividade.sort_values(ascending=False).head(3)
        if not top3.empty:
            top3_str = "\n".join([f"{i+1}º {u}: {p:.1f}" for i, (u, p) in enumerate(top3.items())])
            self.kpi_labels["Top 3 Usuários"].config(text=top3_str)
        else:
            self.kpi_labels["Top 3 Usuários"].config(text="--")

if __name__ == "__main__":
    import sys
    import os

    root = ttk.Window(themename="flatly")

    if not getattr(sys, 'frozen', False):
        root.iconbitmap(os.path.join(os.path.dirname(__file__), "icon.ico"))

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
            app.style.theme_use(app.tema_var.get())
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