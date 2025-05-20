from PyQt6.QtCore import QObject, pyqtSignal, pyqtProperty, pyqtSlot
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox, QMainWindow, QListWidget
from PyQt6.QtQml import QQmlApplicationEngine, QQmlContext
from xls2xlsx import XLS2XLSX
from app_v7 import GerenciadorPesosAgendamento, carregar_planilha, MainWindow
import sys, os, shutil
import pandas as pd
import re
import json
import datetime

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
    mesesAtivosChanged = pyqtSignal()

    def __init__(self, mainwindow):
        super().__init__()
        self.mainwindow = mainwindow
        self._nomes = []
        self._valores = []
        self._arquivosCarregados = ""
        self._mesesAtivos = []

    @pyqtProperty('QStringList', notify=mesesAtivosChanged)
    def mesesAtivos(self):
        return self._mesesAtivos

    
    @pyqtProperty('QStringList', notify=arquivosCarregadosChanged)
    def mesesDisponiveis(self):
        """Retorna todos os meses únicos processados"""
        pasta = "dados_mensais"
        periodos = []
        
        if os.path.exists(pasta):
            arquivos = [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            print(f"⚠️ Arquivos encontrados na pasta: {len(arquivos)}")
            
            for arquivo in arquivos:
                file_path = os.path.join(pasta, arquivo)
                periodo = self.mainwindow.extrair_mes_ano_do_arquivo(file_path)
                if periodo and "/" in periodo:
                    periodos.append(periodo)
                    print(f"Adicionando período: {periodo}")
        
        # Retornar lista ordenada de períodos únicos
        result = sorted(list(set(periodos)), key=self.mainwindow.chave_mes_ano)
        print(f"⚠️ Total de períodos para exibição: {len(result)}, meses = {result}")
        return result



    @pyqtSlot(str)
    def toggleMesAtivo(self, mes):
        """Ativa/desativa um mês na visualização"""
        if mes in self._mesesAtivos:
            self._mesesAtivos.remove(mes)
        else:
            self._mesesAtivos.append(mes)
        
        # Emitir sinal para atualizar os visuais
        self.mesesAtivosChanged.emit()
        
        # Atualizar dados na ordem correta
        print(f"Atualizando dados após toggle de {mes}, meses ativos: {self._mesesAtivos}")
        
        # IMPORTANTE: Ordem correta de atualização dos componentes
        self.atualizar_kpis()
        self.atualizar_grafico()
        self.atualizar_tabela_pesos()
        self.atualizar_grafico()
        
        # Forçar emissão dos sinais para garantir que a UI atualize
        self.valoresChanged.emit()
        self.nomesChanged.emit()
        if hasattr(self.mainwindow, "kpis_qml") and self.mainwindow.kpis_qml is not None:
            self.mainwindow.kpis_qml.kpisChanged.emit()


    @pyqtProperty('QStringList', notify=nomesChanged)
    def nomes(self):
        return self._nomes

    @pyqtProperty(list, notify=valoresChanged)
    def valores(self):
        return self._valores

    @pyqtProperty(list, notify=tabelaPesosChanged)
    def tabela_pesos(self):
        return [
            {'tipo': nome, 'peso': valor}
            for nome, valor in zip(self._nomes, self._valores)
        ]

    @pyqtProperty(str, notify=arquivosCarregadosChanged)
    def arquivosCarregados(self):
        return self._arquivosCarregados

    @pyqtProperty('QStringList', notify=nomesChanged)
    def nomes(self):
        return self._nomes

    @pyqtProperty(list, notify=valoresChanged)
    def valores(self):
        return self._valores

    @pyqtProperty(list, notify=tabelaPesosChanged)
    def tabela_pesos(self):
        return [
            {'tipo': nome, 'peso': valor}
            for nome, valor in zip(self._nomes, self._valores)
        ]

    @pyqtProperty(str, notify=arquivosCarregadosChanged)
    def arquivosCarregados(self):
        return self._arquivosCarregados
    
    def _filtrar_arquivos_por_meses_ativos(self, arquivos):
        """Filtra arquivos pelos meses ativos"""
        if not self._mesesAtivos:
            # CORREÇÃO: Quando nenhum mês está ativo, mostrar TODOS os arquivos
            # em vez de retornar lista vazia []
            return arquivos  
        
        result = []
        for arq in arquivos:
            file_path = os.path.join("dados_mensais", arq)
            periodo = self.mainwindow.extrair_mes_ano_do_arquivo(file_path)
            if periodo and periodo.lower() in [m.lower() for m in self._mesesAtivos]:
                result.append(arq)
        
        print(f"Filtragem: {len(arquivos)} → {len(result)} arquivos")
        return result



    @pyqtSlot()
    def atualizar_kpis(self):
        pasta = "dados_mensais"
        arquivos = (
            [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            if os.path.exists(pasta)
            else []
        )
        
        # Filtrar pelos meses ativos
        arquivos = self._filtrar_arquivos_por_meses_ativos(arquivos)
        
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
        dia_qml = "--"
        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            # Filtrar linhas onde usuário é NaN ou string vazia
            df_total = df_total[df_total["Usuário"].notna() & (df_total["Usuário"] != "")]
            
            # Agrupamento por Usuário
            produtividade = df_total.groupby("Usuário")["peso"].sum()
            
            # Remover índices NaN do resultado do groupby
            produtividade = produtividade[produtividade.index.notna()]
            
            # Remover strings "NAN" explicitamente
            produtividade = produtividade[produtividade.index.str.upper() != "NAN"]
            
            media = produtividade.mean() if not produtividade.empty else 0
            media_qml = f"{media:.2f}"
            top3 = produtividade.sort_values(ascending=False).head(3)
            if not top3.empty:
                top3_qml = "\n".join([f"{u}: {p:.1f}" for u, p in top3.items()])
            else:
                top3_qml = "--"
        elif "Usuário" in df_total.columns:
            # Filtrar linhas onde usuário é NaN ou string vazia
            df_total = df_total[df_total["Usuário"].notna() & (df_total["Usuário"] != "")]
            
            contagem = df_total.groupby("Usuário").size()
            
            # Remover índices NaN do resultado do groupby
            contagem = contagem[contagem.index.notna()]
            
            # Remover strings "NAN" explicitamente
            contagem = contagem[contagem.index.str.upper() != "NAN"]
            
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
        """Atualiza a lista de arquivos carregados e meses disponíveis"""
        pasta = "dados_mensais"
        if not os.path.exists(pasta):
            self._arquivosCarregados = "Nenhum arquivo carregado"
            self._mesesAtivos = []
        else:
            arquivos = sorted([arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")])
            periodos = []
            for arq in arquivos:
                file_path = os.path.join(pasta, arq)
                periodo = self.mainwindow.extrair_mes_ano_do_arquivo(file_path)
                if periodo and "/" in periodo:
                    periodos.append(periodo)
            
            self._arquivosCarregados = " - ".join(periodos) if periodos else "Nenhum arquivo carregado"
            
            # Inicializa todos os meses como ativos se ainda não houver seleção
            if not hasattr(self, "_mesesAtivos") or not self._mesesAtivos:
                self._mesesAtivos = periodos.copy()
                    
        self.arquivosCarregadosChanged.emit()
        self.mesesAtivosChanged.emit()


    @pyqtSlot()
    def atualizar_grafico(self):
        pasta = "dados_mensais"
        arquivos = (
            [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]
            if os.path.exists(pasta)
            else []
        )
        
        # Filtrar pelos meses ativos
        arquivos = self._filtrar_arquivos_por_meses_ativos(arquivos)
        
        dfs = []
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            df = self.mainwindow.carregar_planilha(file_path)
            if "Usuário" in df.columns:
                df["Usuário"] = df["Usuário"].astype(str).str.strip().str.upper()
            df = df.loc[:, ~df.columns.duplicated()]
            if not df.empty:
                dfs.append(df)
        
        # Quando não há dados, mostrar arrays vazios
        if not dfs:
            self._nomes = []
            self._valores = []
            self.nomesChanged.emit()
            self.valoresChanged.emit()
            return

        df_total = pd.concat(dfs, ignore_index=True)
        df_total = df_total.loc[:, ~df_total.columns.duplicated()]
        
        # CORREÇÃO: Garantir limpeza de dados antes de usar
        if "Usuário" in df_total.columns and "peso" in df_total.columns:
            # Filtrar linhas onde usuário é NaN ou string vazia
            df_total = df_total[df_total["Usuário"].notna() & (df_total["Usuário"] != "")]
            
            # Garantir que só agrupa por usuário e peso, não por agendamento
            produtividade = df_total.groupby("Usuário")["peso"].sum()
            
            # Remover índices NaN do resultado do groupby
            produtividade = produtividade[produtividade.index.notna()]
            
            produtividade = produtividade.sort_values(ascending=False)
            self._nomes = list(produtividade.index)
            self._valores = [float(v) for v in produtividade.values]
        # ...resto do método

        elif "Usuário" in df_total.columns:
            produtividade = df_total.groupby("Usuário").size()
            
            # Remover índices NaN do resultado do groupby
            produtividade = produtividade[produtividade.index.notna()]
            
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
        self.atualizar_kpis
        self.atualizar_grafico
        self.atualizar_tabela_pesos

    def _coletar_tipos_brutos(self):
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")] if os.path.exists(pasta) else []
        
        tipos = set()
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            try:
                df = pd.read_excel(file_path, skiprows=1, usecols="G")
                if "Agendamento" in df.columns:
                    # Coleta sem processamento
                    tipos.update(df["Agendamento"].dropna().astype(str).str.strip().tolist())
            except Exception as e:
                print(f"Erro coletando tipos brutos: {str(e)}")
        return sorted(tipos)


    #& Aqui começam os métodos de atribuição de pesos.

    @pyqtSlot()
    def atualizar_tabela_pesos(self):
        tipos_agendamento = []
        
        # Coletar todos os tipos de agendamento
        pasta = "dados_mensais"
        arquivos = [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")] if os.path.exists(pasta) else []
        
        # ALTERAÇÃO AQUI: Filtrar pelos meses ativos
        arquivos = self._filtrar_arquivos_por_meses_ativos(arquivos)
        
        for arq in arquivos:
            file_path = os.path.join(pasta, arq)
            try:
                df_raw = pd.read_excel(file_path, skiprows=1, usecols="G")
                
                if len(df_raw.columns) > 0:
                    df_raw.columns = ["Agendamento"]
                    
                    tipos = df_raw["Agendamento"].dropna().astype(str)
                    tipos = tipos[tipos.str.strip() != ""]
                    tipos = tipos.apply(lambda x: re.sub(r'\s*(\(\d+\)|\d+)$', '', x).strip())
                    
                    tipos_agendamento.extend(tipos.unique())
            except Exception as e:
                print(f"Erro processando {file_path}: {str(e)}")
        
        tipos_agendamento = sorted(list(set(tipos_agendamento)))
        pesos = [self.mainwindow.gerenciador_pesos.obter_peso(tipo) for tipo in tipos_agendamento]
        
        self._nomes = tipos_agendamento
        self._valores = pesos
        self.tabelaPesosChanged.emit()
        self.nomesChanged.emit()
        self.valoresChanged.emit()


    @pyqtSlot(int, float)
    def atualizarPeso(self, index, novo_peso):
        if hasattr(self.mainwindow, "gerenciador_pesos"):
            nome = self.nomes[index] if 0 <= index < len(self.nomes) else None
            if nome:
                # Atualiza e força salvamento imediato
                self.mainwindow.gerenciador_pesos.pesos[nome] = float(novo_peso)
                self.mainwindow.gerenciador_pesos.salvar_pesos()  # Novo
                self._valores[index] = float(novo_peso)
                self.valoresChanged.emit()
    
    @pyqtSlot()
    def salvarPesos(self):
        """Salva os pesos atuais em novo arquivo com timestamp na pasta 'pesos salvos' do projeto"""
        try:
            # Caminho absoluto da pasta do script principal
            pasta_base = os.path.dirname(os.path.abspath(__file__))
            pasta_pesos = os.path.join(pasta_base, "pesos salvos")
            os.makedirs(pasta_pesos, exist_ok=True)

            # Gerar nome do arquivo com timestamp
            timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            arquivo = os.path.join(pasta_pesos, f"pesos_{timestamp}.json")

            # Salvar cópia dos pesos atuais
            with open(arquivo, 'w', encoding='utf-8') as f:
                json.dump(dict(self.mainwindow.gerenciador_pesos.pesos), f, ensure_ascii=False, indent=4)

            # Feedback para o usuário
            QMessageBox.information(self.mainwindow, "Pesos Salvos", f"Arquivo salvo em:\n{arquivo}")

        except Exception as e:
            print(f"ERRO AO SALVAR PESOS: {str(e)}")
            QMessageBox.critical(self.mainwindow, "Erro", f"Falha ao salvar pesos:\n{str(e)}")

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
        if hasattr(self.mainwindow, 'gerenciador_pesos'):
            # 1. Usar os tipos de agendamento atuais (não buscar tipos novos)
            tipos_atuais = self._nomes  # Mantem os nomes atuais na tabela
            
            # 2. Reset no gerenciador de pesos
            self.mainwindow.gerenciador_pesos.pesos = {
                tipo: 1.0 for tipo in tipos_atuais
            }
            self.mainwindow.gerenciador_pesos.salvar_pesos()
            
            # 3. Atualizar apenas os valores sem mexer nos nomes
            self._valores = [1.0] * len(tipos_atuais)
            
            # 4. Forçar sincronização na interface
            self.tabelaPesosChanged.emit()
            self.valoresChanged.emit()

    @pyqtSlot(int, float)
    def atualizarPeso(self, index, novo_peso):
        if hasattr(self.mainwindow, "gerenciador_pesos"):
            nome = self._nomes[index] if 0 <= index < len(self._nomes) else None
            if nome:
                self.mainwindow.gerenciador_pesos.pesos[nome] = float(novo_peso)
                self.mainwindow.gerenciador_pesos.salvar_pesos()
                self._valores[index] = float(novo_peso)
                self.valoresChanged.emit()

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

    def closeEvent(self, event):
        try:
            if hasattr(self, 'gerenciador_pesos'):
                self.gerenciador_pesos.salvar_pesos()
        except Exception as e:
            print(f"Erro ao salvar pesos ao fechar: {str(e)}")
        finally:
            super().closeEvent(event)


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
            return
        pasta = "dados_mensais"
        os.makedirs(pasta, exist_ok=True)
        arquivos_importados = 0
        for path in paths:
            try:
                nome_destino = os.path.basename(path)
                destino = os.path.join(pasta, nome_destino)
                # Se for .xls, converte para .xlsx e salva apenas o .xlsx
                if path.lower().endswith(".xls"):
                    x2x = XLS2XLSX(path)
                    novo_destino = os.path.splitext(destino)[0] + ".xlsx"
                    x2x.to_xlsx(novo_destino)
                    arquivos_importados += 1
                else:
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
        if hasattr(self, "backend_qml") and self.backend_qml is not None:
            self.backend_qml.atualizar_grafico()
            self.backend_qml.atualizar_tabela_pesos()
            self.backend_qml.atualizar_kpis()
            self.backend_qml.atualizar_arquivos_carregados()


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
        try:
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
        except Exception as e:
            print(f"Erro ao carregar planilha {file_path}: {str(e)}")
            return pd.DataFrame()


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
    def __init__(self, arquivo_pesos=None):
        if arquivo_pesos is None:
            pasta_destino = os.path.join(os.path.expanduser("~"), "Analyzer-Dev-APROD", "salvos")
            os.makedirs(pasta_destino, exist_ok=True)
            arquivo_pesos = os.path.join(pasta_destino, "pesos.json")
        
        self.arquivo_pesos = arquivo_pesos
        self.pesos = {}
        self.carregar_pesos()

    def restaurar_pesos_padrao(self, tipos_atuais: list):
        self.pesos = {tipo: 1.0 for tipo in tipos_atuais}  # Sobrescreve completamente
        self.salvar_pesos()  # Persistência imediata


    def carregar_pesos(self):
        try:
            with open(self.arquivo_pesos, 'r', encoding='utf-8') as f:
                dados = json.load(f)
                self.pesos = dict(dados)  # ✅ Converte para dict normal
        except (FileNotFoundError, json.JSONDecodeError):
            self.pesos = {}

    def salvar_pesos(self):
        try:
            with open(self.arquivo_pesos, 'w', encoding='utf-8') as f:
                json.dump(dict(self.pesos), f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"ERRO CRÍTICO AO SALVAR PESOS: {str(e)}")
            raise

    def _valor_padrao(self):
        return 1.0

    def obter_peso(self, tipo_agendamento):
        return self.pesos.get(tipo_agendamento, 1.0)


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        app.setApplicationName("KPI App")
        engine = QQmlApplicationEngine()

        # Criar a pasta de dados se não existir
        os.makedirs("dados_mensais", exist_ok=True)

        mainwindow = MainWindow()
        backend = Backend(mainwindow)
        kpis = KPIs()
        mainwindow.kpis_qml = kpis
        mainwindow.backend_qml = backend

        # Definir propriedades de contexto para QML
        engine.rootContext().setContextProperty("mainwindow", mainwindow)
        engine.rootContext().setContextProperty("backend", backend)
        engine.rootContext().setContextProperty("kpis", kpis)

        # Inicializar dados em ordem segura
        backend.atualizar_arquivos_carregados()  # Primeiro os arquivos e meses
        # NÃO force a leitura de mesesDisponiveis aqui - isso causará problemas
        backend.atualizar_grafico()  # Depois os gráficos
        backend.atualizar_kpis()     # Por fim os KPIs

        # Carregar o QML depois da inicialização dos dados
        qml_file = os.path.join(os.path.dirname(__file__), "MainWindow.qml")
        engine.load(qml_file)
        
        if not engine.rootObjects():
            print("ERROR: Falha ao carregar arquivo QML!")
            sys.exit(1)
            
        sys.exit(app.exec())
    except Exception as e:
        print(f"ERRO DE INICIALIZAÇÃO: {str(e)}")
        import traceback
        traceback.print_exc()


