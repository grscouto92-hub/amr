import pandas as pd
import os
import warnings
import time
import unicodedata
import re
import json
import webbrowser
from datetime import datetime
from decimal import Decimal, InvalidOperation

# Silencia avisos do Excel/Pandas
warnings.simplefilter(action='ignore', category=UserWarning)


class PowerBIQualidadeETL:

    def __init__(self):

        # ==========================================
        # 1. CONFIGURAÇÕES DE DIRETÓRIO E ARQUIVOS
        # ==========================================

        self.pasta_base = (
            r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS"
            r"\CX Qualidade - Qualidade - PBI\2. Novo PowerBI\Genesys"
        )

        self.arquivo_genesys = os.path.join(self.pasta_base, "dados_genesys.xlsx")
        self.arquivo_excel = os.path.join(self.pasta_base, "data.xlsx")
        self.arquivo_exclusoes = os.path.join(self.pasta_base, "monitoriasexcluidas.xlsx")
        self.arquivo_csat = os.path.join(self.pasta_base, "dados_csat.xlsx")
        self.arquivo_itens = os.path.join(self.pasta_base, "base1.xlsx")

        self.erro_excel = None
        self.erro_genesys = None
        self.erro_exclusoes = None
        self.erro_csat = None
        self.erro_itens = None

        # ==========================================
        # EQUIPE DE QUALIDADE
        # ==========================================

        self.time_qualidade = {
            "SUZANE SALES GALVEIA",
            "MAURICIO NOVAES SANTOS",
            "DENISE DE OLIVEIRA BASSO",
            "ALINE DOS SANTOS MALAQUIAS",
            "THAIS SOUZA LEAO",
            "GUILHERME RODRIGUES DA SILVA COUTO",
            "LUAN SOUZA SILVA"
        }

        # ==========================================
        # FILTROS DE ESCOPO DO RELATÓRIO
        # (comparação sem acento, sem espaço/pontuação e em minúsculo)
        # ==========================================

        self.formularios_permitidos = {
            self._limpar(x) for x in [
                "[Américas] Texto",
                "[Américas] Voz",
                "Américas -Texto",
                "Américas -Voz",
                "Texto - Rede Américas",
                "Voz - Rede Américas",
                "Voz - Rede Amériacas",
            ]
        }

        self.status_permitido = "FINISHED"

        self.gerentes_permitidos = {
            self._limpar(x) for x in [
                "Fabiano Feliciano da Silva",
                "Ricardo Del Bianchi Gomes",
            ]
        }

        # Monitoria só vale se feita DEPOIS de 30 dias da admissão (> 30)
        self.dias_casa_minimo = 30

        # Alocações que são desconsideradas APENAS na contagem de Base de HC (data.xlsx)
        # (As monitorias da tabela Fato NÃO são excluídas por alocação)
        self.alocacoes_excluidas_hc = {
            self._limpar(x) for x in [
                "TERCEIRO",
                "TERCEIROS",
                "TERCEIRIZADO",
            ]
        }

        # Cargos que NÃO são elegíveis para monitoria na Base de HC (data.xlsx)
        self.cargos_excluidos_hc = {
            self._limpar(x) for x in [
                "SUPERVISOR ATENDIMENTO CLIENTE",
                "COORDENADOR ATENDIMENTO CLIENTE",
            ]
        }

        # =====================================================================
        # GRUPOS DE PERGUNTAS EXATOS (CONDUÇÃO vs. NCG)
        # =====================================================================

        self.grupos_conducao_permitidos = {
            self._limpar(x) for x in [
                "Condução do Atendimento",
                "Condução (NC)",
                "Condução",
            ]
        }

        self.grupos_ncg_permitidos = {
            self._limpar(x) for x in [
                "Não Conformidade Grave",
                "NCG (Não Conformidades Graves",
                "NCG (Não Conformidades Graves)",
                "Não Conformidade Grave (NCG)",
                "NCG",
            ]
        }

        self.df_fato = pd.DataFrame()
        self.df_dim = pd.DataFrame()
        self.df_exclusoes = pd.DataFrame()
        self.df_csat = pd.DataFrame()
        self.df_fatos_itens = pd.DataFrame()
        self.df_ofensores = pd.DataFrame()
        self.mapa_admissao = {}
        self.mapa_alocacao_login = {}

        # Diagnósticos
        self.diag_bonus = {}
        self.diag_filtros = {}
        self.diag_itens = {}
        self.catalogo_grupos = []

    # ============================================================
    # FUNÇÕES AUXILIARES (BÔNUS DE TABULAÇÃO E NORMALIZAÇÃO)
    # ============================================================

    @staticmethod
    def _limpar(texto):
        """Minúsculo, sem acento, sem espaços/pontuação. 'Não Conforme' -> 'naoconforme'."""
        if pd.isna(texto):
            return ""
        try:
            t = str(texto).strip().lower()
            t = "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")
            return re.sub(r"[\s\W_]+", "", t)
        except Exception:
            return str(texto).strip().lower()

    @staticmethod
    def _norm_id(valor):
        """
        Normaliza IDs para o cruzamento:
        - remove espaços e '.0' no final
        - converte notação científica (1.23E+15) para número inteiro por extenso
        - minúsculo (UUIDs do Genesys não divergem por caixa)
        """
        if pd.isna(valor):
            return ""
        s = str(valor).strip()
        if s.lower() in ("nan", "none", "nat", "n/a", ""):
            return ""
        if re.fullmatch(r"[0-9]+([.,][0-9]+)?[eE][+-]?[0-9]+", s):
            try:
                s = format(Decimal(s.replace(",", ".")), "f")
            except InvalidOperation:
                pass
        s = re.sub(r"\.0+$", "", s)
        return s.lower()

    def _ajustar_supervisor_mariele(self, coordenador, supervisor):
        """
        Se a coordenadora for MARIELE e estiver sem nome de supervisor,
        repete o nome dela (MARIELE).
        """
        coord_str = str(coordenador).strip() if pd.notna(coordenador) else ""
        sup_str = str(supervisor).strip() if pd.notna(supervisor) else ""

        coord_limpo = self._limpar(coord_str)
        sup_limpo = self._limpar(sup_str)

        valores_sem_supervisor = {
            "", "nan", "none", "nat", "na", "semgestor",
            "semsupervisor", "naocadastradonabase", "naocadastrado", "-"
        }

        if "mariele" in coord_limpo and sup_limpo in valores_sem_supervisor:
            return coord_str if coord_str else "MARIELE"

        return sup_str if sup_str else "Sem Gestor"

    def _eh_alocacao_terceiro(self, texto):
        """Retorna True se o valor da alocação na base de HC for Terceiro."""
        limpo = self._limpar(texto)
        if not limpo:
            return False
        return (limpo in self.alocacoes_excluidas_hc) or ("terceir" in limpo)

    def _eh_cargo_excluido(self, texto):
        """
        Retorna True se o cargo for de liderança não elegível para monitoria:
        - SUPERVISOR ATENDIMENTO CLIENTE
        - COORDENADOR ATENDIMENTO CLIENTE
        """
        limpo = self._limpar(texto)
        if not limpo:
            return False
        return limpo in self.cargos_excluidos_hc

    def _achar_col_id_avaliacao(self, df):
        """Coluna 'id' + 'aval' que NÃO seja a do avaliador."""
        return next(
            (c for c in df.columns
             if "id" in str(c).lower() and "aval" in str(c).lower() and "avaliador" not in str(c).lower()),
            None
        )

    def _achar_col_grupo_perguntas(self, df):
        """
        Localiza a coluna 'Grupo de perguntas' olhando tanto o nome da coluna
        quanto a presença dos valores 'Condução do Atendimento' / 'Não Conformidade Grave'.
        """
        candidatas_nome = [
            c for c in df.columns
            if "grupo" in self._limpar(c) and "id" not in self._limpar(c)
        ]

        melhor_col = None
        melhor_score = -1

        colunas_busca = candidatas_nome if candidatas_nome else list(df.columns)
        alvos = self.grupos_conducao_permitidos | self.grupos_ncg_permitidos

        for c in colunas_busca:
            serie_limpa = df[c].apply(self._limpar)
            score = int(serie_limpa.isin(alvos).sum())
            if score > melhor_score:
                melhor_col = c
                melhor_score = score

        if melhor_score > 0:
            return melhor_col

        return candidatas_nome[0] if candidatas_nome else None

    def _escolher_coluna_por_conteudo(self, df, palavra_no_nome, termo_no_valor):
        """
        Escolhe a coluna de texto certa olhando o CONTEÚDO, não só o nome.
        """
        def nome_norm(c):
            return self._limpar(c)

        def eh_coluna_tecnica(c):
            n = str(c).lower()
            n_sem_acento = "".join(ch for ch in unicodedata.normalize("NFD", n) if unicodedata.category(ch) != "Mn")
            if re.search(r"(^|[^a-z])(id|cod|codigo)([^a-z]|$)", n_sem_acento):
                return True
            return any(x in n_sem_acento for x in ("grupo", "peso", "pontu", "nota", "score"))

        def pontuar(cols):
            melhor_col, melhor_score = None, -1
            for c in cols:
                score = int(df[c].apply(self._limpar).str.contains(termo_no_valor, na=False).sum())
                if score > melhor_score:
                    melhor_col, melhor_score = c, score
            return melhor_col, max(melhor_score, 0)

        candidatas = [
            c for c in df.columns
            if palavra_no_nome in nome_norm(c) and not eh_coluna_tecnica(c)
        ]

        col, score = pontuar(candidatas) if candidatas else (None, 0)

        if score == 0:
            todas = [c for c in df.columns if not eh_coluna_tecnica(c)]
            col2, score2 = pontuar(todas) if todas else (None, 0)
            if score2 > 0:
                return col2, score2
            return (candidatas[0] if candidatas else None), 0

        return col, score

    def _classificar_por_grupo_de_perguntas(self, texto_grupo):
        """
        Classifica com 100% de precisão usando o nome do Grupo de perguntas:
        - NCG: 'Não Conformidade Grave', 'NCG (Não Conformidades Graves',
               'NCG (Não Conformidades Graves)', 'Não Conformidade Grave (NCG)'
        - Condução: 'Condução do Atendimento'
        """
        g_limpo = self._limpar(texto_grupo)

        if not g_limpo:
            return "Condução"

        if g_limpo in self.grupos_ncg_permitidos or "naoconformidadegrave" in g_limpo or "ncg" in g_limpo:
            return "NCG"

        if g_limpo in self.grupos_conducao_permitidos or "conducao" in g_limpo:
            return "Condução"

        return "Condução"

    # ============================================================
    # 2. EXTRAÇÃO
    # ============================================================

    def extrair_dados(self):
        try:
            self.df_fato = pd.read_excel(self.arquivo_genesys, dtype=str)
        except Exception as e:
            self.erro_genesys = str(e)

        try:
            self.df_dim = pd.read_excel(self.arquivo_excel, dtype=str)
            self.df_dim.columns = self.df_dim.columns.astype(str).str.strip().str.upper()
        except Exception as e:
            self.erro_excel = str(e)

        try:
            self.df_csat = pd.read_excel(self.arquivo_csat, dtype=str)
        except Exception as e:
            self.erro_csat = str(e)

        try:
            if os.path.exists(self.arquivo_itens):
                self.df_fatos_itens = pd.read_excel(self.arquivo_itens, dtype=str)
            elif os.path.exists(self.arquivo_genesys):
                xls_genesys = pd.ExcelFile(self.arquivo_genesys)
                if len(xls_genesys.sheet_names) > 1:
                    aba_alvo = xls_genesys.sheet_names[1]
                    for aba in xls_genesys.sheet_names:
                        aba_low = aba.lower()
                        if "item" in aba_low or "detalhe" in aba_low or "pergunta" in aba_low:
                            aba_alvo = aba
                            break
                    self.df_fatos_itens = pd.read_excel(xls_genesys, sheet_name=aba_alvo, dtype=str)
        except Exception as e:
            self.erro_itens = str(e)

        if not os.path.exists(self.arquivo_exclusoes):
            self.erro_exclusoes = "ARQUIVO_NAO_EXISTE"
        else:
            try:
                xls = pd.ExcelFile(self.arquivo_exclusoes)
                df_encontrado = pd.DataFrame()

                for aba in xls.sheet_names:
                    temp_df = pd.read_excel(xls, sheet_name=aba, dtype=str)
                    if not temp_df.empty:
                        cols_upper = temp_df.columns.astype(str).str.upper()
                        if any("MONITORIA EXCLU" in c for c in cols_upper):
                            df_encontrado = temp_df
                            break

                if df_encontrado.empty:
                    for aba in xls.sheet_names:
                        temp_df = pd.read_excel(xls, sheet_name=aba, dtype=str)
                        if not temp_df.empty:
                            df_encontrado = temp_df
                            break

                self.df_exclusoes = df_encontrado
                if not self.df_exclusoes.empty:
                    self.df_exclusoes.columns = self.df_exclusoes.columns.astype(str).str.strip().str.upper()

            except Exception as e:
                self.erro_exclusoes = f"ERRO_LEITURA: {str(e)}"

    # ============================================================
    # 3. TRATAMENTO DE DATA DO EXCEL
    # ============================================================

    def parse_data_excel(self, val):
        if pd.isna(val):
            return "0000-00"
        val_str = str(val).strip()
        if val_str in ["nan", "NaT", "N/A", "None", ""]:
            return "0000-00"
        if len(val_str) >= 7 and val_str[4] == "-" and val_str[:4].isdigit():
            return val_str[:7]
        try:
            dt = pd.to_datetime(val, errors="coerce", dayfirst=True)
            if pd.notna(dt):
                return dt.strftime("%Y-%m")
        except Exception:
            pass
        return val_str.replace("/", "-").replace("\\", "-")[:7]

    # ============================================================
    # 4. TRANSFORMAÇÃO DA DIMENSÃO (EXCLUI TERCEIRO E CARGOS DE LIDERANÇA NA BASE HC)
    # ============================================================

    def transformar_dimensao(self):
        df = self.df_dim.copy()
        if df.empty:
            df = pd.DataFrame({
                "ID_AGENTE_CHAVE": ["N/A"],
                "LOGIN_LIMPO": ["N/A"],
                "MATRICULA": ["N/A"],
                "NOME": ["ERRO"],
                "CARGO": ["N/A"],
                "ALOCACAO": ["N/A"],
                "STATUS_DATA": ["Ativo"],
                "ATIVO_FLAG": ["SIM"],
                "ADMISSAO": ["N/A"],
                "DIAS_CASA_DIM": ["N/A"],
                "SUPERVISOR": ["Sem Gestor"],
                "COORDENADOR": ["Sem Gestor"],
                "GERENTE": ["Sem Gestor"],
                "SUPERINTENDENTE": ["Sem Gestor"],
                "MES_ANO_REF": ["N/A"],
                "ELEGIVEL_ESCOPO": ["SIM"]
            })
            self.df_dim = df.astype(str)
            return

        if "ANO/MES" not in df.columns:
            df["ANO/MES"] = "2026-01-01"
        df["MES_ANO_REF"] = df["ANO/MES"].apply(self.parse_data_excel)

        if "ADMISSAO" not in df.columns:
            df["ADMISSAO"] = pd.NaT
        df["ADMISSAO_DT"] = pd.to_datetime(df["ADMISSAO"], errors="coerce", dayfirst=True)

        if "LOGIN GENESYS" not in df.columns:
            df["LOGIN GENESYS"] = "DESCONHECIDO"

        login_excel = df["LOGIN GENESYS"].fillna("DESCONHECIDO").astype(str).str.strip().str.upper()
        df["LOGIN_LIMPO"] = login_excel
        df["ID_AGENTE_CHAVE"] = login_excel + "_" + df["MES_ANO_REF"].astype(str)

        if "CARGO" not in df.columns:
            df["CARGO"] = "N/A"
        df["CARGO"] = df["CARGO"].fillna("N/A").astype(str).str.strip()
        df.loc[df["CARGO"].str.lower().isin(["nan", "none", ""]), "CARGO"] = "N/A"

        col_aloc = next((c for c in df.columns if "ALOCAC" in self._limpar(c).upper()), None)
        if col_aloc:
            df["ALOCACAO"] = df[col_aloc].fillna("PRÓPRIO").astype(str).str.strip().str.upper()
        else:
            df["ALOCACAO"] = "PRÓPRIO"

        mapa_df = (
            pd.DataFrame({
                "LOGIN": login_excel,
                "ADMISSAO_DT": df["ADMISSAO_DT"],
                "ALOCACAO": df["ALOCACAO"]
            })
            .dropna(subset=["LOGIN"])
            .drop_duplicates(subset=["LOGIN"], keep="last")
        )
        self.mapa_admissao = dict(zip(mapa_df["LOGIN"], mapa_df["ADMISSAO_DT"]))
        self.mapa_alocacao_login = dict(zip(mapa_df["LOGIN"], mapa_df["ALOCACAO"]))

        # Identifica Status Ativo / Inativo no data.xlsx
        ativo_mask = pd.Series(True, index=df.index)
        col_status_hc = next(
            (c for c in df.columns if any(k in c for k in ["STATUS", "SITUAC", "SITUAÇ", "ATIVO", "CONDIC", "CONDIÇ"])),
            None
        )
        texto_status_bruto = pd.Series("ATIVO", index=df.index)

        if col_status_hc:
            texto_status_bruto = df[col_status_hc].fillna("ATIVO").astype(str).str.strip().str.upper()
            val_st = texto_status_bruto.apply(self._limpar)
            inativos_termos = ("demit", "deslig", "inativ", "afast", "licenc", "ferias", "suspens")
            ativo_mask = ~(
                val_st.apply(lambda s: any(t in s for t in inativos_termos) if s else False)
                | val_st.isin(["nao", "n", "0", "false", "inativo"])
            )

        col_demissao = next((c for c in df.columns if "DEMISS" in c or "DESLIG" in c), None)
        demitido_por_data = pd.Series(False, index=df.index)
        if col_demissao:
            dem_str = df[col_demissao].fillna("").astype(str).str.strip().str.lower()
            demitido_por_data = ~dem_str.isin(["", "nan", "nat", "none", "n/a", "-", "0"])
            ativo_mask = ativo_mask & (~demitido_por_data)

        def formatar_status_data(idx):
            eh_at = bool(ativo_mask.loc[idx])
            bruto = str(texto_status_bruto.loc[idx]).strip()
            bruto_limpo = self._limpar(bruto)

            if eh_at:
                if bruto_limpo in ("", "nan", "none", "na", "ativo", "sim", "s", "1", "true"):
                    return "Ativo"
                return f"Ativo ({bruto.title()})"
            else:
                if bool(demitido_por_data.loc[idx]) and bruto_limpo in ("", "nan", "none", "na", "ativo", "sim"):
                    return "Inativo (Desligado)"
                if bruto_limpo in ("", "nan", "none", "na", "inativo", "nao", "n", "0", "false"):
                    return "Inativo"
                return f"Inativo ({bruto.title()})"

        df["STATUS_DATA"] = [formatar_status_data(i) for i in df.index]
        df["ATIVO_FLAG"] = "NAO"
        df.loc[ativo_mask, "ATIVO_FLAG"] = "SIM"

        # 1) Exclui Terceiros da Base HC elegível
        nao_terceiro_mask = ~df["ALOCACAO"].apply(self._eh_alocacao_terceiro)

        # 2) Exclui Cargos de Liderança (SUPERVISOR ATENDIMENTO CLIENTE / COORDENADOR ATENDIMENTO CLIENTE)
        nao_cargo_excluido_mask = ~df["CARGO"].apply(self._eh_cargo_excluido)

        fim_mes_ref = pd.to_datetime(df["MES_ANO_REF"] + "-01", errors="coerce") + pd.offsets.MonthEnd(1)
        dias_casa_ref = (fim_mes_ref - df["ADMISSAO_DT"]).dt.days.fillna(-9999).astype("int64")
        df["DIAS_CASA_DIM"] = dias_casa_ref.astype(str).replace("-9999", "N/A")

        # Elegível de escopo: Não Terceiro + Não Supervisor/Coordenador + > 30 dias de casa
        elegivel_mask = (
            nao_terceiro_mask
            & nao_cargo_excluido_mask
            & (dias_casa_ref > self.dias_casa_minimo)
            & (login_excel != "DESCONHECIDO")
        )

        df["ELEGIVEL_ESCOPO"] = "NAO"
        df.loc[elegivel_mask, "ELEGIVEL_ESCOPO"] = "SIM"

        df["ADMISSAO"] = df["ADMISSAO_DT"].dt.strftime("%Y-%m-%d").fillna("N/A")

        colunas_manter = [
            "ID_AGENTE_CHAVE", "LOGIN_LIMPO", "MES_ANO_REF", "MATRICULA", "NOME", "CARGO",
            "ALOCACAO", "STATUS_DATA", "ATIVO_FLAG", "ADMISSAO", "DIAS_CASA_DIM",
            "SUPERVISOR", "COORDENADOR", "GERENTE", "SUPERINTENDENTE", "ELEGIVEL_ESCOPO"
        ]
        for col in colunas_manter:
            if col not in df.columns:
                if col in ["MATRICULA", "NOME", "CARGO", "ALOCACAO", "ADMISSAO", "LOGIN_LIMPO", "DIAS_CASA_DIM"]:
                    df[col] = "N/A"
                elif col == "STATUS_DATA":
                    df[col] = "Ativo"
                elif col in ["ELEGIVEL_ESCOPO", "ATIVO_FLAG"]:
                    df[col] = "SIM"
                else:
                    df[col] = "Sem Gestor"

        for col in ["SUPERVISOR", "COORDENADOR", "GERENTE", "SUPERINTENDENTE"]:
            df[col] = df[col].fillna("Sem Gestor").astype(str).str.strip()
            df.loc[df[col].str.lower().isin(["nan", "none", "n/a", ""]), col] = "Sem Gestor"

        # REGRA MARIELE: Se a coordenadora for MARIELE e estiver sem supervisor, repete o nome dela
        df["SUPERVISOR"] = df.apply(
            lambda r: self._ajustar_supervisor_mariele(r["COORDENADOR"], r["SUPERVISOR"]),
            axis=1
        )

        df["MATRICULA"] = df["MATRICULA"].fillna("N/A").astype(str)
        df["NOME"] = df["NOME"].fillna("N/A").astype(str)

        df = df[colunas_manter].drop_duplicates(subset=["ID_AGENTE_CHAVE"], keep="last").reset_index(drop=True)
        self.df_dim = df.astype(str)

    # ============================================================
    # 5. FILTROS DE ESCOPO DA FATO (MANTÉM AS 532 MONITORIAS)
    # ============================================================

    def _aplicar_filtros_escopo(self, df, col_status):
        col_form = next((c for c in df.columns if "formulario" in self._limpar(c)), None)

        motivos = pd.Series("", index=df.index)

        def marcar(mask, texto):
            motivos.loc[mask] = motivos.loc[mask] + texto + "; "

        if col_form:
            form_limpo = df[col_form].apply(self._limpar)
            marcar(
                ~form_limpo.isin(self.formularios_permitidos),
                "formulário fora do escopo"
            )
            df["formulario_padrao"] = df[col_form].fillna("N/A").astype(str).str.strip()
            df["canal_atendimento"] = "Outros"
            df.loc[form_limpo.str.contains("voz", na=False), "canal_atendimento"] = "Voz"
            df.loc[form_limpo.str.contains("texto", na=False), "canal_atendimento"] = "Texto"
        else:
            df["formulario_padrao"] = "N/A"
            df["canal_atendimento"] = "Geral"
            print("[FILTROS] Coluna nome_do_formulario não encontrada - filtro de formulário ignorado.")

        if col_status:
            marcar(
                df[col_status].astype(str).str.strip().str.upper() != self.status_permitido,
                f"status diferente de {self.status_permitido}"
            )
        else:
            print("[FILTROS] Coluna de status não encontrada - filtro de status ignorado.")

        if "GERENTE" in self.df_dim.columns and "ID_AGENTE_CHAVE" in self.df_dim.columns:
            mapa_gerente = self.df_dim.set_index("ID_AGENTE_CHAVE")["GERENTE"].to_dict()
            gerente_limpo = df["ID_AGENTE_CHAVE"].map(mapa_gerente).apply(self._limpar)
            marcar(~gerente_limpo.isin(self.gerentes_permitidos), "gerente fora do escopo")
        else:
            print("[FILTROS] Coluna GERENTE não encontrada na dimensão - filtro de gerente ignorado.")

        aloc_serie = pd.Series("", index=df.index)
        if "ALOCACAO" in self.df_dim.columns and "ID_AGENTE_CHAVE" in self.df_dim.columns:
            mapa_aloc_chave = self.df_dim.set_index("ID_AGENTE_CHAVE")["ALOCACAO"].to_dict()
            aloc_serie = df["ID_AGENTE_CHAVE"].map(mapa_aloc_chave).fillna("")

        if "LOGIN_AGENTE_LIMPO" in df.columns and self.mapa_alocacao_login:
            aloc_login = df["LOGIN_AGENTE_LIMPO"].map(self.mapa_alocacao_login).fillna("")
            aloc_serie = aloc_serie.where(aloc_serie != "", aloc_login)

        df["alocacao_agente"] = aloc_serie.replace("", "PRÓPRIO")

        dias = pd.to_numeric(df["DIAS_DE_CASA_NA_AVALIACAO"], errors="coerce")
        marcar(~(dias > self.dias_casa_minimo), f"até {self.dias_casa_minimo} dias de casa ou sem admissão")

        mascara = (motivos != "") & (df["validação"] == "monitoria válida")

        df["motivo_expurgo"] = ""
        df.loc[mascara, "motivo_expurgo"] = motivos[mascara].str.rstrip("; ")
        df.loc[mascara, "validação"] = "expurgo por regra de negócio"

        if "nota_numerica" in df.columns:
            df.loc[mascara, "nota_numerica"] = pd.NA
            df.loc[mascara, "nota_da_avaliacao"] = "N/A"

        self.diag_filtros = {
            "total_antes": int(len(df)),
            "expurgadas_regra_negocio": int(mascara.sum()),
            "exclusoes_qualidade": int((df["validação"] == "exclusão feita pela qualidade").sum()),
            "mantidas_validas": int((df["validação"] == "monitoria válida").sum()),
        }
        print("[FILTROS ESCOPO] Diagnóstico:", self.diag_filtros)

    # ============================================================
    # 6. TRANSFORMAÇÃO DA FATO (GENESYS)
    # ============================================================

    def transformar_fato(self):
        df = self.df_fato.copy()
        if df.empty:
            self.df_fato = pd.DataFrame(columns=[
                "ID_AGENTE_CHAVE", "nota_da_avaliacao", "DIAS_DE_CASA_NA_AVALIACAO", "validação"
            ])
            return

        col_aval = self._achar_col_id_avaliacao(df)
        col_int = next((c for c in df.columns if "id" in str(c).lower() and "intera" in str(c).lower()), None)
        col_status = next((c for c in df.columns if str(c).lower() in ["status", "state", "status_avaliacao"]), None)
        col_nota = next((c for c in df.columns if "nota_da_aval" in str(c).lower()), None)
        col_avaliador = next((c for c in df.columns if "nome_do_aval" in str(c).lower()), None)
        col_data = next((c for c in df.columns if "data_da_aval" in str(c).lower()), None)
        col_agente = next((c for c in df.columns if "id_do_agente" in str(c).lower()), None)

        def purificar_texto(texto):
            return "".join(c for c in str(texto).upper() if c.isalnum())

        df["validação"] = "monitoria válida"

        if self.erro_exclusoes == "ARQUIVO_NAO_EXISTE":
            df["validação"] = "ERRO: ARQUIVO DE EXCLUSAO NAO EXISTE NO PC"
        elif self.erro_exclusoes and "ERRO_LEITURA" in self.erro_exclusoes:
            if "Permission" in self.erro_exclusoes or "Acesso negado" in self.erro_exclusoes:
                df["validação"] = "ERRO: FECHE A PLANILHA DE EXCLUSAO, ESTA ABERTA!"
            else:
                df["validação"] = f"ERRO LEITURA: {self.erro_exclusoes[:20]}"
        elif self.df_exclusoes.empty:
            df["validação"] = "ERRO: PLANILHA LIDA, MAS TODAS AS ABAS ESTÃO VAZIAS"
        else:
            coluna_alvo = next((c for c in self.df_exclusoes.columns if "MONITORIA EXCLU" in str(c).upper()), None)
            if not coluna_alvo:
                cols_encontradas = "|".join(list(self.df_exclusoes.columns))[:50]
                df["validação"] = f"ERRO COLUNA. ACHEI: {cols_encontradas}"
            else:
                ids_excluidos = set(self.df_exclusoes[coluna_alvo].dropna().apply(purificar_texto).tolist())
                if len(ids_excluidos) == 0:
                    df["validação"] = "ERRO: A COLUNA EXISTE, MAS NÃO TEM NENHUM ID PREENCHIDO"
                else:
                    mascara_exclusao = pd.Series(False, index=df.index)
                    if col_int:
                        id_int_limpo = df[col_int].apply(purificar_texto)
                        mascara_exclusao = mascara_exclusao | id_int_limpo.isin(ids_excluidos)
                    if col_aval:
                        id_aval_limpo = df[col_aval].apply(purificar_texto)
                        mascara_exclusao = mascara_exclusao | id_aval_limpo.isin(ids_excluidos)

                    qtd_marcadas = mascara_exclusao.sum()
                    if qtd_marcadas == 0:
                        amostra_exc = list(ids_excluidos)[0][:10] if ids_excluidos else "VAZIO"
                        amostra_json = "VAZIO"
                        if col_int:
                            ids_validos = id_int_limpo[id_int_limpo.str.len() > 5]
                            if not ids_validos.empty:
                                amostra_json = ids_validos.iloc[0][:10]
                        df["validação"] = f"FALHA ZERO MATCH -> PBI: {amostra_json} | EXC: {amostra_exc}"
                    else:
                        df["validação"] = "monitoria válida"
                        df.loc[mascara_exclusao, "validação"] = "exclusão feita pela qualidade"

        subset_dedup = []
        if col_aval:
            subset_dedup.append(col_aval)
        elif col_int:
            subset_dedup.append(col_int)
        if col_status:
            subset_dedup.append(col_status)

        if subset_dedup:
            df = df.drop_duplicates(subset=subset_dedup, keep="last").reset_index(drop=True)

        df["ID_INTERACAO_PADRAO"] = df[col_int].fillna("N/A").astype(str) if col_int else "N/A"
        df["ID_AVALIACAO_NORM"] = df[col_aval].apply(self._norm_id) if col_aval else ""

        if col_nota:
            df["nota_numerica"] = pd.to_numeric(df[col_nota], errors="coerce").round(1)

            # --- REGRA DE +6 PONTOS (TABULAÇÃO NÃO CONFORME) ---
            df["status_bonus_tabulacao"] = "Não aplicável"

            if self.df_fatos_itens.empty:
                df["status_bonus_tabulacao"] = "ERRO: Tabela/Aba de Itens vazia ou não encontrada"
            elif not col_aval:
                df["status_bonus_tabulacao"] = "ERRO: Coluna ID de Avaliação não achada na fato"
            else:
                itens = self.df_fatos_itens
                col_id_it = self._achar_col_id_avaliacao(itens)
                col_perg, qtd_perg = self._escolher_coluna_por_conteudo(itens, "pergunta", "tabulacao")
                col_resp, qtd_resp = self._escolher_coluna_por_conteudo(itens, "resposta", "naoconforme")

                self.diag_bonus = {
                    "col_id_fato": col_aval,
                    "col_id_itens": col_id_it,
                    "col_pergunta": col_perg,
                    "linhas_com_tabulacao": qtd_perg,
                    "col_resposta": col_resp,
                    "linhas_com_naoconforme": qtd_resp,
                }

                if not (col_perg and col_resp and col_id_it):
                    cols_encontradas = "|".join(list(itens.columns))[:40]
                    df["status_bonus_tabulacao"] = f"ERRO Colunas. Achei: {cols_encontradas}"
                elif qtd_perg == 0:
                    df["status_bonus_tabulacao"] = f"ERRO: nenhuma pergunta com 'tabulação' na coluna [{col_perg}]"
                else:
                    df["status_bonus_tabulacao"] = "Avaliado - Sem bônus"

                    pergunta_limpa = itens[col_perg].apply(self._limpar)
                    resposta_limpa = itens[col_resp].apply(self._limpar)

                    condicao_bonus = (
                        pergunta_limpa.str.contains("tabulacao", na=False)
                        & resposta_limpa.str.contains("naoconforme", na=False)
                    )

                    ids_com_bonus = set(itens.loc[condicao_bonus, col_id_it].apply(self._norm_id)) - {""}
                    ids_fato_limpos = df["ID_AVALIACAO_NORM"]
                    mascara_bonus = ids_fato_limpos.isin(ids_com_bonus)

                    nota_original = df["nota_numerica"].copy()
                    mascara_nota_valida = mascara_bonus & (nota_original > 0.0)
                    mascara_nota_zero = mascara_bonus & (nota_original <= 0.0)

                    df.loc[mascara_nota_valida, "nota_numerica"] = (
                        (nota_original[mascara_nota_valida] + 6).clip(upper=100.0).round(1)
                    )

                    df.loc[mascara_nota_valida, "status_bonus_tabulacao"] = "Bônus +6 Aplicado"
                    df.loc[mascara_nota_zero, "status_bonus_tabulacao"] = "Bônus negado (Nota Zero)"

                    self.diag_bonus.update({
                        "itens_tabulacao_naoconforme": int(condicao_bonus.sum()),
                        "ids_unicos_com_bonus": len(ids_com_bonus),
                        "avaliacoes_da_fato_que_casaram": int(mascara_bonus.sum()),
                        "bonus_aplicados": int(mascara_nota_valida.sum()),
                        "bonus_negados_nota_zero": int(mascara_nota_zero.sum()),
                    })

                    if len(ids_com_bonus) > 0 and mascara_bonus.sum() == 0:
                        amostra_it = next(iter(ids_com_bonus))[:12]
                        amostra_ft = ids_fato_limpos[ids_fato_limpos != ""].iloc[0][:12] if (ids_fato_limpos != "").any() else "VAZIO"
                        df["status_bonus_tabulacao"] = f"ERRO: IDs nao casam -> ITENS: {amostra_it} | FATO: {amostra_ft}"

                print("[BONUS TABULACAO] Diagnóstico:", self.diag_bonus)

            df.loc[df["validação"] == "exclusão feita pela qualidade", "nota_numerica"] = pd.NA
            df["nota_da_avaliacao"] = df["nota_numerica"].apply(lambda x: str(x).replace(".", ",") if pd.notna(x) else "N/A")
        else:
            df["nota_numerica"] = pd.NA
            df["nota_da_avaliacao"] = "N/A"

        if col_avaliador:
            df["nome_do_avaliador"] = df[col_avaliador].fillna("N/A").astype(str).str.strip().str.upper()
            df["origem_monitoria"] = "Operação"
            df.loc[df["nome_do_avaliador"].isin(self.time_qualidade), "origem_monitoria"] = "Qualidade"
        else:
            df["nome_do_avaliador"] = "N/A"
            df["origem_monitoria"] = "Desconhecida"

        if col_agente:
            agente_chave_login = df[col_agente].fillna("DESCONHECIDO").astype(str).str.strip().str.upper()
        else:
            agente_chave_login = pd.Series("DESCONHECIDO", index=df.index)

        df["LOGIN_AGENTE_LIMPO"] = agente_chave_login
        df["ADMISSAO_DT"] = pd.to_datetime(agente_chave_login.map(self.mapa_admissao), errors="coerce")

        if col_data:
            dt_obj = pd.to_datetime(df[col_data], errors="coerce", dayfirst=True)
            df["data_da_avaliaçao"] = dt_obj.dt.strftime("%Y-%m-%d").fillna("N/A")
            df["semana_da_avaliacao"] = (dt_obj - pd.to_timedelta(dt_obj.dt.weekday, unit="D")).dt.strftime("%Y-%m-%d").fillna("N/A")
            mes_ano_fato = dt_obj.dt.strftime("%Y-%m").fillna("0000-00")
            df["DIAS_DE_CASA_NA_AVALIACAO"] = (
                (dt_obj - df["ADMISSAO_DT"]).dt.days.fillna(-9999).astype("int64").astype(str).replace("-9999", "N/A")
            )
        else:
            df["data_da_avaliaçao"] = "N/A"
            df["semana_da_avaliacao"] = "N/A"
            df["DIAS_DE_CASA_NA_AVALIACAO"] = "N/A"
            mes_ano_fato = pd.Series("0000-00", index=df.index)

        df["MES_ANO_FATO"] = mes_ano_fato.astype(str)
        df["ID_AGENTE_CHAVE"] = agente_chave_login + "_" + mes_ano_fato.astype(str)

        # --- FILTROS DE ESCOPO (formulário, status, gerente, >30d) ---
        self._aplicar_filtros_escopo(df, col_status)

        colunas_ignorar = {"nota_da_avaliacao", "nota_numerica", "ADMISSAO_DT"}
        if col_nota:
            colunas_ignorar.add(col_nota)

        for col in [c for c in df.columns if c not in colunas_ignorar]:
            df[col] = df[col].fillna("N/A")

        df.drop(columns=["ADMISSAO_DT"], inplace=True, errors="ignore")
        df = df.reset_index(drop=True)

        for col in df.columns:
            if col != "nota_numerica":
                df[col] = df[col].astype(str)

        df["nota_numerica"] = pd.to_numeric(df["nota_numerica"], errors="coerce")
        self.df_fato = df

    # ============================================================
    # 7. EXTRAÇÃO DE OFENSORES PELO GRUPO DE PERGUNTAS (CONDUÇÃO vs. NCG)
    # ============================================================

    def processar_itens_ofensores(self):
        if self.df_fatos_itens.empty or self.df_fato.empty:
            self.df_ofensores = pd.DataFrame()
            return

        itens = self.df_fatos_itens.copy()
        col_id_it = self._achar_col_id_avaliacao(itens)
        col_perg, _ = self._escolher_coluna_por_conteudo(itens, "pergunta", "tabulacao")
        col_resp, _ = self._escolher_coluna_por_conteudo(itens, "resposta", "naoconforme")
        col_grupo = self._achar_col_grupo_perguntas(itens)

        if not (col_id_it and col_perg and col_resp):
            self.df_ofensores = pd.DataFrame()
            return

        itens["ID_AVALIACAO_NORM"] = itens[col_id_it].apply(self._norm_id)
        itens["PERGUNTA_TEXTO"] = itens[col_perg].fillna("N/A").astype(str).str.strip()
        itens["RESPOSTA_TEXTO"] = itens[col_resp].fillna("N/A").astype(str).str.strip()
        itens["GRUPO_TEXTO"] = itens[col_grupo].fillna("Condução do Atendimento").astype(str).str.strip() if col_grupo else "Condução do Atendimento"

        grupos_unicos = itens["GRUPO_TEXTO"].dropna().unique().tolist()
        self.catalogo_grupos = []
        print("\n" + "=" * 75)
        print(f"[GRUPOS DE PERGUNTAS NA BASE1.XLSX] Coluna detectada: [{col_grupo}]")
        print("=" * 75)
        for g_nome in sorted(grupos_unicos):
            g_str = str(g_nome).strip()
            if not g_str or g_str.lower() in ("nan", "n/a", "none"):
                continue
            cat = self._classificar_por_grupo_de_perguntas(g_str)
            self.catalogo_grupos.append({
                "grupo_original": g_str,
                "classificacao_painel": cat
            })
            print(f"  [{cat:8s}] <- Grupo: \"{g_str}\"")
        print("=" * 75 + "\n")

        fato_validas = self.df_fato[self.df_fato["validação"] == "monitoria válida"][[
            "ID_AVALIACAO_NORM", "ID_AGENTE_CHAVE", "MES_ANO_FATO", "data_da_avaliaçao",
            "canal_atendimento", "origem_monitoria"
        ]].drop_duplicates(subset=["ID_AVALIACAO_NORM"])

        itens_merged = pd.merge(
            itens,
            fato_validas,
            on="ID_AVALIACAO_NORM",
            how="inner"
        )

        if itens_merged.empty:
            self.df_ofensores = pd.DataFrame()
            return

        r_limpa = itens_merged["RESPOSTA_TEXTO"].apply(self._limpar)
        mask_naoconforme = r_limpa.str.contains("naoconforme", na=False)

        df_of = itens_merged[mask_naoconforme].copy()

        if df_of.empty:
            self.df_ofensores = pd.DataFrame()
            return

        df_of["CATEGORIA_ITEM"] = df_of["GRUPO_TEXTO"].apply(self._classificar_por_grupo_de_perguntas)

        dim_cols = self.df_dim[[
            "ID_AGENTE_CHAVE", "NOME", "SUPERVISOR", "COORDENADOR", "GERENTE", "Quartil_Colaborador"
        ]].drop_duplicates(subset=["ID_AGENTE_CHAVE"]) if "Quartil_Colaborador" in self.df_dim.columns else self.df_dim[[
            "ID_AGENTE_CHAVE", "NOME", "SUPERVISOR", "COORDENADOR", "GERENTE"
        ]].drop_duplicates(subset=["ID_AGENTE_CHAVE"])

        df_of = pd.merge(df_of, dim_cols, on="ID_AGENTE_CHAVE", how="left")

        for c in ["NOME", "SUPERVISOR", "COORDENADOR", "GERENTE"]:
            df_of[c] = df_of[c].fillna("Sem Gestor").astype(str)

        # Garante a regra da coordenadora MARIELE também nos itens ofensores
        df_of["SUPERVISOR"] = df_of.apply(
            lambda r: self._ajustar_supervisor_mariele(r["COORDENADOR"], r["SUPERVISOR"]),
            axis=1
        )

        self.df_ofensores = df_of.reset_index(drop=True)
        self.diag_itens = {
            "coluna_grupo_usada": str(col_grupo),
            "total_ofensores_validos": int(len(df_of)),
            "total_conducao": int((df_of["CATEGORIA_ITEM"] == "Condução").sum()),
            "total_ncg": int((df_of["CATEGORIA_ITEM"] == "NCG").sum()),
        }
        print("[ITENS OFENSORES] Diagnóstico:", self.diag_itens)

    # ============================================================
    # 8. TRATAMENTO CSAT
    # ============================================================

    def transformar_csat(self):
        if self.erro_csat is not None:
            self.df_csat = pd.DataFrame({"STATUS_ERRO": [f"Falha ao ler arquivo: {self.erro_csat}"]})
            return

        df = self.df_csat.copy()

        if df.empty:
            self.df_csat = pd.DataFrame({"STATUS_ERRO": ["Arquivo encontrado, mas está vazio (sem dados)."]})
            return

        df.columns = df.columns.astype(str).str.strip().str.lower()

        colunas_numericas = [
            "nota_0", "nota_1", "nota_2", "nota_3", "nota_4", "nota_5",
            "nota_6", "nota_7", "nota_8", "nota_9", "nota_10",
            "detrator", "neutro", "promotor"
        ]

        for col in colunas_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        for col in df.columns:
            if col not in colunas_numericas:
                df[col] = df[col].fillna("N/A").astype(str)

        self.df_csat = df

    # ============================================================
    # 9. CLASSIFICAÇÃO POR FAIXAS DE META (85 e 75)
    # ============================================================

    def calcular_quartis(self):
        """
        Classifica os colaboradores avaliados nas 3 faixas solicitadas:
        - Acima da Meta (>= 85)
        - Acima de 75 (75 a 84,9)
        - Abaixo de 75 (< 75)
        """
        if "nota_numerica" not in self.df_fato.columns:
            self.df_dim["Quartil_Colaborador"] = "N/A"
            return

        df_validos = self.df_fato[self.df_fato["nota_numerica"].notna()].copy()

        if df_validos.empty:
            self.df_dim["Quartil_Colaborador"] = "N/A"
        else:
            medias = df_validos.groupby("ID_AGENTE_CHAVE", as_index=False)["nota_numerica"].mean()
            medias["Quartil_Colaborador"] = "Abaixo de 75 (< 75)"
            medias.loc[medias["nota_numerica"] >= 75.0, "Quartil_Colaborador"] = "Acima de 75 (75 a 84,9)"
            medias.loc[medias["nota_numerica"] >= 85.0, "Quartil_Colaborador"] = "Acima da Meta (>= 85)"

            medias = medias[["ID_AGENTE_CHAVE", "Quartil_Colaborador"]].drop_duplicates(subset=["ID_AGENTE_CHAVE"])
            self.df_dim = pd.merge(self.df_dim, medias, on="ID_AGENTE_CHAVE", how="left")
            self.df_dim["Quartil_Colaborador"] = self.df_dim["Quartil_Colaborador"].fillna("Sem Avaliação no Mês")

        self.df_dim = self.df_dim.fillna("N/A").astype(str)
        self.df_fato.drop(columns=["nota_numerica"], inplace=True, errors="ignore")

    # ============================================================
    # 10. EXECUÇÃO COMPLETA
    # ============================================================

    def executar_pipeline(self):
        self.extrair_dados()
        self.transformar_dimensao()
        self.transformar_fato()
        self.calcular_quartis()
        self.processar_itens_ofensores()
        self.transformar_csat()

        self.df_fato = self.df_fato.astype(str)
        self.df_dim = self.df_dim.astype(str)
        self.df_csat = self.df_csat.astype(str)

        return self.df_fato, self.df_dim, self.df_csat


# ==============================================================================
# GERADOR DE RELATÓRIO EM HTML - VISÃO EXECUTIVA PARA GERENTES & COORDENADORES
# ==============================================================================

class GeradorRelatorioHTML:

    def __init__(self, etl_instance):
        self.etl = etl_instance
        self.df_fato = etl_instance.df_fato.copy()
        self.df_dim = etl_instance.df_dim.copy()
        self.df_csat = etl_instance.df_csat.copy()
        self.df_ofensores = etl_instance.df_ofensores.copy()

        # Caminho onde o HTML será gerado
        self.pasta_saida = r"C:\Users\F40847326802\OneDrive - REDE AMÉRICAS\AMERICAS\Relatório de Qualidade (BOOK)"
        os.makedirs(self.pasta_saida, exist_ok=True)

        self.caminho_saida = os.path.join(
            self.pasta_saida,
            "Book_Qualidade.html"
        )

    def preparar_dados_json(self):
        df_f = self.df_fato.copy()
        df_d = self.df_dim.copy()

        if "nota_da_avaliacao" in df_f.columns:
            df_f["nota_float"] = pd.to_numeric(
                df_f["nota_da_avaliacao"].str.replace(",", ".", regex=False),
                errors="coerce"
            )
        else:
            df_f["nota_float"] = pd.NA

        df_merged = pd.merge(
            df_f,
            df_d,
            on="ID_AGENTE_CHAVE",
            how="left",
            suffixes=("", "_DIM")
        )

        def extrair_mes_ano(row):
            mes_dim = str(row.get("MES_ANO_REF", "N/A"))
            if mes_dim not in ["N/A", "nan", "0000-00", ""]:
                return mes_dim
            chave = str(row.get("ID_AGENTE_CHAVE", ""))
            if "_" in chave:
                possivel_mes = chave.split("_")[-1]
                if len(possivel_mes) == 7 and "-" in possivel_mes:
                    return possivel_mes
            data_av = str(row.get("data_da_avaliaçao", ""))
            if len(data_av) >= 7 and "-" in data_av:
                return data_av[:7]
            return "Sem Data"

        df_merged["MES_ANO_PAINEL"] = df_merged.apply(extrair_mes_ano, axis=1)

        campos_hierarquia = [
            "NOME", "MATRICULA", "CARGO", "ALOCACAO", "STATUS_DATA", "SUPERVISOR",
            "COORDENADOR", "GERENTE", "SUPERINTENDENTE", "Quartil_Colaborador", "LOGIN_LIMPO"
        ]
        for col in campos_hierarquia:
            if col not in df_merged.columns:
                df_merged[col] = "Não Cadastrado na Base"
            else:
                df_merged[col] = (
                    df_merged[col]
                    .replace({"nan": "Não Cadastrado na Base", "N/A": "Não Cadastrado na Base"})
                    .fillna("Não Cadastrado na Base")
                )

        # Aplica a regra da coordenadora MARIELE também após o merge com a Fato
        df_merged["SUPERVISOR"] = df_merged.apply(
            lambda r: self.etl._ajustar_supervisor_mariele(r.get("COORDENADOR", ""), r.get("SUPERVISOR", "")),
            axis=1
        )

        registros_avaliacoes = []
        for _, row in df_merged.iterrows():
            nota_val = row["nota_float"]
            nota_json = float(round(nota_val, 1)) if pd.notna(nota_val) else None
            login_un = str(row.get("LOGIN_AGENTE_LIMPO", row.get("LOGIN_LIMPO", row.get("ID_AGENTE_CHAVE", "N/A"))))
            aloc_str = str(row.get("alocacao_agente", row.get("ALOCACAO", "PRÓPRIO")))

            registros_avaliacoes.append({
                "id_interacao": str(row.get("ID_INTERACAO_PADRAO", "N/A")),
                "id_chave": str(row.get("ID_AGENTE_CHAVE", "N/A")),
                "login_unico": login_un,
                "mes_ano": str(row.get("MES_ANO_PAINEL", "Sem Data")),
                "data_avaliacao": str(row.get("data_da_avaliaçao", "N/A")),
                "semana": str(row.get("semana_da_avaliacao", "N/A")),
                "nome_agente": str(row.get("NOME", "Não Cadastrado")),
                "matricula": str(row.get("MATRICULA", "N/A")),
                "cargo": str(row.get("CARGO", "N/A")),
                "alocacao": aloc_str,
                "supervisor": str(row.get("SUPERVISOR", "Sem Gestor")),
                "coordenador": str(row.get("COORDENADOR", "Sem Gestor")),
                "gerente": str(row.get("GERENTE", "Sem Gestor")),
                "quartil": str(row.get("Quartil_Colaborador", "N/A")),
                "avaliador": str(row.get("nome_do_avaliador", "N/A")),
                "origem": str(row.get("origem_monitoria", "N/A")),
                "canal": str(row.get("canal_atendimento", "Geral")),
                "formulario": str(row.get("formulario_padrao", "N/A")),
                "validacao": str(row.get("validação", "N/A")),
                "motivo_expurgo": str(row.get("motivo_expurgo", "")),
                "status_bonus": str(row.get("status_bonus_tabulacao", "N/A")),
                "dias_casa": str(row.get("DIAS_DE_CASA_NA_AVALIACAO", "N/A")),
                "nota": nota_json
            })

        meses_validos = sorted([
            r["mes_ano"] for r in registros_avaliacoes
            if r["validacao"] == "monitoria válida" and r["mes_ano"] not in ("Sem Data", "0000-00", "N/A")
        ])
        mes_mais_recente = meses_validos[-1] if meses_validos else "TODOS"

        hc_dimensao = []
        if not df_d.empty and "GERENTE" in df_d.columns:
            df_d_ger = df_d[
                df_d["GERENTE"].apply(self.etl._limpar).isin(self.etl.gerentes_permitidos)
                & (df_d.get("ELEGIVEL_ESCOPO", "SIM") == "SIM")
            ].copy()
            for _, r_dim in df_d_ger.iterrows():
                coord_dim = str(r_dim.get("COORDENADOR", "Sem Gestor"))
                sup_dim = self.etl._ajustar_supervisor_mariele(coord_dim, r_dim.get("SUPERVISOR", "Sem Gestor"))
                hc_dimensao.append({
                    "id_chave": str(r_dim.get("ID_AGENTE_CHAVE", "N/A")),
                    "login_unico": str(r_dim.get("LOGIN_LIMPO", "N/A")),
                    "mes_ano": str(r_dim.get("MES_ANO_REF", "N/A")),
                    "nome": str(r_dim.get("NOME", "N/A")),
                    "matricula": str(r_dim.get("MATRICULA", "N/A")),
                    "cargo": str(r_dim.get("CARGO", "N/A")),
                    "status_data": str(r_dim.get("STATUS_DATA", "Ativo")),
                    "ativo_flag": str(r_dim.get("ATIVO_FLAG", "SIM")),
                    "alocacao": str(r_dim.get("ALOCACAO", "PRÓPRIO")),
                    "admissao": str(r_dim.get("ADMISSAO", "N/A")),
                    "dias_casa_dim": str(r_dim.get("DIAS_CASA_DIM", "N/A")),
                    "supervisor": sup_dim,
                    "coordenador": coord_dim,
                    "gerente": str(r_dim.get("GERENTE", "Sem Gestor")),
                    "quartil": str(r_dim.get("Quartil_Colaborador", "Sem Avaliação no Mês"))
                })

        registros_ofensores = []
        if not self.df_ofensores.empty:
            for _, r_of in self.df_ofensores.iterrows():
                coord_of = str(r_of.get("COORDENADOR", "Sem Gestor"))
                sup_of = self.etl._ajustar_supervisor_mariele(coord_of, r_of.get("SUPERVISOR", "Sem Gestor"))
                registros_ofensores.append({
                    "categoria": str(r_of.get("CATEGORIA_ITEM", "Condução")),
                    "pergunta": str(r_of.get("PERGUNTA_TEXTO", "N/A")),
                    "grupo": str(r_of.get("GRUPO_TEXTO", "")),
                    "resposta": str(r_of.get("RESPOSTA_TEXTO", "Não Conforme")),
                    "mes_ano": str(r_of.get("MES_ANO_FATO", "Sem Data")),
                    "data_avaliacao": str(r_of.get("data_da_avaliaçao", "N/A")),
                    "canal": str(r_of.get("canal_atendimento", "Geral")),
                    "origem": str(r_of.get("origem_monitoria", "N/A")),
                    "gerente": str(r_of.get("GERENTE", "Sem Gestor")),
                    "coordenador": coord_of,
                    "supervisor": sup_of,
                    "quartil": str(r_of.get("Quartil_Colaborador", "N/A")),
                    "nome_agente": str(r_of.get("NOME", "N/A"))
                })

        pacote_dados = {
            "gerado_em": datetime.now().strftime("%d/%m/%Y às %H:%M:%S"),
            "mes_mais_recente": mes_mais_recente,
            "diag_bonus": self.etl.diag_bonus,
            "diag_filtros": self.etl.diag_filtros,
            "diag_itens": self.etl.diag_itens,
            "catalogo_grupos": self.etl.catalogo_grupos,
            "erros": {
                "genesys": self.etl.erro_genesys,
                "excel_hc": self.etl.erro_excel,
                "exclusoes": self.etl.erro_exclusoes,
                "csat": self.etl.erro_csat,
                "itens": self.etl.erro_itens
            },
            "avaliacoes": registros_avaliacoes,
            "hc_dimensao": hc_dimensao,
            "ofensores": registros_ofensores
        }

        return json.dumps(pacote_dados, ensure_ascii=False)

    def gerar_html(self):
        dados_json_str = self.preparar_dados_json()

        html_template = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Book Executivo de Qualidade CX - Visão Gerencial & Coordenação</title>

    <!-- Fonte Oficial Rede Américas: Inter (Fallback: Arial) -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">

    <!-- Tailwind CSS configurado com a Paleta Oficial Rede Américas -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {{
            theme: {{
                extend: {{
                    fontFamily: {{
                        sans: ['Inter', 'Arial', 'sans-serif'],
                    }},
                    colors: {{
                        ra: {{
                            roxo: '#6D2077',       /* Roxo Púrpura Primário */
                            lavanda: '#D9B4D5',    /* Lavanda Primário */
                            verde: '#84BD00',      /* Verde Primário */
                            offwhite: '#F1F1EF',   /* Off White Primário */
                            grafite: '#343434',    /* Grafite Primário */
                            roxoEscuro: '#422140', /* Secundário Apoio */
                            tomRoxo: '#A76FA9',    /* Gradação 1 */
                            tomLavanda: '#CCA2C9', /* Gradação 2 */
                            tomVerde: '#A9B95C',   /* Gradação 3 */
                            verdeEscuro: '#4F7000',/* Secundário Apoio */
                            amarelo: '#FFD661',    /* Secundário Apoio */
                            coral: '#F0734F',      /* Secundário Apoio */
                            vermelho: '#A11430',   /* Secundário Apoio */
                            bege: '#E8E0D6'        /* Secundário Neutro */
                        }}
                    }}
                }}
            }}
        }}
    </script>

    <!-- Chart.js + Plugin DataLabels -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>

    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>

    <style>
        body {{
            font-family: 'Inter', Arial, sans-serif;
            background-color: #F1F1EF;
            color: #343434;
        }}
        .bg-gradiente-ra {{
            background: linear-gradient(90deg, #6D2077 0%, #D9B4D5 50%, #84BD00 100%);
        }}
        .card-coord {{
            transition: all 0.2s ease-in-out;
        }}
        .card-coord:hover {{
            transform: translateY(-3px);
            box-shadow: 0 12px 24px -6px rgba(109, 32, 119, 0.18);
        }}
        .custom-scrollbar::-webkit-scrollbar {{
            width: 6px;
            height: 6px;
        }}
        .custom-scrollbar::-webkit-scrollbar-track {{
            background: #F1F1EF;
        }}
        .custom-scrollbar::-webkit-scrollbar-thumb {{
            background: #CCA2C9;
            border-radius: 6px;
        }}
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {{
            background: #6D2077;
        }}
        .pbi-dropdown-menu {{
            max-height: 280px;
            box-shadow: 0 12px 28px -4px rgba(52, 52, 52, 0.22);
        }}
        @media print {{
            .no-print {{
                display: none !important;
            }}
        }}
    </style>
</head>
<body class="antialiased selection:bg-ra-roxo selection:text-white">

    <!-- FAIXA DE GRADIENTE OFICIAL REDE AMÉRICAS -->
    <div class="h-2.5 w-full bg-gradiente-ra"></div>

    <!-- CABEÇALHO EXECUTIVO COM BOTÃO DE IMAGEM DOS GRÁFICOS PARA O OUTLOOK -->
    <header class="bg-ra-roxo text-white shadow-lg sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div class="flex items-center gap-4">
                <div class="bg-white p-2 rounded-2xl shadow-sm flex items-center justify-center">
                    <svg width="34" height="34" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M12 10H30C33.3 10 36 12.7 36 16H10L12 10Z" fill="#6D2077"/>
                        <path d="M6 19H31C34.3 19 37 21.7 37 25H4L6 19Z" fill="#A76FA9"/>
                        <path d="M9 28H32C35.3 28 38 30.7 38 34H7L9 28Z" fill="#CCA2C9"/>
                        <path d="M36 36H58L56 42H42C38.7 42 36 39.3 36 36Z" fill="#D9B4D5"/>
                        <path d="M37 45H60L58 51H43C39.7 51 37 48.3 37 45Z" fill="#A9B95C"/>
                        <path d="M40 54H54L52 60H46C42.7 60 40 57.3 40 54Z" fill="#84BD00"/>
                    </svg>
                </div>
                <div>
                    <div class="flex items-center gap-2.5">
                        <span class="text-xl font-extrabold tracking-tight text-white">Rede Américas</span>
                        <span class="text-[11px] px-2.5 py-0.5 rounded-full bg-ra-verde text-ra-roxoEscuro font-extrabold uppercase tracking-wider">
                            Book Gerencial CX
                        </span>
                    </div>
                    <p class="text-xs text-ra-lavanda mt-0.5">
                        Visão por Coordenação • Faixas de Meta (85 e 75 pts) • Condução vs. NCG
                    </p>
                </div>
            </div>

            <div class="flex flex-wrap items-center gap-2.5">
                <div class="text-xs bg-ra-roxoEscuro/90 px-3.5 py-2 rounded-xl border border-ra-lavanda/30 flex items-center gap-2">
                    <i data-lucide="clock" class="w-4 h-4 text-ra-verde"></i>
                    <span class="text-ra-lavanda">Atualizado em: <strong id="lbl-data-geracao" class="text-white"></strong></span>
                </div>

                <!-- BOTÃO PRINCIPAL: GERAR IMAGEM ORGANIZADA DOS GRÁFICOS E ABRIR OUTLOOK -->
                <button onclick="prepararImagemGraficosEEnviarOutlook()" class="no-print bg-ra-verde hover:bg-ra-verdeEscuro text-ra-roxoEscuro hover:text-white text-xs font-extrabold px-4 py-2.5 rounded-xl shadow-md flex items-center gap-2 transition">
                    <i data-lucide="image" class="w-4 h-4"></i>
                    Gerar Imagem dos Gráficos &amp; Abrir Outlook
                </button>

                <button onclick="window.print()" class="no-print bg-white/10 hover:bg-white/20 text-white text-xs font-semibold px-3.5 py-2 rounded-xl border border-white/20 flex items-center gap-1.5 transition">
                    <i data-lucide="printer" class="w-4 h-4 text-ra-verde"></i>
                    PDF
                </button>
            </div>
        </div>
    </header>

    <!-- MODAL DE IMAGEM ORGANIZADA COM GRÁFICOS PARA O OUTLOOK -->
    <div id="modal-email-outlook" class="hidden fixed inset-0 z-50 bg-black/70 backdrop-blur-xs flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl border border-ra-bege max-w-6xl w-full max-h-[92vh] flex flex-col overflow-hidden">
            <div class="bg-ra-roxo text-white px-6 py-4 flex items-center justify-between">
                <div class="flex items-center gap-3">
                    <div class="bg-ra-verde text-ra-roxoEscuro p-2 rounded-xl">
                        <i data-lucide="image-plus" class="w-5 h-5"></i>
                    </div>
                    <div>
                        <h3 class="text-base font-extrabold">Painel em Imagem de Alta Resolução Pronto para o E-mail</h3>
                        <p class="text-[11px] text-ra-lavanda">
                            A imagem organizada com os KPIs, Cards de Coordenação e os 7 Gráficos já foi copiada para sua Área de Transferência!
                        </p>
                    </div>
                </div>
                <button onclick="fecharModalEmail()" class="text-white/80 hover:text-white text-sm font-bold px-3 py-1 rounded-lg bg-white/10">Fechar [X]</button>
            </div>

            <!-- Barra de Ações da Imagem -->
            <div class="p-4 bg-ra-offwhite border-b border-ra-bege flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
                <div class="text-xs text-ra-grafite space-y-1">
                    <p>
                        <strong class="text-ra-roxo">Como colar no Outlook:</strong> O seu Outlook foi aberto. Clique no corpo do e-mail e aperte
                        <kbd class="px-2 py-0.5 bg-white border border-ra-roxo/40 rounded font-mono font-extrabold text-ra-roxo">Ctrl + V</kbd>
                        para colar o painel gráfico completo como imagem!
                    </p>
                    <p class="text-[11px] text-ra-grafite/70">
                        Você também pode copiar qualquer gráfico individualmente nos botões abaixo ou baixar o arquivo <code>.png</code>.
                    </p>
                </div>
                <div class="flex flex-wrap items-center gap-2 shrink-0">
                    <button onclick="copiarImagemCompletaParaClipboard()" id="btn-copiar-img-completa" class="bg-ra-roxo hover:bg-ra-roxoEscuro text-white text-xs font-extrabold px-4 py-2.5 rounded-xl shadow-sm flex items-center gap-1.5 transition">
                        <i data-lucide="copy" class="w-4 h-4 text-ra-verde"></i>
                        <span>Copiar Imagem Completa (Ctrl+C)</span>
                    </button>
                    <button onclick="baixarImagemCompletaPNG()" class="bg-white hover:bg-ra-lavanda/30 text-ra-roxo border border-ra-roxo/30 text-xs font-bold px-3.5 py-2.5 rounded-xl shadow-xs flex items-center gap-1.5 transition">
                        <i data-lucide="download" class="w-4 h-4"></i>
                        Baixar PNG
                    </button>
                    <button onclick="acionarMailtoOutlook()" class="bg-ra-verde hover:bg-ra-verdeEscuro text-ra-roxoEscuro hover:text-white text-xs font-extrabold px-4 py-2.5 rounded-xl shadow-sm flex items-center gap-1.5 transition">
                        <i data-lucide="mail" class="w-4 h-4"></i>
                        Abrir Outlook Novamente
                    </button>
                </div>
            </div>

            <!-- Botões rápidos para copiar um gráfico individual -->
            <div class="px-5 py-2.5 bg-white border-b border-ra-bege flex flex-wrap items-center gap-2 text-[11px]">
                <span class="font-bold text-ra-roxo uppercase mr-1">Copiar Gráfico Individual:</span>
                <button onclick="copiarGraficoIndividual('chartCoordenadores', 'Performance por Coordenação', this)" class="px-2.5 py-1 rounded-lg bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-grafite font-semibold border border-ra-bege transition">
                    1. Nota por Coordenação
                </button>
                <button onclick="copiarGraficoIndividual('chartQuartisCoord', 'Faixas de Meta por Coordenação', this)" class="px-2.5 py-1 rounded-lg bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-grafite font-semibold border border-ra-bege transition">
                    2. Faixas de Meta (85/75)
                </button>
                <button onclick="copiarGraficoIndividual('chartComparativoOrigem', 'Volume x Nota (Qualidade vs. Operação)', this)" class="px-2.5 py-1 rounded-lg bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-grafite font-semibold border border-ra-bege transition">
                    3. Qualidade vs. Operação
                </button>
                <button onclick="copiarGraficoIndividual('chartEvolucao', 'Evolução Semanal (Volume vs. Nota)', this)" class="px-2.5 py-1 rounded-lg bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-grafite font-semibold border border-ra-bege transition">
                    4. Evolução Semanal
                </button>
                <button onclick="copiarGraficoIndividual('chartOfensoresConducao', 'Top 10 Ofensores Condução', this)" class="px-2.5 py-1 rounded-lg bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-grafite font-semibold border border-ra-bege transition">
                    5. Ofensores Condução
                </button>
                <button onclick="copiarGraficoIndividual('chartOfensoresNCG', 'Incidências de NCG', this)" class="px-2.5 py-1 rounded-lg bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-grafite font-semibold border border-ra-bege transition">
                    6. Incidências NCG
                </button>
                <button onclick="copiarGraficoIndividual('chartSupervisores', 'Desempenho por Supervisor', this)" class="px-2.5 py-1 rounded-lg bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-grafite font-semibold border border-ra-bege transition">
                    7. Supervisores
                </button>
            </div>

            <!-- Área de Visualização da Imagem Gerada -->
            <div class="p-6 overflow-y-auto custom-scrollbar bg-ra-offwhite/60 flex-1 flex flex-col items-center">
                <img id="img-preview-painel-executivo" alt="Painel Executivo de Gráficos Rede Américas" class="max-w-full rounded-xl shadow-lg border border-ra-bege bg-white" />
                <canvas id="canvas-exportacao-email" class="hidden"></canvas>
            </div>
        </div>
    </div>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">

        <!-- ALERTAS DE LEITURA DE ARQUIVO (SE HOUVER) -->
        <div id="box-alertas" class="hidden bg-white border-l-4 border-ra-vermelho p-4 rounded-2xl shadow-sm">
            <div class="flex items-start gap-3">
                <i data-lucide="alert-triangle" class="w-5 h-5 text-ra-vermelho shrink-0 mt-0.5"></i>
                <div class="text-sm text-ra-grafite space-y-1" id="lista-alertas"></div>
            </div>
        </div>

        <!-- BARRA DE SEGMENTAÇÃO MULTI-SELEÇÃO (ESTILO POWER BI) -->
        <section class="bg-white p-5 rounded-2xl shadow-sm border border-ra-bege space-y-4">
            <div class="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 border-b border-ra-bege pb-3.5">
                <div>
                    <span class="text-xs font-extrabold uppercase tracking-wider text-ra-roxo flex items-center gap-2">
                        <i data-lucide="sliders-horizontal" class="w-4 h-4 text-ra-verde"></i>
                        Segmentação de Dados (Filtros)
                    </span>
                    <p class="text-xs text-ra-grafite/70 mt-0.5">
                        Clique em qualquer lista suspensa para marcar <strong>um, dois ou mais itens simultaneamente</strong> (incluindo Coordenadores e Dias)
                    </p>
                </div>
                <div class="flex flex-wrap items-center gap-2">
                    <div id="botoes-gerentes" class="flex flex-wrap gap-1.5"></div>
                    <button onclick="resetarFiltros()" class="bg-ra-offwhite hover:bg-ra-lavanda/40 text-ra-roxo font-bold text-xs py-2 px-3.5 rounded-xl border border-ra-bege flex items-center gap-1.5 transition">
                        <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i>
                        Limpar Todas as Seleções
                    </button>
                </div>
            </div>

            <!-- Grid de Dropdowns Multi-Seleção -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
                <div id="slicer-mes" class="relative"></div>
                <div id="slicer-data" class="relative"></div>
                <div id="slicer-gerente" class="relative"></div>
                <div id="slicer-coordenador" class="relative"></div>
                <div id="slicer-supervisor" class="relative"></div>
                <div id="slicer-canal" class="relative"></div>
                <div id="slicer-origem" class="relative"></div>
                <div class="grid grid-cols-2 gap-2">
                    <div id="slicer-quartil" class="relative"></div>
                    <div>
                        <label class="block text-[11px] font-bold text-ra-grafite mb-1 uppercase tracking-wide">Escopo</label>
                        <select id="filtro-validacao" onchange="aoMudarStatusEscopo()" class="w-full text-xs border border-ra-bege rounded-xl bg-ra-offwhite px-2.5 py-2 font-semibold text-ra-grafite focus:ring-2 focus:ring-ra-roxo focus:outline-none">
                            <option value="VALIDAS" selected>Válidas</option>
                            <option value="EXPURGADAS">Expurgos</option>
                            <option value="TODOS">Todas</option>
                        </select>
                    </div>
                </div>
            </div>
        </section>

        <!-- KPIS EXECUTIVOS DA GERÊNCIA -->
        <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            <div class="bg-white p-5 rounded-2xl shadow-sm border-t-4 border-ra-roxo flex items-center justify-between">
                <div>
                    <p class="text-[11px] font-bold uppercase tracking-wider text-ra-grafite/70">Nota Média Geral</p>
                    <h3 id="kpi-nota-media" class="text-3xl font-extrabold text-ra-roxo mt-1">0,0</h3>
                    <p id="kpi-sub-calibracao" class="text-[11px] text-ra-grafite/80 mt-1 font-medium">
                        Qualidade: - | Operação: -
                    </p>
                </div>
                <div class="bg-ra-lavanda/30 text-ra-roxo p-3 rounded-2xl">
                    <i data-lucide="award" class="w-7 h-7"></i>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl shadow-sm border-t-4 border-ra-verde flex items-center justify-between">
                <div>
                    <p class="text-[11px] font-bold uppercase tracking-wider text-ra-grafite/70">Monitorias Válidas</p>
                    <h3 id="kpi-total-validas" class="text-3xl font-extrabold text-ra-grafite mt-1">0</h3>
                    <p id="kpi-sub-canal" class="text-[11px] text-ra-grafite/80 mt-1 font-medium">
                        Voz: 0 | Texto: 0
                    </p>
                </div>
                <div class="bg-ra-verde/20 text-ra-verdeEscuro p-3 rounded-2xl">
                    <i data-lucide="check-circle-2" class="w-7 h-7"></i>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl shadow-sm border-t-4 border-ra-tomRoxo flex items-center justify-between">
                <div>
                    <p class="text-[11px] font-bold uppercase tracking-wider text-ra-grafite/70">HC Avaliado (&gt;30d)</p>
                    <h3 id="kpi-agentes-avaliados" class="text-3xl font-extrabold text-ra-roxoEscuro mt-1">0</h3>
                    <p id="kpi-sub-cobertura" class="text-[11px] text-ra-grafite/80 mt-1 font-medium">
                        Pendentes no mês: 0
                    </p>
                </div>
                <div class="bg-ra-lavanda/40 text-ra-roxo p-3 rounded-2xl">
                    <i data-lucide="users" class="w-7 h-7"></i>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl shadow-sm border-t-4 border-ra-verdeEscuro flex items-center justify-between">
                <div>
                    <p class="text-[11px] font-bold uppercase tracking-wider text-ra-grafite/70">Acima da Meta (&ge; 85)</p>
                    <h3 id="kpi-perc-meta" class="text-3xl font-extrabold text-ra-verdeEscuro mt-1">0%</h3>
                    <p id="kpi-sub-faixas" class="text-[10px] text-ra-grafite/80 mt-1 font-semibold">
                        &ge;85: 0 | 75-84: 0 | &lt;75: 0
                    </p>
                </div>
                <div class="bg-ra-tomVerde/30 text-ra-verdeEscuro p-3 rounded-2xl">
                    <i data-lucide="trending-up" class="w-7 h-7"></i>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl shadow-sm border-t-4 border-ra-vermelho flex items-center justify-between">
                <div>
                    <p class="text-[11px] font-bold uppercase tracking-wider text-ra-grafite/70">Itens Não Conformes</p>
                    <h3 id="kpi-total-ofensores" class="text-3xl font-extrabold text-ra-roxo mt-1">0</h3>
                    <p id="kpi-sub-ofensores" class="text-[11px] text-ra-vermelho mt-1 font-bold">
                        Condução: 0 | NCG: 0
                    </p>
                </div>
                <div class="bg-ra-vermelho/10 text-ra-vermelho p-3 rounded-2xl">
                    <i data-lucide="shield-alert" class="w-7 h-7"></i>
                </div>
            </div>
        </section>

        <!-- SCORECARDS VISUAIS DOS COORDENADORES (COM FAIXAS DE META 85/75 E OFENSORES) -->
        <section class="space-y-3">
            <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                <div>
                    <h2 class="text-lg font-extrabold text-ra-roxo flex items-center gap-2">
                        <i data-lucide="layout-grid" class="w-5 h-5 text-ra-verde"></i>
                        Scorecards Executivos por Coordenação
                    </h2>
                    <p class="text-xs text-ra-grafite/75">
                        Visão completa para os Gerentes • <strong>Clique em um ou mais cards</strong> para selecionar ou desmarcar Coordenadores no filtro
                    </p>
                </div>
                <span id="lbl-filtro-coord-ativo" class="hidden text-xs bg-ra-roxo text-white px-3 py-1.5 rounded-xl font-semibold"></span>
            </div>

            <div id="grid-cards-coordenadores" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
            </div>
        </section>

        <!-- TABELA EXPANSÍVEL: OPERADORES ELEGÍVEIS (>30d, Próprios, Exceto Supervisores/Coordenadores) SEM MONITORIA -->
        <section class="bg-white rounded-2xl shadow-sm border-2 border-ra-roxo/25 overflow-hidden">
            <div class="p-5 bg-ra-lavanda/15 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                <div class="flex items-start gap-3">
                    <div class="bg-ra-roxo text-white p-2.5 rounded-xl shrink-0 mt-0.5">
                        <i data-lucide="user-x" class="w-5 h-5"></i>
                    </div>
                    <div>
                        <div class="flex flex-wrap items-center gap-2">
                            <h3 class="text-base font-extrabold text-ra-roxo">
                                Controle de Cobertura — Operadores Elegíveis Sem Monitoria
                            </h3>
                            <span id="badge-qtd-sem-monitoria" class="text-xs font-extrabold px-3 py-0.5 rounded-full bg-ra-roxo text-white">
                                0 pendentes
                            </span>
                        </div>
                        <p class="text-xs text-ra-grafite/80 mt-0.5">
                            Colaboradores na base (<code class="font-semibold">data.xlsx</code>), com mais de 30 dias de casa (exceto Terceiros, Supervisores e Coordenadores) que estão com zero monitorias no período (<strong id="lbl-mes-ref-pendentes">-</strong>)
                        </p>
                    </div>
                </div>

                <div class="flex flex-wrap items-center gap-2 shrink-0">
                    <select id="filtro-status-sem-monitoria" onchange="aplicarFiltros()" class="text-xs border border-ra-roxo/30 rounded-xl bg-white px-3 py-2 font-bold text-ra-roxo focus:outline-none shadow-2xs">
                        <option value="TODOS" selected>Status no Data: Todos (Ativos + Inativos)</option>
                        <option value="SIM">Apenas Ativos no Data</option>
                        <option value="NAO">Apenas Inativos / Afastados</option>
                    </select>
                    <button onclick="exportarSemMonitoriaCSV()" class="bg-ra-verde hover:bg-ra-verdeEscuro text-ra-roxoEscuro hover:text-white text-xs font-extrabold px-3.5 py-2 rounded-xl shadow-sm flex items-center gap-1.5 transition">
                        <i data-lucide="download" class="w-4 h-4"></i>
                        Exportar Pendentes (CSV)
                    </button>
                    <button onclick="alternarTabelaSemMonitoria()" id="btn-toggle-sem-monitoria" class="bg-ra-roxo hover:bg-ra-roxoEscuro text-white text-xs font-bold px-4 py-2 rounded-xl shadow-sm flex items-center gap-1.5 transition">
                        <span id="txt-toggle-sem-monitoria">Expandir Lista Nominal</span>
                        <i data-lucide="chevron-down" id="icon-toggle-sem-monitoria" class="w-4 h-4"></i>
                    </button>
                </div>
            </div>

            <div id="box-tabela-sem-monitoria" class="hidden border-t border-ra-bege">
                <div id="resumo-pendentes-coord" class="p-4 bg-ra-offwhite/70 border-b border-ra-bege flex flex-wrap gap-2">
                </div>

                <div class="overflow-x-auto custom-scrollbar max-h-96">
                    <table class="w-full text-left border-collapse">
                        <thead class="bg-ra-roxoEscuro text-white text-[11px] uppercase sticky top-0 z-10">
                            <tr>
                                <th class="py-3 px-4 font-semibold">Coordenação</th>
                                <th class="py-3 px-4 font-semibold">Supervisor</th>
                                <th class="py-3 px-4 font-semibold">Operador(a) Sem Monitoria</th>
                                <th class="py-3 px-4 font-semibold">Cargo</th>
                                <th class="py-3 px-4 font-semibold text-center">Ativo no Data?</th>
                                <th class="py-3 px-4 font-semibold text-center">Alocação</th>
                                <th class="py-3 px-4 font-semibold text-center">Matrícula / Login</th>
                                <th class="py-3 px-4 font-semibold text-center">Data Admissão</th>
                                <th class="py-3 px-4 font-semibold text-center">Dias de Casa</th>
                                <th class="py-3 px-4 font-semibold text-center">Mês Ref. Base</th>
                            </tr>
                        </thead>
                        <tbody id="tbody-sem-monitoria" class="divide-y divide-ra-bege text-xs text-ra-grafite bg-white">
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- GRÁFICOS EXECUTIVOS DE COORDENAÇÃO -->
        <section class="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-ra-bege">
                <div class="mb-4">
                    <h3 class="text-base font-bold text-ra-roxo flex items-center gap-2">
                        <i data-lucide="bar-chart-3" class="w-5 h-5 text-ra-verde"></i>
                        Nota Média por Coordenação (Geral vs. Qualidade vs. Operação)
                    </h3>
                    <p class="text-xs text-ra-grafite/70">
                        Permite avaliar a performance consolidada e a calibração de nota entre Qualidade e Operação
                    </p>
                </div>
                <div class="h-80 relative">
                    <canvas id="chartCoordenadores"></canvas>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl shadow-sm border border-ra-bege">
                <div class="mb-4">
                    <h3 class="text-base font-bold text-ra-roxo flex items-center gap-2">
                        <i data-lucide="percent" class="w-5 h-5 text-ra-verde"></i>
                        Faixas de Meta por Coordenação (% da Equipe Avaliada)
                    </h3>
                    <p class="text-xs text-ra-grafite/70">
                        Proporção de colaboradores <strong>Acima da Meta (&ge; 85)</strong>, <strong>Acima de 75 (75 a 84,9)</strong> e <strong>Abaixo de 75 (&lt; 75)</strong>
                    </p>
                </div>
                <div class="h-80 relative">
                    <canvas id="chartQuartisCoord"></canvas>
                </div>
            </div>
        </section>

        <!-- LINHA EXECUTIVA: COMPARATIVO QUALIDADE VS OPERAÇÃO & EVOLUÇÃO -->
        <section class="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-ra-bege">
                <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 mb-4">
                    <div>
                        <h3 class="text-base font-bold text-ra-roxo flex items-center gap-2">
                            <i data-lucide="scale" class="w-5 h-5 text-ra-verde"></i>
                            Comparativo Volume x Nota: Qualidade vs. Operação
                        </h3>
                        <p class="text-xs text-ra-grafite/70">
                            Representatividade de avaliações realizadas e calibração de nota média frente à Meta (85,0)
                        </p>
                    </div>
                    <span id="badge-gap-calibracao" class="text-[11px] font-extrabold px-3 py-1 rounded-xl bg-ra-lavanda/35 text-ra-roxo self-start sm:self-center shrink-0">
                        Gap: 0,0 pts
                    </span>
                </div>
                <div class="h-80 relative">
                    <canvas id="chartComparativoOrigem"></canvas>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl shadow-sm border border-ra-bege">
                <div class="mb-4">
                    <h3 class="text-base font-bold text-ra-roxo flex items-center gap-2">
                        <i data-lucide="activity" class="w-5 h-5 text-ra-verde"></i>
                        Evolução Semanal (Volume de Monitorias vs. Nota Média)
                    </h3>
                    <p class="text-xs text-ra-grafite/70">
                        Histórico semanal de tendência da operação com linha de referência da Meta (85,0 pts)
                    </p>
                </div>
                <div class="h-80 relative">
                    <canvas id="chartEvolucao"></canvas>
                </div>
            </div>
        </section>

        <!-- DIAGNÓSTICO DE ITENS OFENSORES PELO GRUPO DE PERGUNTAS (CONDUÇÃO VS. NCG) -->
        <section class="space-y-4">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-ra-bege">
                <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-3 border-b border-ra-bege pb-4 mb-5">
                    <div>
                        <h2 class="text-lg font-extrabold text-ra-roxo flex items-center gap-2">
                            <i data-lucide="clipboard-x" class="w-5 h-5 text-ra-verde"></i>
                            Diagnóstico Geral de Ofensores: Condução do Atendimento vs. NCG
                        </h2>
                        <p class="text-xs text-ra-grafite/75 mt-0.5">
                            Itens sinalizados como <strong>"Não Conforme"</strong> separados entre <strong>Condução do Atendimento</strong> (tiram pontos) e <strong>NCG</strong> (zeram a monitoria)
                        </p>
                    </div>
                    <div class="flex flex-wrap items-center gap-2">
                        <span class="px-3 py-1.5 rounded-xl bg-ra-lavanda/35 text-ra-roxo text-xs font-extrabold" id="badge-resumo-conducao">
                            Condução (NC): 0 desvios
                        </span>
                        <span class="px-3 py-1.5 rounded-xl bg-ra-vermelho/15 text-ra-vermelho text-xs font-extrabold" id="badge-resumo-ncg">
                            NCG (Grave): 0 ocorrências
                        </span>
                    </div>
                </div>

                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    <div class="bg-ra-offwhite/60 p-4 rounded-2xl border border-ra-bege">
                        <h3 class="text-sm font-bold text-ra-roxo flex items-center gap-2 mb-1">
                            <span class="w-3 h-3 rounded-full bg-ra-roxo inline-block"></span>
                            Top 10 Ofensores — Condução do Atendimento (NC)
                        </h3>
                        <p class="text-[11px] text-ra-grafite/70 mb-3">
                            Itens que tiram pontos da monitoria com maior volume de "Não Conforme"
                        </p>
                        <div class="h-80 relative">
                            <canvas id="chartOfensoresConducao"></canvas>
                        </div>
                    </div>

                    <div class="bg-ra-vermelho/5 p-4 rounded-2xl border border-ra-vermelho/20">
                        <h3 class="text-sm font-bold text-ra-vermelho flex items-center gap-2 mb-1">
                            <span class="w-3 h-3 rounded-full bg-ra-vermelho inline-block"></span>
                            Incidências de NCG — Não Conformidade Grave (Zeramento)
                        </h3>
                        <p class="text-[11px] text-ra-grafite/70 mb-3">
                            Itens que zeram a monitoria sinalizados como "Não Conforme"
                        </p>
                        <div class="h-80 relative">
                            <canvas id="chartOfensoresNCG"></canvas>
                        </div>
                    </div>
                </div>

                <div class="mt-6 border-t border-ra-bege pt-5">
                    <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-3">
                        <div>
                            <h4 class="text-sm font-bold text-ra-roxo">
                                Detalhamento de Itens "Não Conforme" por Coordenação
                            </h4>
                            <p class="text-xs text-ra-grafite/70">
                                Veja exatamente qual Coordenador concentra cada desvio de Condução do Atendimento ou NCG
                            </p>
                        </div>
                        <div class="flex items-center gap-2">
                            <label class="text-xs font-semibold text-ra-grafite">Filtrar Grupo:</label>
                            <select id="filtro-tipo-ofensor" onchange="atualizarSecaoOfensores()" class="text-xs border border-ra-bege rounded-xl bg-ra-offwhite px-3 py-1.5 font-bold text-ra-roxo focus:outline-none">
                                <option value="TODOS">Todos os Grupos (Condução + NCG)</option>
                                <option value="Condução">Apenas Condução do Atendimento (NC)</option>
                                <option value="NCG">Apenas Não Conformidade Grave (NCG)</option>
                            </select>
                        </div>
                    </div>

                    <div class="overflow-x-auto custom-scrollbar max-h-80 rounded-xl border border-ra-bege">
                        <table class="w-full text-left border-collapse">
                            <thead class="bg-ra-roxoEscuro text-white text-[11px] uppercase sticky top-0">
                                <tr>
                                    <th class="py-3 px-4 font-semibold">Classificação</th>
                                    <th class="py-3 px-4 font-semibold">Grupo de Perguntas</th>
                                    <th class="py-3 px-4 font-semibold">Coordenação</th>
                                    <th class="py-3 px-4 font-semibold">Pergunta / Item Avaliado</th>
                                    <th class="py-3 px-4 font-semibold text-center">Qtd Não Conforme</th>
                                    <th class="py-3 px-4 font-semibold text-center">% nas Monitorias da Coord.</th>
                                    <th class="py-3 px-4 font-semibold text-center">Voz / Texto</th>
                                </tr>
                            </thead>
                            <tbody id="tbody-ofensores-coord" class="divide-y divide-ra-bege text-xs text-ra-grafite bg-white">
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

        <!-- MATRIZ EXECUTIVA CONSOLIDADA POR COORDENAÇÃO -->
        <section class="bg-white rounded-2xl shadow-sm border border-ra-bege overflow-hidden">
            <div class="p-5 border-b border-ra-bege flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 bg-ra-offwhite/60">
                <div>
                    <h3 class="text-base font-bold text-ra-roxo flex items-center gap-2">
                        <i data-lucide="layers" class="w-5 h-5 text-ra-verde"></i>
                        Matriz Consolidada dos Coordenadores
                    </h3>
                    <p class="text-xs text-ra-grafite/70 mt-0.5">
                        Resumo completo de volume, cobertura de HC elegível (&gt;30d e sem terceiros), faixas de meta (&ge;85, 75-84,9 e &lt;75) e desvios
                    </p>
                </div>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead class="bg-ra-roxo text-white text-[11px] uppercase tracking-wider">
                        <tr>
                            <th class="py-3.5 px-4 font-semibold">Coordenação</th>
                            <th class="py-3.5 px-3 font-semibold">Gerente</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Monitorias (Voz / Texto)</th>
                            <th class="py-3.5 px-3 font-semibold text-center">HC Avaliado / Elegíveis (&gt;30d)</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Sem Monitoria</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Nota Geral</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Nota Qual. | Oper.</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Acima Meta (&ge;85)</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Acima de 75 (75-84)</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Abaixo de 75 (&lt;75)</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Condução (NC)</th>
                            <th class="py-3.5 px-3 font-semibold text-center">NCG (Zera)</th>
                            <th class="py-3.5 px-3 font-semibold text-center">Bônus +6</th>
                        </tr>
                    </thead>
                    <tbody id="tbody-coordenacao" class="divide-y divide-ra-bege text-xs text-ra-grafite">
                    </tbody>
                </table>
            </div>
        </section>

        <!-- GRÁFICO DE DESDOBRAMENTO POR SUPERVISOR -->
        <section class="bg-white p-5 rounded-2xl shadow-sm border border-ra-bege">
            <div class="mb-4">
                <h3 class="text-base font-bold text-ra-roxo flex items-center gap-2">
                    <i data-lucide="git-branch" class="w-5 h-5 text-ra-verde"></i>
                    Desempenho por Supervisor (Top 12 no Recorte)
                </h3>
                <p class="text-xs text-ra-grafite/70">
                    Visão de apoio para o Coordenador identificar os Supervisores de maior e menor nota
                </p>
            </div>
            <div class="h-80 relative">
                <canvas id="chartSupervisores"></canvas>
            </div>
        </section>

        <!-- RADAR DE ATENÇÃO GERENCIAL: OPERADORES ABAIXO DE 75 (< 75 pts) -->
        <section class="bg-white rounded-2xl shadow-sm border-2 border-ra-vermelho/25 overflow-hidden">
            <div class="p-5 border-b border-ra-bege bg-ra-vermelho/5 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                <div>
                    <h3 class="text-base font-bold text-ra-vermelho flex items-center gap-2">
                        <i data-lucide="alert-octagon" class="w-5 h-5 text-ra-vermelho"></i>
                        Pontos de Atenção para os Coordenadores — Colaboradores Abaixo de 75 (&lt; 75 pts)
                    </h3>
                    <p class="text-xs text-ra-grafite/80 mt-0.5">
                        Lista nominal de operadores com média abaixo de 75 pontos no período selecionado para direcionamento de feedback
                    </p>
                </div>
                <span id="badge-qtd-criticos" class="text-xs font-extrabold px-3 py-1.5 rounded-xl bg-ra-vermelho text-white self-start sm:self-center">
                    0 colaboradores abaixo de 75
                </span>
            </div>
            <div class="overflow-x-auto custom-scrollbar max-h-80">
                <table class="w-full text-left border-collapse">
                    <thead class="bg-ra-roxoEscuro text-white text-[11px] uppercase sticky top-0">
                        <tr>
                            <th class="py-3 px-4 font-semibold">Coordenação</th>
                            <th class="py-3 px-4 font-semibold">Supervisor</th>
                            <th class="py-3 px-4 font-semibold">Operador(a)</th>
                            <th class="py-3 px-4 font-semibold text-center">Qtd Monitorias</th>
                            <th class="py-3 px-4 font-semibold text-center">Nota Média no Recorte</th>
                            <th class="py-3 px-4 font-semibold text-center">Faixa no Mês</th>
                        </tr>
                    </thead>
                    <tbody id="tbody-criticos" class="divide-y divide-ra-bege text-xs text-ra-grafite">
                    </tbody>
                </table>
            </div>
        </section>

        <!-- TABELA DETALHADA DE MONITORIAS + AUDITORIA DE ESCOPO -->
        <section class="bg-white rounded-2xl shadow-sm border border-ra-bege overflow-hidden">
            <div class="p-5 border-b border-ra-bege flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                <div>
                    <h3 class="text-base font-bold text-ra-roxo flex items-center gap-2">
                        <i data-lucide="table-2" class="w-5 h-5 text-ra-verde"></i>
                        Base Nominal de Monitorias (Auditoria &amp; Exportação)
                    </h3>
                    <p class="text-xs text-ra-grafite/70 mt-0.5">
                        Exibindo até 500 avaliações em tela • Exporte em CSV para abrir no Excel todas as avaliações filtradas
                    </p>
                </div>
                <div class="flex flex-wrap items-center gap-3">
                    <div class="relative">
                        <input type="text" id="input-busca" oninput="aplicarFiltros()" placeholder="Buscar coordenador, agente, ID..." class="text-xs border border-ra-bege rounded-xl pl-9 pr-3 py-2.5 bg-ra-offwhite focus:ring-2 focus:ring-ra-roxo focus:outline-none w-64">
                        <i data-lucide="search" class="w-4 h-4 text-ra-roxo absolute left-3 top-2.5"></i>
                    </div>
                    <button onclick="exportarCSV()" class="bg-ra-verde hover:bg-ra-verdeEscuro text-ra-roxoEscuro hover:text-white text-xs font-extrabold px-4 py-2.5 rounded-xl shadow-sm flex items-center gap-1.5 transition">
                        <i data-lucide="download" class="w-4 h-4"></i>
                        Exportar Excel (CSV)
                    </button>
                </div>
            </div>

            <div class="overflow-x-auto custom-scrollbar max-h-[460px]">
                <table class="w-full text-left border-collapse">
                    <thead class="bg-ra-grafite text-white text-[11px] uppercase sticky top-0 z-10">
                        <tr>
                            <th class="py-3 px-3 font-semibold">Data</th>
                            <th class="py-3 px-3 font-semibold">ID Interação</th>
                            <th class="py-3 px-3 font-semibold">Gerente</th>
                            <th class="py-3 px-3 font-semibold">Coordenação</th>
                            <th class="py-3 px-3 font-semibold">Supervisor</th>
                            <th class="py-3 px-3 font-semibold">Colaborador</th>
                            <th class="py-3 px-3 font-semibold text-center">Alocação</th>
                            <th class="py-3 px-3 font-semibold text-center">Canal</th>
                            <th class="py-3 px-3 font-semibold text-center">Origem</th>
                            <th class="py-3 px-3 font-semibold text-center">Nota Final</th>
                            <th class="py-3 px-3 font-semibold text-center">Faixa de Meta</th>
                            <th class="py-3 px-3 font-semibold text-center">Bônus (+6)</th>
                            <th class="py-3 px-3 font-semibold">Validação / Regra Escopo</th>
                        </tr>
                    </thead>
                    <tbody id="tbody-monitorias" class="divide-y divide-ra-bege text-xs text-ra-grafite">
                    </tbody>
                </table>
            </div>
            <div class="bg-ra-offwhite px-5 py-3 border-t border-ra-bege text-xs text-ra-grafite flex flex-col sm:flex-row justify-between items-center gap-2">
                <span id="lbl-contagem-tabela" class="font-semibold">Mostrando 0 registros</span>
                <span id="lbl-diag-bonus" class="font-mono text-[11px] text-ra-roxo"></span>
            </div>
        </section>

    </main>

    <!-- RODAPÉ INSTITUCIONAL -->
    <footer class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 text-center text-xs text-ra-grafite/60">
        Rede Américas • Paixão por Cuidar • Gestão Nacional de Qualidade CX
    </footer>

    <!-- LÓGICA JAVASCRIPT DO PAINEL EXECUTIVO -->
    <script>
        Chart.register(ChartDataLabels);
        Chart.defaults.font.family = "'Inter', 'Arial', sans-serif";
        Chart.defaults.color = "#343434";

        const DADOS_RELATORIO = {dados_json_str};

        const ehExpurgada = v => v === "exclusão feita pela qualidade" || v === "expurgo por regra de negócio";

        const ehAcimaMeta = q => String(q).includes("Acima da Meta");
        const ehAcima75 = q => String(q).includes("Acima de 75");
        const ehAbaixo75 = q => String(q).includes("Abaixo de 75");

        const CORES_RA = {{
            roxo: "#6D2077",
            roxoEscuro: "#422140",
            lavanda: "#D9B4D5",
            verde: "#84BD00",
            verdeEscuro: "#4F7000",
            offwhite: "#F1F1EF",
            grafite: "#343434",
            tomRoxo: "#A76FA9",
            tomLavanda: "#CCA2C9",
            tomVerde: "#A9B95C",
            amarelo: "#FFD661",
            coral: "#F0734F",
            vermelho: "#A11430",
            bege: "#E8E0D6"
        }};

        function criarIconeLinhaLegenda() {{
            const c = document.createElement("canvas");
            c.width = 28;
            c.height = 14;
            const ctx = c.getContext("2d");

            ctx.strokeStyle = CORES_RA.roxo;
            ctx.lineWidth = 3;
            ctx.lineCap = "round";
            ctx.beginPath();
            ctx.moveTo(2, 7);
            ctx.lineTo(26, 7);
            ctx.stroke();

            ctx.fillStyle = CORES_RA.verde;
            ctx.strokeStyle = CORES_RA.roxo;
            ctx.lineWidth = 2;
            ctx.beginPath();
            ctx.arc(14, 7, 4.5, 0, Math.PI * 2);
            ctx.fill();
            ctx.stroke();

            return c;
        }}

        function criarIconeMetaTracejadaLegenda() {{
            const c = document.createElement("canvas");
            c.width = 28;
            c.height = 14;
            const ctx = c.getContext("2d");

            ctx.strokeStyle = CORES_RA.verdeEscuro;
            ctx.lineWidth = 2.5;
            ctx.setLineDash([5, 4]);
            ctx.beginPath();
            ctx.moveTo(2, 7);
            ctx.lineTo(26, 7);
            ctx.stroke();

            return c;
        }}

        function criarIconeBarraLegenda() {{
            const c = document.createElement("canvas");
            c.width = 24;
            c.height = 14;
            const ctx = c.getContext("2d");

            ctx.fillStyle = CORES_RA.lavanda;
            ctx.beginPath();
            if (ctx.roundRect) {{
                ctx.roundRect(2, 2, 20, 10, 3);
            }} else {{
                ctx.rect(2, 2, 20, 10);
            }}
            ctx.fill();

            return c;
        }}

        const FILTROS_MULTI = {{
            mes: new Set(),
            data: new Set(),
            gerente: new Set(),
            coordenador: new Set(),
            supervisor: new Set(),
            canal: new Set(),
            origem: new Set(),
            quartil: new Set()
        }};

        const CONFIG_SLICERS = [
            {{ chave: "mes", containerId: "slicer-mes", titulo: "Mês / Ano", placeholder: "Todos os Meses", campo: "mes_ano" }},
            {{ chave: "data", containerId: "slicer-data", titulo: "Dia / Data Avaliação", placeholder: "Todos os Dias", campo: "data_avaliacao" }},
            {{ chave: "gerente", containerId: "slicer-gerente", titulo: "Gerente", placeholder: "Todos os Gerentes", campo: "gerente" }},
            {{ chave: "coordenador", containerId: "slicer-coordenador", titulo: "Coordenação (Multi)", placeholder: "Todas as Coordenações", campo: "coordenador", destaque: true }},
            {{ chave: "supervisor", containerId: "slicer-supervisor", titulo: "Supervisor(a)", placeholder: "Todos os Supervisores", campo: "supervisor" }},
            {{ chave: "canal", containerId: "slicer-canal", titulo: "Canal (Formulário)", placeholder: "Voz + Texto", campo: "canal" }},
            {{ chave: "origem", containerId: "slicer-origem", titulo: "Avaliador (Origem)", placeholder: "Qualidade + Operação", campo: "origem" }},
            {{ chave: "quartil", containerId: "slicer-quartil", titulo: "Faixa de Meta", placeholder: "Todas as Faixas", campo: "quartil" }}
        ];

        let slicerAbertoAtual = null;
        let chartCoordenadoresInstance = null;
        let chartQuartisCoordInstance = null;
        let chartComparativoOrigemInstance = null;
        let chartSupervisoresInstance = null;
        let chartEvolucaoInstance = null;
        let chartConducaoInstance = null;
        let chartNCGInstance = null;
        let dadosFiltradosAtuais = [];
        let ofensoresFiltradosAtuais = [];
        let pendentesSemMonitoriaAtuais = [];
        let ultimoAssuntoEmail = "";
        let ultimoBlobImagemCompleta = null;

        function inicializarPainel() {{
            lucide.createIcons();
            document.getElementById("lbl-data-geracao").textContent = DADOS_RELATORIO.gerado_em || "-";

            document.addEventListener("click", function(e) {{
                if (!e.target.closest(".pbi-slicer-wrapper")) {{
                    fecharTodosSlicers();
                }}
            }});

            verificarAlertas();
            construirSlicersMultiSelect();
            criarBotoesGerentes();
            aplicarFiltros();
        }}

        function fecharTodosSlicers() {{
            document.querySelectorAll(".pbi-dropdown-menu").forEach(el => el.classList.add("hidden"));
            slicerAbertoAtual = null;
        }}

        function toggleSlicerMenu(chave, event) {{
            event.stopPropagation();
            const menu = document.getElementById(`menu-slicer-${{chave}}`);
            const estavaOculto = menu.classList.contains("hidden");
            fecharTodosSlicers();
            if (estavaOculto) {{
                menu.classList.remove("hidden");
                slicerAbertoAtual = chave;
                const inp = document.getElementById(`busca-slicer-${{chave}}`);
                if (inp) inp.focus();
            }}
        }}

        function obterBasePorStatusEscopo() {{
            const avs = DADOS_RELATORIO.avaliacoes || [];
            const fVal = document.getElementById("filtro-validacao").value;
            return avs.filter(a => {{
                if (fVal === "VALIDAS" && ehExpurgada(a.validacao)) return false;
                if (fVal === "EXPURGADAS" && !ehExpurgada(a.validacao)) return false;
                return true;
            }});
        }}

        function construirSlicersMultiSelect() {{
            const base = obterBasePorStatusEscopo();

            CONFIG_SLICERS.forEach(cfg => {{
                const wrapper = document.getElementById(cfg.containerId);
                if (!wrapper) return;

                const baseContextual = base.filter(item => {{
                    if (cfg.chave !== "gerente" && FILTROS_MULTI.gerente.size > 0 && !FILTROS_MULTI.gerente.has(item.gerente)) return false;
                    if (cfg.chave !== "mes" && FILTROS_MULTI.mes.size > 0 && !FILTROS_MULTI.mes.has(item.mes_ano)) return false;
                    if (cfg.chave === "supervisor" && FILTROS_MULTI.coordenador.size > 0 && !FILTROS_MULTI.coordenador.has(item.coordenador)) return false;
                    return true;
                }});

                let opcoes = Array.from(new Set(
                    baseContextual.map(i => i[cfg.campo]).filter(v => v && v !== "N/A" && v !== "nan")
                )).sort();

                if (cfg.chave === "data" || cfg.chave === "mes") {{
                    opcoes = opcoes.sort().reverse();
                }}

                const setAtual = FILTROS_MULTI[cfg.chave];
                Array.from(setAtual).forEach(val => {{
                    if (!opcoes.includes(val)) setAtual.delete(val);
                }});

                let textoBotao = cfg.placeholder;
                if (setAtual.size === 1) {{
                    textoBotao = Array.from(setAtual)[0];
                }} else if (setAtual.size > 1) {{
                    textoBotao = `${{setAtual.size}} selecionados`;
                }}

                const temFiltroAtivo = setAtual.size > 0;
                const corBorda = temFiltroAtivo
                    ? "border-ra-roxo bg-ra-lavanda/25 text-ra-roxo font-bold"
                    : (cfg.destaque ? "border-ra-roxo/40 bg-ra-lavanda/10 text-ra-roxo font-semibold" : "border-ra-bege bg-ra-offwhite text-ra-grafite font-medium");

                const listaHTML = opcoes.map(op => {{
                    const checked = setAtual.has(op) ? "checked" : "";
                    const opEsc = String(op).replace(/"/g, "&quot;");
                    return `
                        <label class="item-slicer-${{cfg.chave}} flex items-center gap-2.5 px-2.5 py-1.5 rounded-lg hover:bg-ra-lavanda/25 cursor-pointer text-xs text-ra-grafite select-none" data-valor="${{opEsc.toLowerCase()}}">
                            <input type="checkbox" value="${{opEsc}}" ${{checked}}
                                   onchange="alternarOpcaoSlicer('${{cfg.chave}}', this.value, this.checked)"
                                   class="w-3.5 h-3.5 rounded text-ra-roxo focus:ring-ra-roxo border-ra-bege accent-[#6D2077]">
                            <span class="truncate" title="${{opEsc}}">${{op}}</span>
                        </label>
                    `;
                }}).join("");

                wrapper.className = "relative pbi-slicer-wrapper";
                wrapper.innerHTML = `
                    <label class="block text-[11px] font-bold ${{cfg.destaque ? 'text-ra-roxo' : 'text-ra-grafite'}} mb-1 uppercase tracking-wide flex items-center justify-between">
                        <span>${{cfg.titulo}}</span>
                        ${{temFiltroAtivo ? `<span class="text-[10px] px-1.5 py-0.2 rounded bg-ra-roxo text-white">${{setAtual.size}}</span>` : ""}}
                    </label>
                    <button type="button" onclick="toggleSlicerMenu('${{cfg.chave}}', event)"
                            class="w-full text-xs rounded-xl px-3 py-2 border flex items-center justify-between gap-2 transition focus:outline-none ${{corBorda}}">
                        <span class="truncate">${{textoBotao}}</span>
                        <span class="text-[10px] opacity-70">▼</span>
                    </button>

                    <div id="menu-slicer-${{cfg.chave}}" onclick="event.stopPropagation()"
                         class="pbi-dropdown-menu hidden absolute left-0 right-0 mt-1.5 bg-white border border-ra-bege rounded-2xl z-50 p-3 flex flex-col gap-2">
                        <div class="flex items-center justify-between gap-1 border-b border-ra-bege pb-2">
                            <button type="button" onclick="marcarTodosSlicer('${{cfg.chave}}')"
                                    class="text-[11px] font-bold text-ra-roxo hover:underline px-1.5 py-0.5 rounded hover:bg-ra-offwhite">
                                Selecionar Todos
                            </button>
                            <button type="button" onclick="limparSlicer('${{cfg.chave}}')"
                                    class="text-[11px] font-bold text-ra-vermelho hover:underline px-1.5 py-0.5 rounded hover:bg-ra-offwhite">
                                Limpar (${{setAtual.size}})
                            </button>
                        </div>
                        <input type="text" id="busca-slicer-${{cfg.chave}}"
                               oninput="filtrarOpcoesDentroDoSlicer('${{cfg.chave}}', this.value)"
                               placeholder="Pesquisar em ${{cfg.titulo}}..."
                               class="w-full text-xs border border-ra-bege rounded-lg px-2.5 py-1.5 bg-ra-offwhite focus:outline-none focus:ring-1 focus:ring-ra-roxo">
                        <div class="overflow-y-auto custom-scrollbar max-h-44 divide-y divide-ra-offwhite">
                            ${{listaHTML || '<p class="text-xs text-center py-3 text-ra-grafite/50">Sem opções</p>'}}
                        </div>
                    </div>
                `;
            }});
        }}

        function filtrarOpcoesDentroDoSlicer(chave, texto) {{
            const termo = (texto || "").trim().toLowerCase();
            document.querySelectorAll(`.item-slicer-${{chave}}`).forEach(lbl => {{
                const val = lbl.getAttribute("data-valor") || "";
                lbl.style.display = val.includes(termo) ? "flex" : "none";
            }});
        }}

        function alternarOpcaoSlicer(chave, valor, marcado) {{
            if (marcado) {{
                FILTROS_MULTI[chave].add(valor);
            }} else {{
                FILTROS_MULTI[chave].delete(valor);
            }}
            const menuEstavaAberto = slicerAbertoAtual;
            construirSlicersMultiSelect();
            criarBotoesGerentes();
            if (menuEstavaAberto) {{
                const menu = document.getElementById(`menu-slicer-${{menuEstavaAberto}}`);
                if (menu) menu.classList.remove("hidden");
                slicerAbertoAtual = menuEstavaAberto;
            }}
            aplicarFiltros();
        }}

        function marcarTodosSlicer(chave) {{
            const checkboxes = document.querySelectorAll(`.item-slicer-${{chave}} input[type="checkbox"]`);
            checkboxes.forEach(cb => {{
                if (cb.closest("label").style.display !== "none") {{
                    FILTROS_MULTI[chave].add(cb.value);
                }}
            }});
            construirSlicersMultiSelect();
            criarBotoesGerentes();
            const menu = document.getElementById(`menu-slicer-${{chave}}`);
            if (menu) menu.classList.remove("hidden");
            slicerAbertoAtual = chave;
            aplicarFiltros();
        }}

        function limparSlicer(chave) {{
            FILTROS_MULTI[chave].clear();
            construirSlicersMultiSelect();
            criarBotoesGerentes();
            const menu = document.getElementById(`menu-slicer-${{chave}}`);
            if (menu) menu.classList.remove("hidden");
            slicerAbertoAtual = chave;
            aplicarFiltros();
        }}

        function aoMudarStatusEscopo() {{
            construirSlicersMultiSelect();
            criarBotoesGerentes();
            aplicarFiltros();
        }}

        function alternarTabelaSemMonitoria() {{
            const box = document.getElementById("box-tabela-sem-monitoria");
            const txt = document.getElementById("txt-toggle-sem-monitoria");
            const oculto = box.classList.contains("hidden");
            if (oculto) {{
                box.classList.remove("hidden");
                txt.textContent = "Recolher Lista Nominal";
            }} else {{
                box.classList.add("hidden");
                txt.textContent = "Expandir Lista Nominal";
            }}
        }}

        function verificarAlertas() {{
            const erros = DADOS_RELATORIO.erros || {{}};
            const alertas = [];
            if (erros.genesys) alertas.push("<strong>Erro Base Genesys:</strong> " + erros.genesys);
            if (erros.excel_hc) alertas.push("<strong>Erro Base Dimensão (data.xlsx):</strong> " + erros.excel_hc);
            if (erros.exclusoes && erros.exclusoes !== "None") alertas.push("<strong>Aviso Planilha Exclusões:</strong> " + erros.exclusoes);
            if (erros.itens) alertas.push("<strong>Aviso Base Itens (base1.xlsx):</strong> " + erros.itens);

            const box = document.getElementById("box-alertas");
            const lista = document.getElementById("lista-alertas");
            if (alertas.length > 0) {{
                lista.innerHTML = alertas.map(a => `<p>• ${{a}}</p>`).join("");
                box.classList.remove("hidden");
            }}

            const d = DADOS_RELATORIO.diag_bonus || {{}};
            const f = DADOS_RELATORIO.diag_filtros || {{}};
            const it = DADOS_RELATORIO.diag_itens || {{}};
            let textoDiag = "";
            if (d.col_pergunta) {{
                textoDiag += `Bônus +6 Aplicados: ${{d.bonus_aplicados || 0}}`;
            }}
            if (it.total_ofensores_validos !== undefined) {{
                textoDiag += ` | Ofensores -> Condução (NC): ${{it.total_conducao || 0}} | NCG: ${{it.total_ncg || 0}}`;
            }}
            if (f.total_antes !== undefined) {{
                textoDiag += ` | Funil (${{f.total_antes}}) -> Válidas no Escopo: ${{f.mantidas_validas}}`;
            }}
            document.getElementById("lbl-diag-bonus").textContent = textoDiag;
        }}

        function criarBotoesGerentes() {{
            const avs = DADOS_RELATORIO.avaliacoes || [];
            const validas = avs.filter(a => !ehExpurgada(a.validacao));
            const gerentesUnicos = Array.from(new Set(
                validas.map(a => a.gerente).filter(g => g && g !== "Sem Gestor" && g !== "Não Cadastrado na Base")
            )).sort();

            const container = document.getElementById("botoes-gerentes");
            const todosAtivo = FILTROS_MULTI.gerente.size === 0;

            let html = `<button onclick="selecionarGerenteRapido('TODOS')" class="text-xs px-3.5 py-2 rounded-xl border transition ${{todosAtivo ? 'bg-ra-roxo text-white border-ra-roxo shadow-sm font-bold' : 'bg-ra-offwhite text-ra-grafite border-ra-bege hover:bg-ra-lavanda/30 font-semibold'}}">
                Todos os Gerentes
            </button>`;

            html += gerentesUnicos.map(g => {{
                const ativo = FILTROS_MULTI.gerente.has(g);
                const classes = ativo
                    ? "bg-ra-roxo text-white border-ra-roxo shadow-sm font-bold"
                    : "bg-ra-offwhite text-ra-grafite border-ra-bege hover:bg-ra-lavanda/30 font-semibold";
                return `<button onclick="selecionarGerenteRapido('${{g.replace(/'/g, "\\\\'")}}')" class="text-xs px-3.5 py-2 rounded-xl border transition ${{classes}}">
                    ${{g}}
                </button>`;
            }}).join("");

            container.innerHTML = html;
        }}

        function selecionarGerenteRapido(nomeGerente) {{
            if (nomeGerente === "TODOS") {{
                FILTROS_MULTI.gerente.clear();
            }} else {{
                if (FILTROS_MULTI.gerente.size === 1 && FILTROS_MULTI.gerente.has(nomeGerente)) {{
                    FILTROS_MULTI.gerente.clear();
                }} else {{
                    FILTROS_MULTI.gerente.clear();
                    FILTROS_MULTI.gerente.add(nomeGerente);
                }}
            }}
            FILTROS_MULTI.coordenador.clear();
            FILTROS_MULTI.supervisor.clear();
            construirSlicersMultiSelect();
            criarBotoesGerentes();
            aplicarFiltros();
        }}

        function selecionarCoordenadorCard(nomeCoord) {{
            if (FILTROS_MULTI.coordenador.has(nomeCoord)) {{
                FILTROS_MULTI.coordenador.delete(nomeCoord);
            }} else {{
                FILTROS_MULTI.coordenador.add(nomeCoord);
            }}
            construirSlicersMultiSelect();
            aplicarFiltros();
        }}

        function resetarFiltros() {{
            Object.values(FILTROS_MULTI).forEach(s => s.clear());
            document.getElementById("filtro-validacao").value = "VALIDAS";
            document.getElementById("filtro-tipo-ofensor").value = "TODOS";
            document.getElementById("filtro-status-sem-monitoria").value = "TODOS";
            document.getElementById("input-busca").value = "";
            construirSlicersMultiSelect();
            criarBotoesGerentes();
            aplicarFiltros();
        }}

        function passaFiltroMulti(setFiltro, valorItem) {{
            if (!setFiltro || setFiltro.size === 0) return true;
            return setFiltro.has(String(valorItem));
        }}

        function obterMesesReferenciaHC() {{
            if (FILTROS_MULTI.mes.size > 0) {{
                return FILTROS_MULTI.mes;
            }}
            const mRec = DADOS_RELATORIO.mes_mais_recente || "TODOS";
            return mRec !== "TODOS" ? new Set([mRec]) : new Set();
        }}

        function calcularPendentesSemMonitoria() {{
            const hcDim = DADOS_RELATORIO.hc_dimensao || [];
            const todasAvs = DADOS_RELATORIO.avaliacoes || [];
            const mesesRefSet = obterMesesReferenciaHC();
            const filtroStatusData = document.getElementById("filtro-status-sem-monitoria")
                ? document.getElementById("filtro-status-sem-monitoria").value
                : "TODOS";

            const txtMesRef = mesesRefSet.size > 0 ? Array.from(mesesRefSet).join(", ") : "Todos";
            document.getElementById("lbl-mes-ref-pendentes").textContent = txtMesRef;

            const loginsAvaliadosNoPeriodo = new Set(
                todasAvs
                    .filter(a => {{
                        if (ehExpurgada(a.validacao)) return false;
                        if (mesesRefSet.size > 0 && !mesesRefSet.has(a.mes_ano)) return false;
                        if (!passaFiltroMulti(FILTROS_MULTI.data, a.data_avaliacao)) return false;
                        return true;
                    }})
                    .map(a => a.login_unico)
            );

            const mapaElegiveisUnicos = new Map();
            hcDim.forEach(h => {{
                if (mesesRefSet.size > 0 && !mesesRefSet.has(h.mes_ano)) return;
                if (!passaFiltroMulti(FILTROS_MULTI.gerente, h.gerente)) return;
                if (!passaFiltroMulti(FILTROS_MULTI.coordenador, h.coordenador)) return;
                if (!passaFiltroMulti(FILTROS_MULTI.supervisor, h.supervisor)) return;
                if (filtroStatusData !== "TODOS" && h.ativo_flag !== filtroStatusData) return;

                if (!loginsAvaliadosNoPeriodo.has(h.login_unico)) {{
                    mapaElegiveisUnicos.set(h.login_unico, h);
                }}
            }});

            return Array.from(mapaElegiveisUnicos.values()).sort((a, b) => {{
                if (a.coordenador !== b.coordenador) return a.coordenador.localeCompare(b.coordenador);
                if (a.supervisor !== b.supervisor) return a.supervisor.localeCompare(b.supervisor);
                return a.nome.localeCompare(b.nome);
            }});
        }}

        function atualizarSecaoSemMonitoria(listaPendentes) {{
            const qtdAtivos = listaPendentes.filter(p => p.ativo_flag === "SIM").length;
            const qtdInativos = listaPendentes.length - qtdAtivos;
            document.getElementById("badge-qtd-sem-monitoria").textContent =
                `${{listaPendentes.length}} sem monitoria (Ativos: ${{qtdAtivos}} | Inativos: ${{qtdInativos}})`;

            const resumoContainer = document.getElementById("resumo-pendentes-coord");
            const contagemCoord = {{}};
            listaPendentes.forEach(p => {{
                contagemCoord[p.coordenador] = (contagemCoord[p.coordenador] || 0) + 1;
            }});

            const coordsOrd = Object.keys(contagemCoord).sort((a, b) => contagemCoord[b] - contagemCoord[a]);
            if (coordsOrd.length === 0) {{
                resumoContainer.innerHTML = `<span class="text-xs font-bold text-ra-verdeEscuro">100% dos operadores elegíveis possuem monitoria no período!</span>`;
            }} else {{
                resumoContainer.innerHTML = coordsOrd.map(c => `
                    <span class="text-xs px-3 py-1 rounded-xl bg-white border border-ra-bege text-ra-grafite shadow-sm">
                        <strong class="text-ra-roxo">${{c}}:</strong> ${{contagemCoord[c]}} sem avaliação
                    </span>
                `).join("");
            }}

            const tbody = document.getElementById("tbody-sem-monitoria");
            if (listaPendentes.length === 0) {{
                tbody.innerHTML = `<tr><td colspan="10" class="py-6 text-center text-ra-verdeEscuro font-semibold">Excelente! Todos os operadores elegíveis (&gt;30 dias, exceto terceiros e liderança) já foram avaliados no período selecionado.</td></tr>`;
                return;
            }}

            tbody.innerHTML = listaPendentes.map(p => {{
                const badgeStatus = p.ativo_flag === "SIM"
                    ? `<span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-ra-verde/25 text-ra-verdeEscuro">${{p.status_data || 'Ativo'}}</span>`
                    : `<span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-ra-vermelho/15 text-ra-vermelho">${{p.status_data || 'Inativo'}}</span>`;

                return `
                <tr class="hover:bg-ra-lavanda/15 transition">
                    <td class="py-2.5 px-4 font-bold text-ra-roxo">${{p.coordenador}}</td>
                    <td class="py-2.5 px-4 text-ra-grafite">${{p.supervisor}}</td>
                    <td class="py-2.5 px-4 font-semibold text-ra-grafite">${{p.nome}}</td>
                    <td class="py-2.5 px-4 text-ra-grafite/90 font-medium">${{p.cargo || 'N/A'}}</td>
                    <td class="py-2.5 px-4 text-center whitespace-nowrap">${{badgeStatus}}</td>
                    <td class="py-2.5 px-4 text-center"><span class="px-2 py-0.5 rounded bg-ra-offwhite text-[10px] font-bold text-ra-roxoEscuro">${{p.alocacao || 'PRÓPRIO'}}</span></td>
                    <td class="py-2.5 px-4 text-center font-mono text-[11px] text-ra-grafite/80">${{p.matricula !== "N/A" ? p.matricula : p.login_unico}}</td>
                    <td class="py-2.5 px-4 text-center font-mono">${{p.admissao}}</td>
                    <td class="py-2.5 px-4 text-center font-bold text-ra-roxoEscuro">${{p.dias_casa_dim !== "N/A" ? p.dias_casa_dim + " dias" : "&gt;30 dias"}}</td>
                    <td class="py-2.5 px-4 text-center font-mono text-[11px]">${{p.mes_ano}}</td>
                </tr>
            `;
            }}).join("");
        }}

        function exportarSemMonitoriaCSV() {{
            if (!pendentesSemMonitoriaAtuais || pendentesSemMonitoriaAtuais.length === 0) {{
                alert("Não há operadores pendentes sem monitoria no filtro atual.");
                return;
            }}

            const cabecalhos = [
                "Mes_Referencia", "Gerente", "Coordenador", "Supervisor",
                "Operador_Sem_Monitoria", "Cargo", "Status_Ativo_Data",
                "Alocacao", "Matricula", "Login_Genesys", "Data_Admissao", "Dias_de_Casa"
            ];

            const linhas = pendentesSemMonitoriaAtuais.map(p => [
                p.mes_ano,
                `"${{(p.gerente || "").replace(/"/g, '""')}}"`,
                `"${{(p.coordenador || "").replace(/"/g, '""')}}"`,
                `"${{(p.supervisor || "").replace(/"/g, '""')}}"`,
                `"${{(p.nome || "").replace(/"/g, '""')}}"`,
                `"${{(p.cargo || "N/A").replace(/"/g, '""')}}"`,
                `"${{(p.status_data || "Ativo").replace(/"/g, '""')}}"`,
                `"${{(p.alocacao || "PRÓPRIO").replace(/"/g, '""')}}"`,
                p.matricula,
                p.login_unico,
                p.admissao,
                p.dias_casa_dim
            ].join(";"));

            const conteudoCSV = "\\uFEFF" + [cabecalhos.join(";"), ...linhas].join("\\n");
            const blob = new Blob([conteudoCSV], {{ type: "text/csv;charset=utf-8;" }});
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `Operadores_Elegiveis_Sem_Monitoria_${{new Date().toISOString().slice(0,10)}}.csv`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }}

        function aplicarFiltros() {{
            const fVal = document.getElementById("filtro-validacao").value;
            const termoBusca = (document.getElementById("input-busca").value || "").trim().toLowerCase();

            const badgeCoord = document.getElementById("lbl-filtro-coord-ativo");
            if (FILTROS_MULTI.coordenador.size > 0) {{
                const nomes = Array.from(FILTROS_MULTI.coordenador).join(" | ");
                badgeCoord.textContent = `Coordenações selecionadas (${{FILTROS_MULTI.coordenador.size}}): ${{nomes}}`;
                badgeCoord.classList.remove("hidden");
            }} else {{
                badgeCoord.classList.add("hidden");
            }}

            const todas = DADOS_RELATORIO.avaliacoes || [];

            dadosFiltradosAtuais = todas.filter(item => {{
                if (fVal === "VALIDAS" && ehExpurgada(item.validacao)) return false;
                if (fVal === "EXPURGADAS" && !ehExpurgada(item.validacao)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.mes, item.mes_ano)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.data, item.data_avaliacao)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.gerente, item.gerente)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.coordenador, item.coordenador)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.supervisor, item.supervisor)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.canal, item.canal)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.origem, item.origem)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.quartil, item.quartil)) return false;

                if (termoBusca) {{
                    const alvo = `${{item.gerente}} ${{item.coordenador}} ${{item.supervisor}} ${{item.nome_agente}} ${{item.id_interacao}} ${{item.avaliador}} ${{item.motivo_expurgo}}`.toLowerCase();
                    if (!alvo.includes(termoBusca)) return false;
                }}
                return true;
            }});

            const todosOfensores = DADOS_RELATORIO.ofensores || [];
            ofensoresFiltradosAtuais = todosOfensores.filter(of => {{
                if (!passaFiltroMulti(FILTROS_MULTI.mes, of.mes_ano)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.data, of.data_avaliacao)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.gerente, of.gerente)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.coordenador, of.coordenador)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.supervisor, of.supervisor)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.canal, of.canal)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.origem, of.origem)) return false;
                if (!passaFiltroMulti(FILTROS_MULTI.quartil, of.quartil)) return false;
                return true;
            }});

            pendentesSemMonitoriaAtuais = calcularPendentesSemMonitoria();

            atualizarKPIs(dadosFiltradosAtuais, ofensoresFiltradosAtuais, pendentesSemMonitoriaAtuais);
            renderizarCardsCoordenadores(dadosFiltradosAtuais, ofensoresFiltradosAtuais);
            atualizarSecaoSemMonitoria(pendentesSemMonitoriaAtuais);
            atualizarTabelaCoordenacao(dadosFiltradosAtuais, ofensoresFiltradosAtuais);
            atualizarGraficos(dadosFiltradosAtuais);
            atualizarSecaoOfensores();
            atualizarTabelaCriticos(dadosFiltradosAtuais);
            atualizarTabelaNominal(dadosFiltradosAtuais);
        }}

        function atualizarKPIs(lista, listaOfensores, listaPendentes) {{
            const comNota = lista.filter(i => i.nota !== null && !isNaN(i.nota));
            const somaNotas = comNota.reduce((acc, cur) => acc + cur.nota, 0);
            const media = comNota.length > 0 ? (somaNotas / comNota.length) : 0;

            const avQual = comNota.filter(i => i.origem === "Qualidade");
            const avOper = comNota.filter(i => i.origem === "Operação");
            const mediaQual = avQual.length > 0 ? (avQual.reduce((a, c) => a + c.nota, 0) / avQual.length).toFixed(1).replace(".", ",") : "-";
            const mediaOper = avOper.length > 0 ? (avOper.reduce((a, c) => a + c.nota, 0) / avOper.length).toFixed(1).replace(".", ",") : "-";

            document.getElementById("kpi-nota-media").textContent = media.toFixed(1).replace(".", ",");
            document.getElementById("kpi-sub-calibracao").textContent = `Qualidade: ${{mediaQual}} | Operação: ${{mediaOper}}`;

            document.getElementById("kpi-total-validas").textContent = lista.length.toLocaleString("pt-BR");
            const qtdVoz = lista.filter(i => i.canal === "Voz").length;
            const qtdTexto = lista.filter(i => i.canal === "Texto").length;
            document.getElementById("kpi-sub-canal").textContent = `Voz: ${{qtdVoz}} | Texto: ${{qtdTexto}}`;

            const mapaAgentes = new Map();
            comNota.forEach(i => {{
                const chaveAg = i.login_unico && i.login_unico !== "N/A" ? i.login_unico : i.id_chave;
                if (chaveAg && chaveAg !== "N/A") {{
                    mapaAgentes.set(chaveAg, i.quartil);
                }}
            }});
            const totalAgentesUnicos = mapaAgentes.size;
            document.getElementById("kpi-agentes-avaliados").textContent = totalAgentesUnicos.toLocaleString("pt-BR");
            document.getElementById("kpi-sub-cobertura").textContent = `Sem monitoria no período: ${{listaPendentes.length}} ag.`;

            let agAcimaMeta = 0;
            let agAcima75 = 0;
            let agAbaixo75 = 0;
            mapaAgentes.forEach(q => {{
                if (ehAcimaMeta(q)) agAcimaMeta++;
                else if (ehAcima75(q)) agAcima75++;
                else if (ehAbaixo75(q)) agAbaixo75++;
            }});
            const percMeta = totalAgentesUnicos > 0 ? Math.round((agAcimaMeta / totalAgentesUnicos) * 100) : 0;
            document.getElementById("kpi-perc-meta").textContent = `${{percMeta}}%`;
            document.getElementById("kpi-sub-faixas").textContent = `>=85: ${{agAcimaMeta}} | 75-84: ${{agAcima75}} | <75: ${{agAbaixo75}}`;

            const qtdConducao = listaOfensores.filter(o => o.categoria === "Condução").length;
            const qtdNCG = listaOfensores.filter(o => o.categoria === "NCG").length;
            document.getElementById("kpi-total-ofensores").textContent = (qtdConducao + qtdNCG).toLocaleString("pt-BR");
            document.getElementById("kpi-sub-ofensores").textContent = `Condução: ${{qtdConducao}} | NCG: ${{qtdNCG}}`;
        }}

        function consolidarPorCoordenador(lista, listaOfensores = []) {{
            const coordMap = {{}};
            const hcDim = DADOS_RELATORIO.hc_dimensao || [];
            const todasAvs = DADOS_RELATORIO.avaliacoes || [];
            const mesesRefSet = obterMesesReferenciaHC();
            const filtroStatusData = document.getElementById("filtro-status-sem-monitoria")
                ? document.getElementById("filtro-status-sem-monitoria").value
                : "TODOS";

            const loginsAvaliadosNoPeriodo = new Set(
                todasAvs
                    .filter(a => {{
                        if (ehExpurgada(a.validacao)) return false;
                        if (mesesRefSet.size > 0 && !mesesRefSet.has(a.mes_ano)) return false;
                        if (!passaFiltroMulti(FILTROS_MULTI.data, a.data_avaliacao)) return false;
                        return true;
                    }})
                    .map(a => a.login_unico)
            );

            lista.forEach(i => {{
                const c = i.coordenador || "Sem Gestor";
                if (!coordMap[c]) {{
                    coordMap[c] = {{
                        coordenador: c,
                        gerente: i.gerente || "Sem Gestor",
                        qtdMonitorias: 0,
                        qtdVoz: 0,
                        qtdTexto: 0,
                        somaNota: 0,
                        qtdComNota: 0,
                        somaQual: 0,
                        qtdQual: 0,
                        somaOper: 0,
                        qtdOper: 0,
                        agentesMap: new Map(),
                        bonus: 0,
                        desviosConducao: 0,
                        ocorrenciasNCG: 0,
                        mapaPerguntasConducao: {{}},
                        mapaPerguntasNCG: {{}}
                    }};
                }}
                coordMap[c].qtdMonitorias += 1;
                if (i.canal === "Voz") coordMap[c].qtdVoz += 1;
                if (i.canal === "Texto") coordMap[c].qtdTexto += 1;

                const chaveAg = i.login_unico && i.login_unico !== "N/A" ? i.login_unico : i.id_chave;
                if (chaveAg && chaveAg !== "N/A") {{
                    coordMap[c].agentesMap.set(chaveAg, i.quartil);
                }}

                if (i.nota !== null && !isNaN(i.nota)) {{
                    coordMap[c].somaNota += i.nota;
                    coordMap[c].qtdComNota += 1;
                    if (i.origem === "Qualidade") {{
                        coordMap[c].somaQual += i.nota;
                        coordMap[c].qtdQual += 1;
                    }} else if (i.origem === "Operação") {{
                        coordMap[c].somaOper += i.nota;
                        coordMap[c].qtdOper += 1;
                    }}
                }}

                if (i.status_bonus === "Bônus +6 Aplicado") coordMap[c].bonus += 1;
            }});

            listaOfensores.forEach(of => {{
                const c = of.coordenador || "Sem Gestor";
                if (coordMap[c]) {{
                    const p = of.pergunta || "Item não especificado";
                    if (of.categoria === "NCG") {{
                        coordMap[c].ocorrenciasNCG += 1;
                        coordMap[c].mapaPerguntasNCG[p] = (coordMap[c].mapaPerguntasNCG[p] || 0) + 1;
                    }} else {{
                        coordMap[c].desviosConducao += 1;
                        coordMap[c].mapaPerguntasConducao[p] = (coordMap[c].mapaPerguntasConducao[p] || 0) + 1;
                    }}
                }}
            }});

            return Object.values(coordMap).map(x => {{
                let acimaMeta = 0, acima75 = 0, abaixo75 = 0;
                x.agentesMap.forEach(faixa => {{
                    if (ehAcimaMeta(faixa)) acimaMeta++;
                    else if (ehAcima75(faixa)) acima75++;
                    else if (ehAbaixo75(faixa)) abaixo75++;
                }});

                const loginsElegiveisCoord = Array.from(new Set(
                    hcDim.filter(h =>
                        h.coordenador === x.coordenador &&
                        (mesesRefSet.size === 0 || mesesRefSet.has(h.mes_ano)) &&
                        (filtroStatusData === "TODOS" || h.ativo_flag === filtroStatusData)
                    ).map(h => h.login_unico)
                ));

                const hcElegivelMes = loginsElegiveisCoord.length;
                const qtdSemMonitoria = loginsElegiveisCoord.filter(lg => !loginsAvaliadosNoPeriodo.has(lg)).length;

                const hcAvaliado = x.agentesMap.size;
                const percCobertura = hcElegivelMes > 0 ? Math.round((hcAvaliado / hcElegivelMes) * 100) : null;
                const mediaGeral = x.qtdComNota > 0 ? (x.somaNota / x.qtdComNota) : 0;
                const mediaQual = x.qtdQual > 0 ? (x.somaQual / x.qtdQual) : null;
                const mediaOper = x.qtdOper > 0 ? (x.somaOper / x.qtdOper) : null;

                const topConducao = Object.entries(x.mapaPerguntasConducao)
                    .map(([pergunta, qtd]) => ({{ pergunta, qtd }}))
                    .sort((a, b) => b.qtd - a.qtd)
                    .slice(0, 3);

                const topNCG = Object.entries(x.mapaPerguntasNCG)
                    .map(([pergunta, qtd]) => ({{ pergunta, qtd }}))
                    .sort((a, b) => b.qtd - a.qtd)
                    .slice(0, 2);

                return {{
                    ...x,
                    hcAvaliado,
                    hcElegivelMes,
                    qtdSemMonitoria,
                    percCobertura,
                    mediaGeral,
                    mediaQual,
                    mediaOper,
                    acimaMeta,
                    acima75,
                    abaixo75,
                    topConducao,
                    topNCG
                }};
            }}).sort((a, b) => b.mediaGeral - a.mediaGeral);
        }}

        function renderizarCardsCoordenadores(lista, listaOfensores) {{
            const coords = consolidarPorCoordenador(lista, listaOfensores);
            const grid = document.getElementById("grid-cards-coordenadores");

            if (coords.length === 0) {{
                grid.innerHTML = `<div class="col-span-3 bg-white p-6 rounded-2xl text-center text-sm text-ra-grafite/60 border border-ra-bege">
                    Nenhuma coordenação encontrada para os filtros selecionados.
                </div>`;
                return;
            }}

            grid.innerHTML = coords.map(c => {{
                const selecionado = FILTROS_MULTI.coordenador.has(c.coordenador);
                const classeAtivo = selecionado ? "ring-4 ring-ra-roxo bg-ra-lavanda/10" : "bg-white";
                const mediaStr = c.mediaGeral.toFixed(1).replace(".", ",");
                const qualStr = c.mediaQual !== null ? c.mediaQual.toFixed(1).replace(".", ",") : "-";
                const operStr = c.mediaOper !== null ? c.mediaOper.toFixed(1).replace(".", ",") : "-";

                const totalAg = Math.max(c.hcAvaliado, 1);
                const pMeta = Math.round((c.acimaMeta / totalAg) * 100);
                const p75 = Math.round((c.acima75 / totalAg) * 100);
                const pAbaixo = Math.max(0, 100 - (pMeta + p75));

                const corNota = c.mediaGeral >= 85 ? "text-ra-verdeEscuro" : (c.mediaGeral >= 75 ? "text-ra-roxo" : "text-ra-vermelho");
                const textoBaseHC = c.hcElegivelMes > 0
                    ? `Elegíveis: ${{c.hcElegivelMes}} | Sem Mon.: ${{c.qtdSemMonitoria}}`
                    : `Agentes únicos avaliados`;

                const htmlItensNCG = c.topNCG.length > 0
                    ? c.topNCG.map(item => `
                        <div class="flex items-start justify-between gap-2 text-[11px] bg-ra-vermelho/10 text-ra-vermelho px-2.5 py-1 rounded-lg border border-ra-vermelho/20">
                            <span class="font-bold line-clamp-1" title="[NCG] ${{item.pergunta.replace(/"/g, '&quot;')}}">• [NCG] ${{item.pergunta}}</span>
                            <span class="font-extrabold shrink-0 bg-ra-vermelho text-white px-1.5 rounded text-[10px]">${{item.qtd}}x</span>
                        </div>
                    `).join("")
                    : "";

                const htmlItensCond = c.topConducao.length > 0
                    ? c.topConducao.map(item => `
                        <div class="flex items-start justify-between gap-2 text-[11px] bg-ra-offwhite text-ra-grafite px-2.5 py-1 rounded-lg border border-ra-bege/80">
                            <span class="font-medium line-clamp-1" title="[Condução] ${{item.pergunta.replace(/"/g, '&quot;')}}">• ${{item.pergunta}}</span>
                            <span class="font-extrabold shrink-0 bg-ra-roxo text-white px-1.5 rounded text-[10px]">${{item.qtd}}x</span>
                        </div>
                    `).join("")
                    : `<p class="text-[11px] text-ra-verdeEscuro font-semibold py-1">Nenhum desvio de Condução no recorte.</p>`;

                return `
                <div onclick="selecionarCoordenadorCard('${{c.coordenador.replace(/'/g, "\\\\'")}}')"
                     class="card-coord cursor-pointer p-5 rounded-2xl shadow-sm border border-ra-bege flex flex-col justify-between ${{classeAtivo}}">
                    <div>
                        <div class="flex items-start justify-between gap-2">
                            <div>
                                <span class="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-md bg-ra-lavanda/40 text-ra-roxoEscuro">
                                    Gerente: ${{c.gerente}}
                                </span>
                                <h3 class="text-base font-extrabold text-ra-roxo mt-1.5 leading-snug">${{c.coordenador}}</h3>
                            </div>
                            <div class="text-right shrink-0">
                                <span class="text-[10px] font-bold uppercase text-ra-grafite/60 block">Nota Média</span>
                                <span class="text-2xl font-extrabold ${{corNota}}">${{mediaStr}}</span>
                            </div>
                        </div>

                        <div class="grid grid-cols-3 gap-2 my-3 p-2.5 rounded-xl bg-ra-offwhite text-center">
                            <div>
                                <span class="text-[10px] text-ra-grafite/70 block">Monitorias</span>
                                <strong class="text-xs font-extrabold text-ra-grafite">${{c.qtdMonitorias}}</strong>
                                <span class="text-[10px] text-ra-grafite/60 block">${{c.qtdVoz}}V / ${{c.qtdTexto}}T</span>
                            </div>
                            <div class="border-x border-ra-bege px-1">
                                <span class="text-[10px] text-ra-grafite/70 block">HC Avaliado (&gt;30d)</span>
                                <strong class="text-xs font-extrabold text-ra-roxo">${{c.hcAvaliado}} ag.</strong>
                                <span class="text-[9px] text-ra-grafite/80 font-semibold block leading-tight">${{textoBaseHC}}</span>
                            </div>
                            <div>
                                <span class="text-[10px] text-ra-grafite/70 block">Qual. vs Oper.</span>
                                <strong class="text-xs font-extrabold text-ra-roxoEscuro">${{qualStr}} | ${{operStr}}</strong>
                                <span class="text-[10px] text-ra-grafite/60 block">Calibração</span>
                            </div>
                        </div>

                        <div class="mb-3.5 pt-2.5 border-t border-ra-bege">
                            <div class="flex items-center justify-between mb-1.5">
                                <span class="text-[10px] font-extrabold uppercase tracking-wider text-ra-roxoEscuro">
                                    Principais Ofensores da Coordenação
                                </span>
                                <div class="flex items-center gap-1">
                                    <span class="text-[10px] font-bold px-1.5 py-0.2 rounded bg-ra-lavanda/40 text-ra-roxo">
                                        Cond.: ${{c.desviosConducao}}
                                    </span>
                                    <span class="text-[10px] font-bold px-1.5 py-0.2 rounded ${{c.ocorrenciasNCG > 0 ? 'bg-ra-vermelho text-white' : 'bg-ra-offwhite text-ra-grafite/60'}}">
                                        NCG: ${{c.ocorrenciasNCG}}
                                    </span>
                                </div>
                            </div>
                            <div class="space-y-1">
                                ${{htmlItensNCG}}
                                ${{htmlItensCond}}
                            </div>
                        </div>
                    </div>

                    <div>
                        <div class="flex justify-between text-[10px] font-bold mb-1">
                            <span class="text-ra-verdeEscuro">&ge;85 (Meta): ${{c.acimaMeta}} (${{pMeta}}%)</span>
                            <span class="text-ra-roxo">75-84: ${{c.acima75}} (${{p75}}%)</span>
                            <span class="text-ra-vermelho">&lt;75: ${{c.abaixo75}}</span>
                        </div>
                        <div class="w-full h-2.5 rounded-full bg-ra-offwhite overflow-hidden flex">
                            <div style="width: ${{pMeta}}%; background-color: #84BD00;" title="Acima da Meta (>= 85): ${{pMeta}}%"></div>
                            <div style="width: ${{p75}}%; background-color: #6D2077;" title="Acima de 75 (75 a 84,9): ${{p75}}%"></div>
                            <div style="width: ${{pAbaixo}}%; background-color: #A11430;" title="Abaixo de 75 (< 75): ${{pAbaixo}}%"></div>
                        </div>
                    </div>
                </div>`;
            }}).join("");
        }}

        function atualizarTabelaCoordenacao(lista, listaOfensores) {{
            const arr = consolidarPorCoordenador(lista, listaOfensores);
            const tbody = document.getElementById("tbody-coordenacao");

            if (arr.length === 0) {{
                tbody.innerHTML = `<tr><td colspan="13" class="py-6 text-center text-ra-grafite/60">Nenhum dado encontrado para os filtros selecionados.</td></tr>`;
                return;
            }}

            tbody.innerHTML = arr.map(item => {{
                const mediaStr = item.mediaGeral.toFixed(1).replace(".", ",");
                const qualStr = item.mediaQual !== null ? item.mediaQual.toFixed(1).replace(".", ",") : "-";
                const operStr = item.mediaOper !== null ? item.mediaOper.toFixed(1).replace(".", ",") : "-";
                const cobStr = item.hcElegivelMes > 0
                    ? `${{item.hcAvaliado}} / ${{item.hcElegivelMes}} (${{item.percCobertura}}%)`
                    : `${{item.hcAvaliado}} ag.`;

                return `<tr class="hover:bg-ra-lavanda/15 transition">
                    <td class="py-3 px-4 font-bold text-ra-roxo">${{item.coordenador}}</td>
                    <td class="py-3 px-3 text-ra-grafite/80 font-medium">${{item.gerente}}</td>
                    <td class="py-3 px-3 text-center font-semibold">${{item.qtdMonitorias}} <span class="text-[11px] text-ra-grafite/60 font-normal">(${{item.qtdVoz}}V / ${{item.qtdTexto}}T)</span></td>
                    <td class="py-3 px-3 text-center font-semibold">${{cobStr}}</td>
                    <td class="py-3 px-3 text-center">
                        <span class="px-2 py-0.5 rounded-md font-bold ${{item.qtdSemMonitoria > 0 ? 'bg-ra-amarelo/50 text-ra-roxoEscuro' : 'bg-ra-verde/20 text-ra-verdeEscuro'}}">${{item.qtdSemMonitoria}} ag.</span>
                    </td>
                    <td class="py-3 px-3 text-center">
                        <span class="px-2.5 py-1 rounded-lg font-extrabold bg-ra-roxo text-white">${{mediaStr}}</span>
                    </td>
                    <td class="py-3 px-3 text-center font-semibold text-ra-roxoEscuro">${{qualStr}} | ${{operStr}}</td>
                    <td class="py-3 px-3 text-center font-bold text-ra-verdeEscuro">${{item.acimaMeta}}</td>
                    <td class="py-3 px-3 text-center font-semibold text-ra-roxo">${{item.acima75}}</td>
                    <td class="py-3 px-3 text-center font-bold text-ra-vermelho">${{item.abaixo75}}</td>
                    <td class="py-3 px-3 text-center font-bold text-ra-roxo">${{item.desviosConducao}}</td>
                    <td class="py-3 px-3 text-center">
                        <span class="px-2 py-0.5 rounded-md font-extrabold ${{item.ocorrenciasNCG > 0 ? 'bg-ra-vermelho/15 text-ra-vermelho' : 'text-ra-grafite/50'}}">${{item.ocorrenciasNCG}}</span>
                    </td>
                    <td class="py-3 px-3 text-center font-bold text-ra-roxoEscuro">${{item.bonus}}</td>
                </tr>`;
            }}).join("");
        }}

        function atualizarSecaoOfensores() {{
            const listaOf = ofensoresFiltradosAtuais || [];
            const tipoFiltro = document.getElementById("filtro-tipo-ofensor").value;

            const listaConducao = listaOf.filter(o => o.categoria === "Condução");
            const listaNCG = listaOf.filter(o => o.categoria === "NCG");

            document.getElementById("badge-resumo-conducao").textContent = `Condução (NC): ${{listaConducao.length}} desvios`;
            document.getElementById("badge-resumo-ncg").textContent = `NCG (Grave): ${{listaNCG.length}} ocorrências`;

            const mapCond = {{}};
            listaConducao.forEach(o => {{
                mapCond[o.pergunta] = (mapCond[o.pergunta] || 0) + 1;
            }});
            const arrCond = Object.keys(mapCond)
                .map(p => ({{ pergunta: p, qtd: mapCond[p] }}))
                .sort((a, b) => b.qtd - a.qtd)
                .slice(0, 10);

            const ctxCond = document.getElementById("chartOfensoresConducao").getContext("2d");
            if (chartConducaoInstance) chartConducaoInstance.destroy();
            chartConducaoInstance = new Chart(ctxCond, {{
                type: "bar",
                data: {{
                    labels: arrCond.map(x => x.pergunta.length > 48 ? x.pergunta.slice(0, 48) + "..." : x.pergunta),
                    datasets: [{{
                        label: "Condução (NC)",
                        data: arrCond.map(x => x.qtd),
                        backgroundColor: CORES_RA.roxo,
                        borderRadius: 6
                    }}]
                }},
                options: {{
                    indexAxis: "y",
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {{ padding: {{ right: 55 }} }},
                    scales: {{
                        x: {{ beginAtZero: true, grid: {{ color: "#E8E0D6" }} }},
                        y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10, weight: "600" }} }} }}
                    }},
                    plugins: {{
                        legend: {{ display: false }},
                        datalabels: {{
                            anchor: "end",
                            align: "end",
                            color: CORES_RA.roxoEscuro,
                            font: {{ weight: "bold", size: 11 }},
                            formatter: v => v + "x"
                        }}
                    }}
                }}
            }});

            const mapNCG = {{}};
            listaNCG.forEach(o => {{
                mapNCG[o.pergunta] = (mapNCG[o.pergunta] || 0) + 1;
            }});
            const arrNCG = Object.keys(mapNCG)
                .map(p => ({{ pergunta: p, qtd: mapNCG[p] }}))
                .sort((a, b) => b.qtd - a.qtd)
                .slice(0, 10);

            const ctxNCG = document.getElementById("chartOfensoresNCG").getContext("2d");
            if (chartNCGInstance) chartNCGInstance.destroy();
            chartNCGInstance = new Chart(ctxNCG, {{
                type: "bar",
                data: {{
                    labels: arrNCG.length > 0
                        ? arrNCG.map(x => x.pergunta.length > 48 ? x.pergunta.slice(0, 48) + "..." : x.pergunta)
                        : ["Nenhuma ocorrência de NCG no recorte"],
                    datasets: [{{
                        label: "Ocorrências de NCG",
                        data: arrNCG.length > 0 ? arrNCG.map(x => x.qtd) : [0],
                        backgroundColor: CORES_RA.vermelho,
                        borderRadius: 6
                    }}]
                }},
                options: {{
                    indexAxis: "y",
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {{ padding: {{ right: 55 }} }},
                    scales: {{
                        x: {{ beginAtZero: true, grid: {{ color: "#E8E0D6" }} }},
                        y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10, weight: "600" }} }} }}
                    }},
                    plugins: {{
                        legend: {{ display: false }},
                        datalabels: {{
                            anchor: "end",
                            align: "end",
                            color: CORES_RA.vermelho,
                            font: {{ weight: "bold", size: 11 }},
                            formatter: v => v > 0 ? v + "x" : ""
                        }}
                    }}
                }}
            }});

            const totalMonPorCoord = {{}};
            dadosFiltradosAtuais.forEach(m => {{
                totalMonPorCoord[m.coordenador] = (totalMonPorCoord[m.coordenador] || 0) + 1;
            }});

            const listaTabela = listaOf.filter(o => tipoFiltro === "TODOS" || o.categoria === tipoFiltro);
            const mapTab = {{}};
            listaTabela.forEach(o => {{
                const ch = `${{o.categoria}}|||${{o.grupo}}|||${{o.coordenador}}|||${{o.pergunta}}`;
                if (!mapTab[ch]) {{
                    mapTab[ch] = {{
                        categoria: o.categoria,
                        grupo: o.grupo,
                        coordenador: o.coordenador,
                        pergunta: o.pergunta,
                        qtd: 0,
                        voz: 0,
                        texto: 0
                    }};
                }}
                mapTab[ch].qtd += 1;
                if (o.canal === "Voz") mapTab[ch].voz += 1;
                if (o.canal === "Texto") mapTab[ch].texto += 1;
            }});

            const arrTab = Object.values(mapTab).sort((a, b) => {{
                if (a.categoria !== b.categoria) return a.categoria === "NCG" ? -1 : 1;
                return b.qtd - a.qtd;
            }});

            const tbody = document.getElementById("tbody-ofensores-coord");
            if (arrTab.length === 0) {{
                tbody.innerHTML = `<tr><td colspan="7" class="py-6 text-center text-ra-grafite/60">Nenhum item "Não Conforme" encontrado para os filtros atuais.</td></tr>`;
                return;
            }}

            tbody.innerHTML = arrTab.map(item => {{
                const totMon = totalMonPorCoord[item.coordenador] || 1;
                const percImpacto = Math.round((item.qtd / totMon) * 100);
                const badgeTipo = item.categoria === "NCG"
                    ? `<span class="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold bg-ra-vermelho text-white uppercase">NCG (Zera)</span>`
                    : `<span class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-ra-lavanda/50 text-ra-roxo uppercase">Condução (NC)</span>`;

                return `<tr class="hover:bg-ra-offwhite transition">
                    <td class="py-2.5 px-4 whitespace-nowrap">${{badgeTipo}}</td>
                    <td class="py-2.5 px-4 text-[11px] text-ra-grafite/80">${{item.grupo}}</td>
                    <td class="py-2.5 px-4 font-bold text-ra-roxo">${{item.coordenador}}</td>
                    <td class="py-2.5 px-4 font-medium text-ra-grafite">${{item.pergunta}}</td>
                    <td class="py-2.5 px-4 text-center font-extrabold text-ra-roxoEscuro">${{item.qtd}}</td>
                    <td class="py-2.5 px-4 text-center font-semibold text-ra-grafite">${{percImpacto}}% das mon.</td>
                    <td class="py-2.5 px-4 text-center text-[11px] text-ra-grafite/70">${{item.voz}}V / ${{item.texto}}T</td>
                </tr>`;
            }}).join("");
        }}

        function atualizarGraficos(lista) {{
            const coordArray = consolidarPorCoordenador(lista, ofensoresFiltradosAtuais);

            // 1. GRÁFICO NOTA MÉDIA POR COORDENAÇÃO
            const ctxC = document.getElementById("chartCoordenadores").getContext("2d");
            if (chartCoordenadoresInstance) chartCoordenadoresInstance.destroy();
            chartCoordenadoresInstance = new Chart(ctxC, {{
                type: "bar",
                data: {{
                    labels: coordArray.map(x => x.coordenador),
                    datasets: [
                        {{
                            label: "Nota Geral",
                            data: coordArray.map(x => +x.mediaGeral.toFixed(1)),
                            backgroundColor: CORES_RA.roxo,
                            borderRadius: 6
                        }},
                        {{
                            label: "Nota Qualidade",
                            data: coordArray.map(x => x.mediaQual !== null ? +x.mediaQual.toFixed(1) : 0),
                            backgroundColor: CORES_RA.verde,
                            borderRadius: 6
                        }},
                        {{
                            label: "Nota Operação",
                            data: coordArray.map(x => x.mediaOper !== null ? +x.mediaOper.toFixed(1) : 0),
                            backgroundColor: CORES_RA.lavanda,
                            borderRadius: 6
                        }}
                    ]
                }},
                options: {{
                    indexAxis: "y",
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {{ padding: {{ right: 45 }} }},
                    scales: {{
                        x: {{ beginAtZero: true, max: 110, grid: {{ color: "#E8E0D6" }} }},
                        y: {{ grid: {{ display: false }}, ticks: {{ font: {{ weight: "700", size: 11 }} }} }}
                    }},
                    plugins: {{
                        legend: {{ position: "top", labels: {{ boxWidth: 12, font: {{ size: 11, weight: "600" }} }} }},
                        datalabels: {{
                            anchor: "end",
                            align: "end",
                            color: CORES_RA.roxoEscuro,
                            font: {{ weight: "bold", size: 10 }},
                            formatter: v => v > 0 ? v.toFixed(1).replace(".", ",") : ""
                        }}
                    }}
                }}
            }});

            // 2. GRÁFICO FAIXAS DE META POR COORDENAÇÃO
            const ctxQC = document.getElementById("chartQuartisCoord").getContext("2d");
            if (chartQuartisCoordInstance) chartQuartisCoordInstance.destroy();
            chartQuartisCoordInstance = new Chart(ctxQC, {{
                type: "bar",
                data: {{
                    labels: coordArray.map(x => x.coordenador),
                    datasets: [
                        {{
                            label: "Acima da Meta (>= 85)",
                            data: coordArray.map(x => x.hcAvaliado > 0 ? Math.round((x.acimaMeta / x.hcAvaliado) * 100) : 0),
                            backgroundColor: CORES_RA.verde
                        }},
                        {{
                            label: "Acima de 75 (75 a 84,9)",
                            data: coordArray.map(x => x.hcAvaliado > 0 ? Math.round((x.acima75 / x.hcAvaliado) * 100) : 0),
                            backgroundColor: CORES_RA.roxo
                        }},
                        {{
                            label: "Abaixo de 75 (< 75)",
                            data: coordArray.map(x => x.hcAvaliado > 0 ? Math.round((x.abaixo75 / x.hcAvaliado) * 100) : 0),
                            backgroundColor: CORES_RA.vermelho
                        }}
                    ]
                }},
                options: {{
                    indexAxis: "y",
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {{
                        x: {{ stacked: true, max: 100, grid: {{ color: "#E8E0D6" }}, ticks: {{ callback: v => v + "%" }} }},
                        y: {{ stacked: true, grid: {{ display: false }}, ticks: {{ font: {{ weight: "700", size: 11 }} }} }}
                    }},
                    plugins: {{
                        legend: {{ position: "top", labels: {{ boxWidth: 12, font: {{ size: 11, weight: "600" }} }} }},
                        datalabels: {{
                            color: "#FFFFFF",
                            font: {{ weight: "bold", size: 10 }},
                            formatter: v => v >= 8 ? v + "%" : ""
                        }}
                    }}
                }}
            }});

            // 3. COMPARATIVO VOLUME X NOTA (QUALIDADE VS. OPERAÇÃO)
            const origensAlvo = ["Qualidade", "Operação"];
            const dadosOrig = {{
                "Qualidade": {{ vol: 0, soma: 0, qtdNota: 0 }},
                "Operação": {{ vol: 0, soma: 0, qtdNota: 0 }}
            }};

            lista.forEach(i => {{
                const orig = i.origem === "Qualidade" ? "Qualidade" : "Operação";
                dadosOrig[orig].vol += 1;
                if (i.nota !== null && !isNaN(i.nota)) {{
                    dadosOrig[orig].soma += i.nota;
                    dadosOrig[orig].qtdNota += 1;
                }}
            }});

            const totalVolOrig = Math.max(dadosOrig["Qualidade"].vol + dadosOrig["Operação"].vol, 1);
            const mediaQualNum = dadosOrig["Qualidade"].qtdNota > 0
                ? +(dadosOrig["Qualidade"].soma / dadosOrig["Qualidade"].qtdNota).toFixed(1)
                : 0;
            const mediaOperNum = dadosOrig["Operação"].qtdNota > 0
                ? +(dadosOrig["Operação"].soma / dadosOrig["Operação"].qtdNota).toFixed(1)
                : 0;

            const badgeGap = document.getElementById("badge-gap-calibracao");
            if (dadosOrig["Qualidade"].qtdNota > 0 && dadosOrig["Operação"].qtdNota > 0) {{
                const gap = +(mediaOperNum - mediaQualNum).toFixed(1);
                const sinal = gap > 0 ? "+" : "";
                badgeGap.textContent = `Gap (Oper. - Qual.): ${{sinal}}${{gap.toFixed(1).replace(".", ",")}} pts`;
            }} else {{
                badgeGap.textContent = "Gap Calibração: N/A";
            }}

            const ctxCO = document.getElementById("chartComparativoOrigem").getContext("2d");
            if (chartComparativoOrigemInstance) chartComparativoOrigemInstance.destroy();
            chartComparativoOrigemInstance = new Chart(ctxCO, {{
                type: "bar",
                data: {{
                    labels: origensAlvo,
                    datasets: [
                        {{
                            type: "line",
                            label: "Nota Média",
                            data: [mediaQualNum, mediaOperNum],
                            borderColor: CORES_RA.roxo,
                            backgroundColor: CORES_RA.roxo,
                            pointBackgroundColor: CORES_RA.verde,
                            pointBorderColor: CORES_RA.roxo,
                            pointRadius: 6,
                            borderWidth: 3,
                            yAxisID: "y1",
                            tension: 0.1,
                            datalabels: {{
                                align: "top",
                                anchor: "end",
                                backgroundColor: CORES_RA.roxo,
                                color: "#FFFFFF",
                                borderRadius: 6,
                                padding: {{ top: 3, bottom: 3, left: 7, right: 7 }},
                                font: {{ weight: "bold", size: 11 }},
                                formatter: v => v > 0 ? v.toFixed(1).replace(".", ",") + " pts" : ""
                            }}
                        }},
                        {{
                            type: "line",
                            label: "Meta (85,0)",
                            data: [85, 85],
                            borderColor: CORES_RA.verdeEscuro,
                            borderDash: [6, 5],
                            borderWidth: 2,
                            pointRadius: 0,
                            fill: false,
                            yAxisID: "y1",
                            datalabels: {{
                                display: ctx => ctx.dataIndex === 1,
                                align: "bottom",
                                anchor: "end",
                                backgroundColor: CORES_RA.verdeEscuro,
                                color: "#FFFFFF",
                                borderRadius: 4,
                                padding: {{ top: 2, bottom: 2, left: 5, right: 5 }},
                                font: {{ weight: "bold", size: 9 }},
                                formatter: () => "Meta: 85,0"
                            }}
                        }},
                        {{
                            type: "bar",
                            label: "Volume de Monitorias",
                            data: [dadosOrig["Qualidade"].vol, dadosOrig["Operação"].vol],
                            backgroundColor: [CORES_RA.tomRoxo, CORES_RA.lavanda],
                            borderRadius: 8,
                            barThickness: 75,
                            yAxisID: "y",
                            datalabels: {{
                                anchor: "center",
                                align: "center",
                                color: CORES_RA.roxoEscuro,
                                font: {{ weight: "bold", size: 11 }},
                                formatter: function(v) {{
                                    if (v <= 0) return "";
                                    const p = Math.round((v / totalVolOrig) * 100);
                                    return v + " mon.\\n(" + p + "%)";
                                }},
                                textAlign: "center"
                            }}
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {{ padding: {{ top: 22 }} }},
                    plugins: {{
                        legend: {{
                            position: "top",
                            labels: {{
                                usePointStyle: true,
                                boxWidth: 28,
                                boxHeight: 14,
                                padding: 14,
                                font: {{ size: 11, weight: "600" }},
                                generateLabels: function(chart) {{
                                    const padrao = Chart.defaults.plugins.legend.labels.generateLabels(chart);
                                    return padrao.map(item => {{
                                        if (item.datasetIndex === 0) item.pointStyle = criarIconeLinhaLegenda();
                                        else if (item.datasetIndex === 1) item.pointStyle = criarIconeMetaTracejadaLegenda();
                                        else item.pointStyle = criarIconeBarraLegenda();
                                        return item;
                                    }});
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            type: "linear",
                            position: "left",
                            beginAtZero: true,
                            grid: {{ color: "#E8E0D6" }},
                            title: {{ display: true, text: "Volume", font: {{ weight: "bold", size: 10 }} }}
                        }},
                        y1: {{
                            type: "linear",
                            position: "right",
                            min: 0,
                            max: 115,
                            grid: {{ drawOnChartArea: false }},
                            title: {{ display: true, text: "Nota", font: {{ weight: "bold", size: 10 }} }}
                        }}
                    }}
                }}
            }});

            // 4. GRÁFICO POR SUPERVISOR
            const supMap = {{}};
            lista.forEach(i => {{
                if (i.nota !== null && !isNaN(i.nota)) {{
                    const s = i.supervisor || "Sem Gestor";
                    if (!supMap[s]) supMap[s] = {{ soma: 0, qtd: 0 }};
                    supMap[s].soma += i.nota;
                    supMap[s].qtd += 1;
                }}
            }});
            const supArray = Object.keys(supMap).map(s => ({{
                supervisor: s,
                media: +(supMap[s].soma / supMap[s].qtd).toFixed(1),
                qtd: supMap[s].qtd
            }})).sort((a, b) => b.media - a.media).slice(0, 12);

            const ctxS = document.getElementById("chartSupervisores").getContext("2d");
            if (chartSupervisoresInstance) chartSupervisoresInstance.destroy();
            chartSupervisoresInstance = new Chart(ctxS, {{
                type: "bar",
                data: {{
                    labels: supArray.map(x => x.supervisor),
                    datasets: [{{
                        label: "Nota Média",
                        data: supArray.map(x => x.media),
                        backgroundColor: CORES_RA.tomRoxo,
                        hoverBackgroundColor: CORES_RA.roxo,
                        borderRadius: 6
                    }}]
                }},
                options: {{
                    indexAxis: "y",
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {{ padding: {{ right: 85 }} }},
                    scales: {{
                        x: {{ beginAtZero: true, max: 105, grid: {{ color: "#E8E0D6" }} }},
                        y: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 10, weight: "600" }} }} }}
                    }},
                    plugins: {{
                        legend: {{ display: false }},
                        datalabels: {{
                            anchor: "end",
                            align: "end",
                            color: CORES_RA.roxoEscuro,
                            font: {{ weight: "bold", size: 10 }},
                            formatter: (v, ctx) => {{
                                const q = supArray[ctx.dataIndex]?.qtd || 0;
                                return v.toFixed(1).replace(".", ",") + " (" + q + " mon.)";
                            }}
                        }}
                    }}
                }}
            }});

            // 5. EVOLUÇÃO SEMANAL COM LINHA TRACEJADA DA META DE 85
            const semMap = {{}};
            lista.forEach(i => {{
                if (i.semana && i.semana !== "N/A") {{
                    if (!semMap[i.semana]) semMap[i.semana] = {{ soma: 0, qtdNota: 0, vol: 0 }};
                    semMap[i.semana].vol += 1;
                    if (i.nota !== null && !isNaN(i.nota)) {{
                        semMap[i.semana].soma += i.nota;
                        semMap[i.semana].qtdNota += 1;
                    }}
                }}
            }});
            const semanasOrd = Object.keys(semMap).sort();

            const ctxE = document.getElementById("chartEvolucao").getContext("2d");
            if (chartEvolucaoInstance) chartEvolucaoInstance.destroy();
            chartEvolucaoInstance = new Chart(ctxE, {{
                type: "bar",
                data: {{
                    labels: semanasOrd,
                    datasets: [
                        {{
                            type: "line",
                            label: "Nota Média",
                            data: semanasOrd.map(s => semMap[s].qtdNota > 0 ? +(semMap[s].soma / semMap[s].qtdNota).toFixed(1) : 0),
                            borderColor: CORES_RA.roxo,
                            backgroundColor: CORES_RA.roxo,
                            pointBackgroundColor: CORES_RA.verde,
                            pointBorderColor: CORES_RA.roxo,
                            pointRadius: 5,
                            borderWidth: 3,
                            yAxisID: "y1",
                            tension: 0.3,
                            datalabels: {{
                                align: "top",
                                anchor: "end",
                                backgroundColor: CORES_RA.roxo,
                                color: "#FFFFFF",
                                borderRadius: 4,
                                padding: {{ top: 2, bottom: 2, left: 5, right: 5 }},
                                font: {{ weight: "bold", size: 10 }},
                                formatter: v => v.toFixed(1).replace(".", ",")
                            }}
                        }},
                        {{
                            type: "line",
                            label: "Meta (85,0)",
                            data: semanasOrd.map(() => 85),
                            borderColor: CORES_RA.verdeEscuro,
                            borderDash: [6, 5],
                            borderWidth: 2,
                            pointRadius: 0,
                            fill: false,
                            yAxisID: "y1",
                            datalabels: {{
                                display: ctx => semanasOrd.length > 0 && ctx.dataIndex === semanasOrd.length - 1,
                                align: "bottom",
                                anchor: "end",
                                backgroundColor: CORES_RA.verdeEscuro,
                                color: "#FFFFFF",
                                borderRadius: 4,
                                padding: {{ top: 2, bottom: 2, left: 5, right: 5 }},
                                font: {{ weight: "bold", size: 9 }},
                                formatter: () => "Meta: 85,0"
                            }}
                        }},
                        {{
                            type: "bar",
                            label: "Volume de Monitorias",
                            data: semanasOrd.map(s => semMap[s].vol),
                            backgroundColor: CORES_RA.lavanda,
                            borderRadius: 6,
                            yAxisID: "y",
                            datalabels: {{
                                anchor: "center",
                                align: "center",
                                color: CORES_RA.roxoEscuro,
                                font: {{ weight: "bold", size: 11 }},
                                formatter: v => v > 0 ? v : ""
                            }}
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    layout: {{ padding: {{ top: 20 }} }},
                    plugins: {{
                        legend: {{
                            position: "top",
                            labels: {{
                                usePointStyle: true,
                                boxWidth: 28,
                                boxHeight: 14,
                                padding: 14,
                                font: {{ size: 11, weight: "600" }},
                                generateLabels: function(chart) {{
                                    const padrao = Chart.defaults.plugins.legend.labels.generateLabels(chart);
                                    return padrao.map(item => {{
                                        if (item.datasetIndex === 0) {{
                                            item.pointStyle = criarIconeLinhaLegenda();
                                        }} else if (item.datasetIndex === 1) {{
                                            item.pointStyle = criarIconeMetaTracejadaLegenda();
                                        }} else {{
                                            item.pointStyle = criarIconeBarraLegenda();
                                        }}
                                        return item;
                                    }});
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{ type: "linear", position: "left", beginAtZero: true, grid: {{ color: "#E8E0D6" }} }},
                        y1: {{ type: "linear", position: "right", min: 0, max: 115, grid: {{ drawOnChartArea: false }} }}
                    }}
                }}
            }});
        }}

        function badgeQuartil(q) {{
            if (ehAcimaMeta(q)) return `<span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-ra-verde/25 text-ra-verdeEscuro">${{q}}</span>`;
            if (ehAcima75(q)) return `<span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-ra-lavanda/50 text-ra-roxo">${{q}}</span>`;
            if (ehAbaixo75(q)) return `<span class="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-ra-vermelho/15 text-ra-vermelho">${{q}}</span>`;
            return `<span class="px-2.5 py-0.5 rounded-full text-[11px] bg-ra-offwhite text-ra-grafite">${{q}}</span>`;
        }}

        function atualizarTabelaCriticos(lista) {{
            const agMap = {{}};
            lista.forEach(i => {{
                if (i.nota !== null && !isNaN(i.nota) && ehAbaixo75(i.quartil)) {{
                    const ch = i.login_unico && i.login_unico !== "N/A" ? i.login_unico : i.nome_agente;
                    if (!agMap[ch]) {{
                        agMap[ch] = {{
                            coordenador: i.coordenador,
                            supervisor: i.supervisor,
                            nome: i.nome_agente,
                            quartil: i.quartil,
                            soma: 0,
                            qtd: 0
                        }};
                    }}
                    agMap[ch].soma += i.nota;
                    agMap[ch].qtd += 1;
                }}
            }});

            const arrCriticos = Object.values(agMap)
                .map(a => ({{ ...a, media: a.soma / a.qtd }}))
                .sort((a, b) => a.media - b.media);

            document.getElementById("badge-qtd-criticos").textContent = `${{arrCriticos.length}} colaboradores abaixo de 75`;

            const tbody = document.getElementById("tbody-criticos");
            if (arrCriticos.length === 0) {{
                tbody.innerHTML = `<tr><td colspan="6" class="py-6 text-center text-ra-verdeEscuro font-semibold">Excelente! Nenhum colaborador abaixo de 75 pontos no recorte selecionado.</td></tr>`;
                return;
            }}

            tbody.innerHTML = arrCriticos.map(a => `
                <tr class="hover:bg-ra-vermelho/5 transition">
                    <td class="py-2.5 px-4 font-bold text-ra-roxo">${{a.coordenador}}</td>
                    <td class="py-2.5 px-4 text-ra-grafite">${{a.supervisor}}</td>
                    <td class="py-2.5 px-4 font-semibold text-ra-grafite">${{a.nome}}</td>
                    <td class="py-2.5 px-4 text-center font-semibold">${{a.qtd}}</td>
                    <td class="py-2.5 px-4 text-center font-extrabold text-ra-vermelho">${{a.media.toFixed(1).replace(".", ",")}}</td>
                    <td class="py-2.5 px-4 text-center">${{badgeQuartil(a.quartil)}}</td>
                </tr>
            `).join("");
        }}

        function atualizarTabelaNominal(lista) {{
            const tbody = document.getElementById("tbody-monitorias");
            const limite = 500;
            const fatia = lista.slice(0, limite);

            tbody.innerHTML = fatia.map(item => {{
                const notaStr = (item.nota !== null && !isNaN(item.nota)) ? item.nota.toFixed(1).replace(".", ",") : "N/A";
                const corOrigem = item.origem === "Qualidade" ? "bg-ra-roxo text-white" : "bg-ra-bege text-ra-grafite";
                const corBonus = item.status_bonus === "Bônus +6 Aplicado"
                    ? "bg-ra-verde/25 text-ra-verdeEscuro font-bold"
                    : "text-ra-grafite/60";
                const textoVal = item.motivo_expurgo ? `${{item.validacao}} (${{item.motivo_expurgo}})` : item.validacao;

                return `<tr class="hover:bg-ra-offwhite transition">
                    <td class="py-2.5 px-3 whitespace-nowrap font-mono">${{item.data_avaliacao}}</td>
                    <td class="py-2.5 px-3 whitespace-nowrap font-mono text-[11px] text-ra-grafite/70">${{item.id_interacao}}</td>
                    <td class="py-2.5 px-3 text-ra-grafite/80">${{item.gerente}}</td>
                    <td class="py-2.5 px-3 font-bold text-ra-roxo">${{item.coordenador}}</td>
                    <td class="py-2.5 px-3 text-ra-grafite">${{item.supervisor}}</td>
                    <td class="py-2.5 px-3 font-medium text-ra-grafite">${{item.nome_agente}}</td>
                    <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded bg-ra-offwhite text-[10px] font-bold text-ra-roxoEscuro">${{item.alocacao}}</span></td>
                    <td class="py-2.5 px-3 text-center font-semibold text-ra-roxoEscuro">${{item.canal}}</td>
                    <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded-lg text-[11px] font-semibold ${{corOrigem}}">${{item.origem}}</span></td>
                    <td class="py-2.5 px-3 text-center font-extrabold text-ra-roxo">${{notaStr}}</td>
                    <td class="py-2.5 px-3 text-center">${{badgeQuartil(item.quartil)}}</td>
                    <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded-lg text-[11px] ${{corBonus}}">${{item.status_bonus}}</span></td>
                    <td class="py-2.5 px-3 text-[11px] text-ra-grafite/75">${{textoVal}}</td>
                </tr>`;
            }}).join("");

            document.getElementById("lbl-contagem-tabela").textContent =
                `Exibindo ${{fatia.length}} de ${{lista.length.toLocaleString("pt-BR")}} monitorias filtradas`;
        }}

        // =========================================================================
        // MOTOR DE GERAÇÃO DE IMAGEM EXECUTIVA ORGANIZADA COM OS GRÁFICOS (CANVAS 2D)
        // =========================================================================

        function desenharRetanguloArredondado(ctx, x, y, w, h, r, fill, stroke) {{
            ctx.beginPath();
            ctx.moveTo(x + r, y);
            ctx.lineTo(x + w - r, y);
            ctx.quadraticCurveTo(x + w, y, x + w, y + r);
            ctx.lineTo(x + w, y + h - r);
            ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
            ctx.lineTo(x + r, y + h);
            ctx.quadraticCurveTo(x, y + h, x, y + h - r);
            ctx.lineTo(x, y + r);
            ctx.quadraticCurveTo(x, y, x + r, y);
            ctx.closePath();
            if (fill) {{
                ctx.fillStyle = fill;
                ctx.fill();
            }}
            if (stroke) {{
                ctx.strokeStyle = stroke;
                ctx.lineWidth = 1.5;
                ctx.stroke();
            }}
        }}

        function truncarTextoCanvas(ctx, texto, maxWidth) {{
            const str = String(texto || "");
            if (ctx.measureText(str).width <= maxWidth) return str;
            let s = str;
            while (s.length > 0 && ctx.measureText(s + "...").width > maxWidth) {{
                s = s.slice(0, -1);
            }}
            return s + "...";
        }}

        function comporPainelExecutivoEmCanvas() {{
            const lista = dadosFiltradosAtuais || [];
            const listaOf = ofensoresFiltradosAtuais || [];
            const coords = consolidarPorCoordenador(lista, listaOf);

            const gerentesFiltro = FILTROS_MULTI.gerente.size > 0
                ? Array.from(FILTROS_MULTI.gerente).join(" / ")
                : "Visão Consolidada (Todos os Gerentes)";

            const mesesFiltro = FILTROS_MULTI.mes.size > 0
                ? Array.from(FILTROS_MULTI.mes).join(", ")
                : (DADOS_RELATORIO.mes_mais_recente || "Período Atual");

            const diasFiltro = FILTROS_MULTI.data.size > 0
                ? ` | Dias selecionados: ${{FILTROS_MULTI.data.size}}`
                : "";

            ultimoAssuntoEmail = `[Qualidade CX - Rede Américas] Painel Executivo por Coordenação`;

            const escala = 1.5;
            const W = 1520;
            const pad = 28;
            const colGap = 20;

            const qtdLinhasCoords = Math.ceil(Math.max(coords.length, 1) / 3);
            const alturaCardCoord = 175;
            const alturaSecaoCoords = qtdLinhasCoords * (alturaCardCoord + 16) + 42;

            const alturaCardGrafico = 370;
            const alturaTotal = 105 + 115 + alturaSecaoCoords + (4 * (alturaCardGrafico + 22)) + 45;

            const canvas = document.getElementById("canvas-exportacao-email");
            canvas.width = W * escala;
            canvas.height = alturaTotal * escala;

            const ctx = canvas.getContext("2d");
            ctx.scale(escala, escala);

            // 1. Fundo Off White Oficial (#F1F1EF)
            ctx.fillStyle = CORES_RA.offwhite;
            ctx.fillRect(0, 0, W, alturaTotal);

            // 2. Barra Superior de Gradiente Oficial Rede Américas
            const grad = ctx.createLinearGradient(0, 0, W, 0);
            grad.addColorStop(0, CORES_RA.roxo);
            grad.addColorStop(0.5, CORES_RA.lavanda);
            grad.addColorStop(1, CORES_RA.verde);
            ctx.fillStyle = grad;
            ctx.fillRect(0, 0, W, 8);

            // 3. Cabeçalho Institucional Roxo Púrpura (#6D2077)
            ctx.fillStyle = CORES_RA.roxo;
            ctx.fillRect(0, 8, W, 82);

            ctx.fillStyle = "#FFFFFF";
            ctx.font = "800 22px Inter, Arial, sans-serif";
            ctx.fillText("Rede Américas • Book Executivo de Qualidade CX", pad, 42);

            ctx.fillStyle = CORES_RA.lavanda;
            ctx.font = "500 13px Inter, Arial, sans-serif";
            ctx.fillText(`Gerência: ${{gerentesFiltro}}   |   Período: ${{mesesFiltro}}${{diasFiltro}}`, pad, 68);

            ctx.fillStyle = CORES_RA.verde;
            ctx.font = "700 12px Inter, Arial, sans-serif";
            ctx.textAlign = "right";
            ctx.fillText(`Atualizado em: ${{DADOS_RELATORIO.gerado_em || "-"}}`, W - pad, 55);
            ctx.textAlign = "left";

            // 4. Linha de 5 KPIs Executivos
            let yCursor = 108;
            const kpis = [
                {{
                    titulo: "NOTA MÉDIA GERAL",
                    valor: document.getElementById("kpi-nota-media").textContent,
                    sub: document.getElementById("kpi-sub-calibracao").textContent,
                    corTopo: CORES_RA.roxo,
                    corValor: CORES_RA.roxo
                }},
                {{
                    titulo: "MONITORIAS VÁLIDAS",
                    valor: document.getElementById("kpi-total-validas").textContent,
                    sub: document.getElementById("kpi-sub-canal").textContent,
                    corTopo: CORES_RA.verde,
                    corValor: CORES_RA.grafite
                }},
                {{
                    titulo: "HC AVALIADO (>30 DIAS)",
                    valor: document.getElementById("kpi-agentes-avaliados").textContent,
                    sub: document.getElementById("kpi-sub-cobertura").textContent,
                    corTopo: CORES_RA.tomRoxo,
                    corValor: CORES_RA.roxoEscuro
                }},
                {{
                    titulo: "ACIMA DA META (>= 85)",
                    valor: document.getElementById("kpi-perc-meta").textContent,
                    sub: document.getElementById("kpi-sub-faixas").textContent,
                    corTopo: CORES_RA.verdeEscuro,
                    corValor: CORES_RA.verdeEscuro
                }},
                {{
                    titulo: "ITENS NÃO CONFORMES",
                    valor: document.getElementById("kpi-total-ofensores").textContent,
                    sub: document.getElementById("kpi-sub-ofensores").textContent,
                    corTopo: CORES_RA.vermelho,
                    corValor: CORES_RA.vermelho
                }}
            ];

            const kpiW = (W - (pad * 2) - (colGap * 4)) / 5;
            const kpiH = 94;

            kpis.forEach((k, idx) => {{
                const kx = pad + idx * (kpiW + colGap);
                desenharRetanguloArredondado(ctx, kx, yCursor, kpiW, kpiH, 14, "#FFFFFF", CORES_RA.bege);

                ctx.fillStyle = k.corTopo;
                ctx.fillRect(kx + 10, yCursor, kpiW - 20, 4);

                ctx.fillStyle = "#666666";
                ctx.font = "700 10px Inter, Arial, sans-serif";
                ctx.fillText(k.titulo, kx + 16, yCursor + 24);

                ctx.fillStyle = k.corValor;
                ctx.font = "800 28px Inter, Arial, sans-serif";
                ctx.fillText(k.valor, kx + 16, yCursor + 58);

                ctx.fillStyle = CORES_RA.grafite;
                ctx.font = "600 11px Inter, Arial, sans-serif";
                ctx.fillText(truncarTextoCanvas(ctx, k.sub, kpiW - 28), kx + 16, yCursor + 80);
            }});

            yCursor += kpiH + 24;

            // 5. Seção de Scorecards Visuais dos Coordenadores
            ctx.fillStyle = CORES_RA.roxo;
            ctx.font = "800 16px Inter, Arial, sans-serif";
            ctx.fillText("1. Scorecards de Performance por Coordenação (Meta 85 / 75 & Ofensores)", pad, yCursor + 14);
            yCursor += 28;

            const cardW = (W - (pad * 2) - (colGap * 2)) / 3;

            coords.forEach((c, idx) => {{
                const colIdx = idx % 3;
                const rowIdx = Math.floor(idx / 3);
                const cx = pad + colIdx * (cardW + colGap);
                const cy = yCursor + rowIdx * (alturaCardCoord + 16);

                desenharRetanguloArredondado(ctx, cx, cy, cardW, alturaCardCoord, 14, "#FFFFFF", CORES_RA.bege);

                ctx.fillStyle = CORES_RA.roxoEscuro;
                ctx.font = "700 10px Inter, Arial, sans-serif";
                ctx.fillText(truncarTextoCanvas(ctx, `GERENTE: ${{c.gerente}}`, cardW - 110), cx + 16, cy + 22);

                ctx.fillStyle = CORES_RA.roxo;
                ctx.font = "800 15px Inter, Arial, sans-serif";
                ctx.fillText(truncarTextoCanvas(ctx, c.coordenador, cardW - 110), cx + 16, cy + 44);

                const corNota = c.mediaGeral >= 85 ? CORES_RA.verdeEscuro : (c.mediaGeral >= 75 ? CORES_RA.roxo : CORES_RA.vermelho);
                ctx.textAlign = "right";
                ctx.fillStyle = "#777777";
                ctx.font = "700 9px Inter, Arial, sans-serif";
                ctx.fillText("NOTA MÉDIA", cx + cardW - 16, cy + 20);
                ctx.fillStyle = corNota;
                ctx.font = "800 24px Inter, Arial, sans-serif";
                ctx.fillText(c.mediaGeral.toFixed(1).replace(".", ","), cx + cardW - 16, cy + 44);
                ctx.textAlign = "left";

                desenharRetanguloArredondado(ctx, cx + 14, cy + 54, cardW - 28, 36, 8, CORES_RA.offwhite, null);
                const qualStr = c.mediaQual !== null ? c.mediaQual.toFixed(1).replace(".", ",") : "-";
                const operStr = c.mediaOper !== null ? c.mediaOper.toFixed(1).replace(".", ",") : "-";

                ctx.fillStyle = CORES_RA.grafite;
                ctx.font = "700 11px Inter, Arial, sans-serif";
                ctx.fillText(`Mon.: ${{c.qtdMonitorias}} (${{c.qtdVoz}}V/${{c.qtdTexto}}T)`, cx + 24, cy + 76);
                ctx.fillText(`HC Aval.: ${{c.hcAvaliado}} (Pend.: ${{c.qtdSemMonitoria}})`, cx + 165, cy + 76);
                ctx.fillStyle = CORES_RA.roxoEscuro;
                ctx.fillText(`Qual/Oper: ${{qualStr}} | ${{operStr}}`, cx + 325, cy + 76);

                ctx.fillStyle = CORES_RA.roxoEscuro;
                ctx.font = "700 10px Inter, Arial, sans-serif";
                ctx.fillText(`Ofensores -> Condução: ${{c.desviosConducao}}  |  NCG: ${{c.ocorrenciasNCG}}`, cx + 16, cy + 108);

                const topOfTexto = c.topNCG.length > 0
                    ? `[NCG] ${{c.topNCG[0].pergunta}} (${{c.topNCG[0].qtd}}x)`
                    : (c.topConducao.length > 0 ? `Top Cond.: ${{c.topConducao[0].pergunta}} (${{c.topConducao[0].qtd}}x)` : "Sem ofensores no recorte");

                ctx.fillStyle = c.topNCG.length > 0 ? CORES_RA.vermelho : CORES_RA.grafite;
                ctx.font = "600 10px Inter, Arial, sans-serif";
                ctx.fillText(truncarTextoCanvas(ctx, topOfTexto, cardW - 32), cx + 16, cy + 125);

                const totalAg = Math.max(c.hcAvaliado, 1);
                const pMeta = Math.round((c.acimaMeta / totalAg) * 100);
                const p75 = Math.round((c.acima75 / totalAg) * 100);
                const pAbaixo = Math.max(0, 100 - (pMeta + p75));

                ctx.font = "700 10px Inter, Arial, sans-serif";
                ctx.fillStyle = CORES_RA.verdeEscuro;
                ctx.fillText(`>=85 (Meta): ${{c.acimaMeta}} (${{pMeta}}%)`, cx + 16, cy + 148);

                ctx.fillStyle = CORES_RA.roxo;
                ctx.fillText(`75-84: ${{c.acima75}} (${{p75}}%)`, cx + 185, cy + 148);

                ctx.fillStyle = CORES_RA.vermelho;
                ctx.fillText(`<75: ${{c.abaixo75}}`, cx + 335, cy + 148);

                const barX = cx + 16;
                const barY = cy + 155;
                const barW = cardW - 32;
                const barH = 8;
                desenharRetanguloArredondado(ctx, barX, barY, barW, barH, 4, CORES_RA.offwhite, null);

                const w1 = (pMeta / 100) * barW;
                const w2 = (p75 / 100) * barW;
                const w3 = Math.max(0, barW - w1 - w2);

                if (w1 > 0) {{ ctx.fillStyle = CORES_RA.verde; ctx.fillRect(barX, barY, w1, barH); }}
                if (w2 > 0) {{ ctx.fillStyle = CORES_RA.roxo; ctx.fillRect(barX + w1, barY, w2, barH); }}
                if (w3 > 0) {{ ctx.fillStyle = CORES_RA.vermelho; ctx.fillRect(barX + w1 + w2, barY, w3, barH); }}
            }});

            yCursor += qtdLinhasCoords * (alturaCardCoord + 16) + 12;

            // 6. Desenha os 7 Gráficos Organizados
            const graficosConfig = [
                {{ id: "chartCoordenadores", titulo: "Nota Média por Coordenação (Geral vs. Qualidade vs. Operação)" }},
                {{ id: "chartQuartisCoord", titulo: "Faixas de Meta por Coordenação (>=85, 75 a 84,9 e <75)" }},
                {{ id: "chartComparativoOrigem", titulo: "Comparativo Volume x Nota: Qualidade vs. Operação" }},
                {{ id: "chartEvolucao", titulo: "Evolução Semanal (Volume de Monitorias vs. Nota Média)" }},
                {{ id: "chartOfensoresConducao", titulo: "Top 10 Ofensores — Condução do Atendimento (NC)" }},
                {{ id: "chartOfensoresNCG", titulo: "Incidências de NCG — Não Conformidade Grave (Zeramento)" }},
                {{ id: "chartSupervisores", titulo: "Desempenho por Supervisor (Top 12 no Recorte)", larguraCompleta: true }}
            ];

            const chartCardW = (W - (pad * 2) - colGap) / 2;
            const fullChartW = W - (pad * 2);

            graficosConfig.forEach((g, idx) => {{
                const colIdx = idx % 2;
                const rowIdx = Math.floor(idx / 2);
                const larguraCard = g.larguraCompleta ? fullChartW : chartCardW;
                const gx = g.larguraCompleta ? pad : (pad + colIdx * (chartCardW + colGap));
                const gy = yCursor + rowIdx * (alturaCardGrafico + 22);

                desenharRetanguloArredondado(ctx, gx, gy, larguraCard, alturaCardGrafico, 14, "#FFFFFF", CORES_RA.bege);

                ctx.fillStyle = g.id === "chartOfensoresNCG" ? CORES_RA.vermelho : CORES_RA.roxo;
                ctx.font = "800 14px Inter, Arial, sans-serif";
                ctx.fillText(g.titulo, gx + 18, gy + 28);

                const canvasOrig = document.getElementById(g.id);
                if (canvasOrig && canvasOrig.width > 0 && canvasOrig.height > 0) {{
                    const areaW = larguraCard - 36;
                    const areaH = alturaCardGrafico - 52;
                    ctx.drawImage(canvasOrig, gx + 18, gy + 40, areaW, areaH);
                }}
            }});

            ctx.fillStyle = "#777777";
            ctx.font = "600 11px Inter, Arial, sans-serif";
            ctx.textAlign = "center";
            ctx.fillText("Rede Américas • Paixão por Cuidar • Gestão Nacional de Qualidade CX", W / 2, alturaTotal - 14);
            ctx.textAlign = "left";

            return canvas;
        }}

        function prepararImagemGraficosEEnviarOutlook() {{
            const canvas = comporPainelExecutivoEmCanvas();
            const dataUrl = canvas.toDataURL("image/png");
            document.getElementById("img-preview-painel-executivo").src = dataUrl;
            document.getElementById("modal-email-outlook").classList.remove("hidden");
            lucide.createIcons();

            canvas.toBlob(async function(blob) {{
                ultimoBlobImagemCompleta = blob;
                await copiarImagemCompletaParaClipboard();
                acionarMailtoOutlook();
            }}, "image/png");
        }}

        async function copiarImagemCompletaParaClipboard() {{
            const btn = document.getElementById("btn-copiar-img-completa");
            const canvas = document.getElementById("canvas-exportacao-email");

            try {{
                const blob = ultimoBlobImagemCompleta || await new Promise(res => canvas.toBlob(res, "image/png"));
                if (navigator.clipboard && window.ClipboardItem && blob) {{
                    await navigator.clipboard.write([
                        new ClipboardItem({{ "image/png": blob }})
                    ]);
                    if (btn) {{
                        const orig = btn.innerHTML;
                        btn.innerHTML = `<i data-lucide="check" class="w-4 h-4 text-ra-verde"></i><span>Imagem Copiada! Aperte Ctrl+V no Outlook</span>`;
                        lucide.createIcons();
                        setTimeout(() => {{ btn.innerHTML = orig; lucide.createIcons(); }}, 3500);
                    }}
                }}
            }} catch (err) {{
                console.warn("Aviso ao copiar imagem via Clipboard API:", err);
            }}
        }}

        async function copiarGraficoIndividual(idCanvas, tituloGrafico, btnEl) {{
            const origCanvas = document.getElementById(idCanvas);
            if (!origCanvas) return;

            const tempCanvas = document.createElement("canvas");
            const pad = 24;
            tempCanvas.width = origCanvas.width + pad * 2;
            tempCanvas.height = origCanvas.height + pad * 2 + 36;
            const tCtx = tempCanvas.getContext("2d");

            tCtx.fillStyle = "#FFFFFF";
            tCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);

            tCtx.fillStyle = "#6D2077";
            tCtx.font = "bold 18px Inter, Arial, sans-serif";
            tCtx.fillText(`Rede Américas • ${{tituloGrafico}}`, pad, pad + 18);

            tCtx.drawImage(origCanvas, pad, pad + 32);

            tempCanvas.toBlob(async function(blob) {{
                try {{
                    if (navigator.clipboard && window.ClipboardItem && blob) {{
                        await navigator.clipboard.write([
                            new ClipboardItem({{ "image/png": blob }})
                        ]);
                        const txtOrig = btnEl.textContent;
                        btnEl.textContent = "✓ Copiado (Ctrl+V)";
                        btnEl.classList.add("bg-ra-verde", "text-ra-roxoEscuro");
                        setTimeout(() => {{
                            btnEl.textContent = txtOrig;
                            btnEl.classList.remove("bg-ra-verde", "text-ra-roxoEscuro");
                        }}, 2500);
                    }}
                }} catch (e) {{
                    alert("Clique com o botão direito sobre a imagem abaixo e selecione 'Copiar imagem'.");
                }}
            }}, "image/png");
        }}

        function baixarImagemCompletaPNG() {{
            const canvas = document.getElementById("canvas-exportacao-email");
            const link = document.createElement("a");
            link.download = `Painel_Graficos_Qualidade_RedeAmericas_${{new Date().toISOString().slice(0,10)}}.png`;
            link.href = canvas.toDataURL("image/png");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }}

        function acionarMailtoOutlook() {{
            const assuntoEnc = encodeURIComponent(ultimoAssuntoEmail);
            const instrucaoCorpo = encodeURIComponent(
                "Olá,\\n\\nSegue abaixo o Painel Executivo de Qualidade CX por Coordenação:\\n\\n[PRESSIONE CTRL + V AQUI PARA COLAR A IMAGEM ORGANIZADA DOS GRÁFICOS]\\n\\nAtenciosamente,\\nQualidade CX - Rede Américas"
            );
            window.location.href = `mailto:?subject=${{assuntoEnc}}&body=${{instrucaoCorpo}}`;
        }}

        function fecharModalEmail() {{
            document.getElementById("modal-email-outlook").classList.add("hidden");
        }}

        function exportarCSV() {{
            if (!dadosFiltradosAtuais || dadosFiltradosAtuais.length === 0) {{
                alert("Não há dados filtrados para exportar.");
                return;
            }}

            const cabecalhos = [
                "Data_Avaliacao", "Mes_Ano", "ID_Interacao", "ID_Agente_Chave",
                "Gerente", "Coordenador", "Supervisor", "Colaborador", "Cargo", "Matricula", "Alocacao",
                "Canal", "Formulario", "Avaliador", "Origem", "Nota", "Faixa_Meta",
                "Status_Bonus", "Validacao", "Motivo_Expurgo", "Dias_de_Casa"
            ];

            const linhas = dadosFiltradosAtuais.map(i => [
                i.data_avaliacao,
                i.mes_ano,
                i.id_interacao,
                i.id_chave,
                `"${{(i.gerente || "").replace(/"/g, '""')}}"`,
                `"${{(i.coordenador || "").replace(/"/g, '""')}}"`,
                `"${{(i.supervisor || "").replace(/"/g, '""')}}"`,
                `"${{(i.nome_agente || "").replace(/"/g, '""')}}"`,
                `"${{(i.cargo || "N/A").replace(/"/g, '""')}}"`,
                i.matricula,
                `"${{(i.alocacao || "").replace(/"/g, '""')}}"`,
                i.canal,
                `"${{(i.formulario || "").replace(/"/g, '""')}}"`,
                `"${{(i.avaliador || "").replace(/"/g, '""')}}"`,
                i.origem,
                (i.nota !== null && !isNaN(i.nota)) ? String(i.nota).replace(".", ",") : "N/A",
                `"${{(i.quartil || "").replace(/"/g, '""')}}"`,
                `"${{(i.status_bonus || "").replace(/"/g, '""')}}"`,
                `"${{(i.validacao || "").replace(/"/g, '""')}}"`,
                `"${{(i.motivo_expurgo || "").replace(/"/g, '""')}}"`,
                i.dias_casa
            ].join(";"));

            const conteudoCSV = "\\uFEFF" + [cabecalhos.join(";"), ...linhas].join("\\n");
            const blob = new Blob([conteudoCSV], {{ type: "text/csv;charset=utf-8;" }});
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `Book_Gerencial_Qualidade_Coordenadores_${{new Date().toISOString().slice(0,10)}}.csv`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }}

        window.addEventListener("DOMContentLoaded", inicializarPainel);
    </script>
</body>
</html>
"""

        with open(self.caminho_saida, "w", encoding="utf-8") as f:
            f.write(html_template)

        print(f"\n[SUCESSO] Book Gerencial HTML gerado em:\n -> {self.caminho_saida}")
        return self.caminho_saida


# ================================================================
# EXECUÇÃO PARA GERAR O RELATÓRIO HTML E ABRIR NO NAVEGADOR
# ================================================================
if __name__ == "__main__":
    print("Iniciando processamento do ETL de Qualidade Rede Américas...")
    etl_qualidade = PowerBIQualidadeETL()
    df_fatos_qualidade, df_dim_colaboradores, df_fatos_csat = etl_qualidade.executar_pipeline()

    gerador_html = GeradorRelatorioHTML(etl_qualidade)
    caminho_arquivo_html = gerador_html.gerar_html()

    try:
        webbrowser.open(f"file:///{os.path.abspath(caminho_arquivo_html)}")
    except Exception as e:
        print(f"Não foi possível abrir o navegador automaticamente: {e}")