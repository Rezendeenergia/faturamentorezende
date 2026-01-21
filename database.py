import sqlite3
from datetime import datetime, timedelta
import json


class Database:
    def __init__(self, db_path='sistema_nf.db'):
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Inicializa o banco de dados"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Tabela de Notas Fiscais
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS notas_fiscais (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_emissao TEXT NOT NULL,
                numero_nf TEXT NOT NULL UNIQUE,
                tipo TEXT NOT NULL,
                valor_bruto REAL NOT NULL,
                localidade TEXT,
                tomador TEXT,
                inss REAL,
                iss REAL,
                retencao_equatorial REAL,
                pis_cofins_retido INTEGER DEFAULT 0,
                pis_cofins_csll REAL,
                valor_nominal_conferencia REAL,
                valor_nominal_calculado REAL,
                valor_liquido_vinci REAL,
                foi_adiantado INTEGER DEFAULT 0,
                data_adiantamento TEXT,
                percentual_adiantamento REAL,
                valor_retido_vinci REAL,
                data_vencimento TEXT,
                dias_para_receber INTEGER,
                status_recebimento TEXT DEFAULT 'PENDENTE',
                criado_em TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Tabela de Extrato (Recebimentos)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS extrato (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_recebimento TEXT NOT NULL,
                valor_recebido REAL NOT NULL,
                nfs_referentes TEXT NOT NULL,
                tipo_recebimento TEXT NOT NULL,
                complemento TEXT,
                foi_adiantado INTEGER DEFAULT 0,
                criado_em TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Tabela de Conciliação (relaciona NFs com Recebimentos)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS conciliacao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nota_fiscal_id INTEGER NOT NULL,
                extrato_id INTEGER NOT NULL,
                valor_conciliado REAL NOT NULL,
                tipo_recebimento TEXT NOT NULL,
                FOREIGN KEY (nota_fiscal_id) REFERENCES notas_fiscais(id),
                FOREIGN KEY (extrato_id) REFERENCES extrato(id)
            )
        ''')

        conn.commit()
        conn.close()

    def calcular_prazo_recebimento(self, tipo, data_emissao):
        """Calcula data de vencimento baseado no tipo"""
        prazos = {
            'TRANSPORTE': 60,
            'TRANSPORTE_CTE': 60,
            'ENSAIO DIELETRICO': 30,
            'CONSTRUCAO': 30
        }

        dias = prazos.get(tipo, 30)
        data = datetime.strptime(data_emissao, '%Y-%m-%d')
        data_vencimento = data + timedelta(days=dias)

        return data_vencimento.strftime('%Y-%m-%d'), dias

    def inserir_nota(self, dados):
        """Insere uma nota fiscal no banco"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Calcula prazo de recebimento
        data_vencimento, dias = self.calcular_prazo_recebimento(
            dados['tipo'],
            dados['data_emissao']
        )

        cursor.execute('''
            INSERT INTO notas_fiscais (
                data_emissao, numero_nf, tipo, valor_bruto, localidade, tomador,
                inss, iss, retencao_equatorial, pis_cofins_retido, pis_cofins_csll,
                valor_nominal_conferencia, valor_nominal_calculado, valor_liquido_vinci,
                foi_adiantado, data_adiantamento, data_vencimento, dias_para_receber,
                valor_retido_vinci, percentual_adiantamento
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            dados['data_emissao'],
            dados['numero_nf'],
            dados['tipo'],
            dados['valor_bruto'],
            dados.get('localidade'),
            dados.get('tomador'),
            dados.get('inss', 0),
            dados.get('iss', 0),
            dados.get('retencao_equatorial', 0),
            1 if dados.get('pis_cofins_retido') else 0,
            dados.get('pis_cofins_csll', 0),
            dados.get('valor_nominal_conferencia'),
            dados.get('valor_nominal_calculado'),
            dados.get('valor_liquido_vinci'),
            1 if dados.get('foi_adiantado') else 0,
            dados.get('data_adiantamento'),
            data_vencimento,
            dias,
            dados.get('valor_retido_vinci'),
            dados.get('percentual_adiantamento')
        ))

        nf_id = cursor.lastrowid
        conn.commit()
        conn.close()

        return nf_id

    def inserir_recebimento(self, dados):
        """Insere um recebimento no extrato"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT INTO extrato (
                data_recebimento, valor_recebido, nfs_referentes, 
                tipo_recebimento, complemento, foi_adiantado
            ) VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            dados['data_recebimento'],
            dados['valor_recebido'],
            dados['nfs_referentes'],
            dados['tipo_recebimento'],
            dados.get('complemento', ''),
            dados.get('foi_adiantado', 0)
        ))

        extrato_id = cursor.lastrowid
        conn.commit()

        # Concilia com as NFs
        self._conciliar_recebimento(cursor, extrato_id, dados)

        conn.commit()
        conn.close()

        return extrato_id

    def _conciliar_recebimento(self, cursor, extrato_id, dados):
        """Concilia recebimento com notas fiscais"""
        nfs_str = dados['nfs_referentes'].strip()

        # Se não tem NF ou está vazio, registra recebimento sem conciliar
        if not nfs_str or nfs_str == '':
            print(f"   ℹ️  Recebimento sem NF específica - R$ {dados['valor_recebido']:.2f}")
            return

        nfs = [nf.strip() for nf in nfs_str.split(',')]

        # Para cada NF mencionada
        for nf in nfs:
            # Normaliza o número da NF
            nf_normalizado = str(nf).replace('.0', '').strip()

            # Busca a NF e seu valor líquido esperado
            cursor.execute('''
                SELECT 
                    id, 
                    valor_liquido_vinci, 
                    valor_nominal_conferencia, 
                    valor_nominal_calculado,
                    valor_bruto
                FROM notas_fiscais 
                WHERE TRIM(REPLACE(numero_nf, '.0', '')) = ?
                OR numero_nf = ?
            ''', (nf_normalizado, nf_normalizado))
            result = cursor.fetchone()

            if result:
                nf_id = result[0]
                valor_liquido_vinci = result[1]
                valor_nominal_conferencia = result[2]
                valor_nominal_calculado = result[3]
                valor_bruto = result[4]

                # HIERARQUIA CORRIGIDA: Ignora valores NULL e zerados
                if valor_nominal_conferencia and valor_nominal_conferencia > 0:
                    valor_esperado = valor_nominal_conferencia
                elif valor_liquido_vinci and valor_liquido_vinci > 0:
                    valor_esperado = valor_liquido_vinci
                elif valor_nominal_calculado and valor_nominal_calculado > 0:
                    valor_esperado = valor_nominal_calculado
                else:
                    valor_esperado = valor_bruto

                # Registra o recebimento com o valor esperado dessa NF
                cursor.execute('''
                    INSERT INTO conciliacao (
                        nota_fiscal_id, extrato_id, valor_conciliado, tipo_recebimento
                    ) VALUES (?, ?, ?, ?)
                ''', (nf_id, extrato_id, valor_esperado, dados['tipo_recebimento']))

                print(f"   ✅ NF {nf_normalizado} conciliada - {dados['tipo_recebimento']} - R$ {valor_esperado:.2f}")

                # Atualiza status da NF
                self._atualizar_status_nf(cursor, nf_id)
            else:
                print(f"   ⚠️  NF {nf_normalizado} não encontrada no banco")

    def _atualizar_status_nf(self, cursor, nf_id):
        """Atualiza status de recebimento da NF"""
        # Soma total recebido
        cursor.execute('''
            SELECT SUM(valor_conciliado) 
            FROM conciliacao 
            WHERE nota_fiscal_id = ?
        ''', (nf_id,))

        total_recebido = cursor.fetchone()[0] or 0

        # Busca valor esperado (HIERARQUIA CORRIGIDA)
        cursor.execute('''
            SELECT 
                valor_nominal_conferencia, 
                valor_liquido_vinci, 
                valor_nominal_calculado,
                valor_bruto
            FROM notas_fiscais 
            WHERE id = ?
        ''', (nf_id,))

        result = cursor.fetchone()
        valor_nominal_conferencia = result[0]
        valor_liquido_vinci = result[1]
        valor_nominal_calculado = result[2]
        valor_bruto = result[3]

        # Ignora valores NULL e zerados
        if valor_nominal_conferencia and valor_nominal_conferencia > 0:
            valor_esperado = valor_nominal_conferencia
        elif valor_liquido_vinci and valor_liquido_vinci > 0:
            valor_esperado = valor_liquido_vinci
        elif valor_nominal_calculado and valor_nominal_calculado > 0:
            valor_esperado = valor_nominal_calculado
        else:
            valor_esperado = valor_bruto

        # Define status
        if total_recebido >= valor_esperado:
            status = 'RECEBIDO'
        elif total_recebido > 0:
            status = 'PARCIAL'
        else:
            status = 'PENDENTE'

        # Atualiza
        cursor.execute('''
            UPDATE notas_fiscais 
            SET status_recebimento = ?
            WHERE id = ?
        ''', (status, nf_id))

    def listar_pendentes(self):
        """Lista NFs pendentes de recebimento"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        hoje = datetime.now().strftime('%Y-%m-%d')

        cursor.execute('''
            SELECT 
                id, numero_nf, data_emissao, tipo, valor_bruto,
                -- HIERARQUIA CORRIGIDA: Ignora NULL e valores zerados
                CASE 
                    WHEN valor_nominal_conferencia IS NOT NULL AND valor_nominal_conferencia > 0 
                        THEN valor_nominal_conferencia
                    WHEN valor_liquido_vinci IS NOT NULL AND valor_liquido_vinci > 0 
                        THEN valor_liquido_vinci
                    WHEN valor_nominal_calculado IS NOT NULL AND valor_nominal_calculado > 0 
                        THEN valor_nominal_calculado
                    ELSE valor_bruto
                END as valor_liquido,
                data_vencimento, tomador, localidade,
                status_recebimento,
                CASE 
                    WHEN date(data_vencimento) < date(?) THEN 'ATRASADO'
                    ELSE 'A_RECEBER'
                END as situacao,
                CAST(julianday(?) - julianday(data_vencimento) as INTEGER) as dias_diferenca
            FROM notas_fiscais
            WHERE status_recebimento != 'RECEBIDO'
            ORDER BY data_vencimento ASC
        ''', (hoje, hoje))

        notas = [dict(row) for row in cursor.fetchall()]
        conn.close()

        return notas

    def dashboard_recebimentos(self):
        """Retorna dados para dashboard de recebimentos"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        hoje = datetime.now().strftime('%Y-%m-%d')

        # Total A Receber - HIERARQUIA CORRIGIDA
        cursor.execute('''
            SELECT 
                COUNT(*) as qtd,
                SUM(
                    CASE 
                        WHEN valor_nominal_conferencia IS NOT NULL AND valor_nominal_conferencia > 0 
                            THEN valor_nominal_conferencia
                        WHEN valor_liquido_vinci IS NOT NULL AND valor_liquido_vinci > 0 
                            THEN valor_liquido_vinci
                        WHEN valor_nominal_calculado IS NOT NULL AND valor_nominal_calculado > 0 
                            THEN valor_nominal_calculado
                        ELSE valor_bruto
                    END
                ) as total
            FROM notas_fiscais
            WHERE status_recebimento != 'RECEBIDO'
            AND date(data_vencimento) >= date(?)
        ''', (hoje,))

        a_receber = dict(cursor.fetchone())

        # Total Atrasado - HIERARQUIA CORRIGIDA
        cursor.execute('''
            SELECT 
                COUNT(*) as qtd,
                SUM(
                    CASE 
                        WHEN valor_nominal_conferencia IS NOT NULL AND valor_nominal_conferencia > 0 
                            THEN valor_nominal_conferencia
                        WHEN valor_liquido_vinci IS NOT NULL AND valor_liquido_vinci > 0 
                            THEN valor_liquido_vinci
                        WHEN valor_nominal_calculado IS NOT NULL AND valor_nominal_calculado > 0 
                            THEN valor_nominal_calculado
                        ELSE valor_bruto
                    END
                ) as total
            FROM notas_fiscais
            WHERE status_recebimento != 'RECEBIDO'
            AND date(data_vencimento) < date(?)
        ''', (hoje,))

        atrasado = dict(cursor.fetchone())

        # Total Recebido - HIERARQUIA CORRIGIDA
        cursor.execute('''
            SELECT 
                COUNT(*) as qtd,
                SUM(
                    CASE 
                        WHEN valor_nominal_conferencia IS NOT NULL AND valor_nominal_conferencia > 0 
                            THEN valor_nominal_conferencia
                        WHEN valor_liquido_vinci IS NOT NULL AND valor_liquido_vinci > 0 
                            THEN valor_liquido_vinci
                        WHEN valor_nominal_calculado IS NOT NULL AND valor_nominal_calculado > 0 
                            THEN valor_nominal_calculado
                        ELSE valor_bruto
                    END
                ) as total
            FROM notas_fiscais
            WHERE status_recebimento = 'RECEBIDO'
        ''')

        recebido = dict(cursor.fetchone())

        conn.close()

        return {
            'a_receber': a_receber,
            'atrasado': atrasado,
            'recebido': recebido
        }

    def adiantar_nota(self, nota_id, dados):
        """Registra adiantamento de uma nota fiscal E cria lançamento no extrato"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Busca dados da nota
        cursor.execute('''
            SELECT numero_nf, valor_nominal_calculado, pis_cofins_csll, valor_nominal_conferencia
            FROM notas_fiscais 
            WHERE id = ?
        ''', (nota_id,))

        resultado = cursor.fetchone()
        if not resultado:
            conn.close()
            raise Exception("Nota fiscal não encontrada")

        numero_nf = resultado[0]
        valor_nominal = resultado[1]
        pis_cofins_csll = resultado[2]
        valor_nominal_conferencia = resultado[3]

        # Calcula valores
        valor_liquido = dados['valor_liquido_vinci']

        # Se PIS retido, desconta do nominal
        if dados['pis_cofins_retido']:
            valor_nominal_final = valor_nominal - pis_cofins_csll
        else:
            valor_nominal_final = valor_nominal

        valor_retido = valor_nominal_final - valor_liquido
        percentual_adiantamento = (valor_retido / valor_nominal_final) * 100 if valor_nominal_final > 0 else 0

        # Atualiza nota
        cursor.execute('''
            UPDATE notas_fiscais 
            SET 
                pis_cofins_retido = ?,
                foi_adiantado = 1,
                data_adiantamento = ?,
                valor_liquido_vinci = ?,
                percentual_adiantamento = ?,
                valor_retido_vinci = ?,
                valor_nominal_conferencia = ?
            WHERE id = ?
        ''', (
            1 if dados['pis_cofins_retido'] else 0,
            dados['data_adiantamento'],
            valor_liquido,
            percentual_adiantamento,
            valor_retido,
            valor_liquido,  # Atualiza o Valor Nominal Conferência
            nota_id
        ))

        # NOVO: Cria lançamento automático no extrato
        cursor.execute('''
            INSERT INTO extrato (
                data_recebimento, valor_recebido, nfs_referentes, 
                tipo_recebimento, complemento, foi_adiantado
            ) VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            dados['data_adiantamento'],
            valor_liquido,
            numero_nf,
            'Adiantamento',
            f'Adiantamento automático - {percentual_adiantamento:.1f}% de taxa',
            1
        ))

        extrato_id = cursor.lastrowid

        # Concilia automaticamente
        cursor.execute('''
            INSERT INTO conciliacao (
                nota_fiscal_id, extrato_id, valor_conciliado, tipo_recebimento
            ) VALUES (?, ?, ?, ?)
        ''', (nota_id, extrato_id, valor_liquido, 'Adiantamento'))

        # Atualiza status da NF
        self._atualizar_status_nf(cursor, nota_id)

        conn.commit()
        conn.close()

        return {
            'valor_retido': valor_retido,
            'percentual': percentual_adiantamento,
            'extrato_id': extrato_id
        }

    def analise_financeira(self):
        """Retorna análise financeira completa"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Total de juros de adiantamento (coluna R - Valor Retido Vinci)
        cursor.execute('''
            SELECT SUM(valor_retido_vinci) as total_juros
            FROM notas_fiscais
            WHERE foi_adiantado = 1
        ''')
        juros = cursor.fetchone()['total_juros'] or 0

        # Total de Retenção Equatorial (coluna J)
        cursor.execute('''
            SELECT SUM(retencao_equatorial) as total_retencao
            FROM notas_fiscais
        ''')
        retencao_equatorial = cursor.fetchone()['total_retencao'] or 0

        # Total de ISS retido (coluna H)
        cursor.execute('''
            SELECT SUM(iss) as total_iss
            FROM notas_fiscais
        ''')
        iss = cursor.fetchone()['total_iss'] or 0

        # Total de INSS retido (coluna F)
        cursor.execute('''
            SELECT SUM(inss) as total_inss
            FROM notas_fiscais
        ''')
        inss = cursor.fetchone()['total_inss'] or 0

        conn.close()

        return {
            'juros': juros,
            'retencao_equatorial': retencao_equatorial,
            'iss': iss,
            'inss': inss
        }

    def listar_todas_notas(self):
        """Lista todas as notas fiscais"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute('''
            SELECT *,
                CASE 
                    WHEN valor_nominal_conferencia IS NOT NULL AND valor_nominal_conferencia > 0 
                        THEN valor_nominal_conferencia
                    WHEN valor_liquido_vinci IS NOT NULL AND valor_liquido_vinci > 0 
                        THEN valor_liquido_vinci
                    WHEN valor_nominal_calculado IS NOT NULL AND valor_nominal_calculado > 0 
                        THEN valor_nominal_calculado
                    ELSE valor_bruto
                END as valor_liquido_exibicao
            FROM notas_fiscais
            ORDER BY data_emissao DESC
        ''')

        notas = [dict(row) for row in cursor.fetchall()]
        conn.close()

        return notas

    def listar_extrato(self, filtro_adiantamento=None):
        """Lista todos os lançamentos do extrato com filtro opcional"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        query = '''
            SELECT 
                id,
                data_recebimento,
                valor_recebido,
                nfs_referentes,
                tipo_recebimento,
                complemento,
                foi_adiantado,
                criado_em
            FROM extrato
        '''

        params = []

        # Aplica filtro se fornecido
        if filtro_adiantamento is not None:
            query += ' WHERE foi_adiantado = ?'
            params.append(1 if filtro_adiantamento else 0)

        query += ' ORDER BY data_recebimento DESC'

        cursor.execute(query, params)
        extrato = [dict(row) for row in cursor.fetchall()]
        conn.close()

        return extrato

    def exportar_para_excel(self, tipo_relatorio):
        """Exporta relatório para formato Excel (dados em dict)"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        if tipo_relatorio == 'todas_notas':
            cursor.execute('''
                SELECT 
                    numero_nf as "Nº NF",
                    data_emissao as "Data Emissão",
                    tipo as "Tipo",
                    valor_bruto as "Valor Bruto",
                    localidade as "Localidade",
                    tomador as "Tomador",
                    inss as "INSS",
                    iss as "ISS",
                    retencao_equatorial as "Retenção Equatorial",
                    pis_cofins_csll as "PIS/COFINS/CSLL",
                    COALESCE(valor_nominal_conferencia, valor_nominal_calculado) as "Valor Nominal",
                    valor_liquido_vinci as "Valor Líquido Vinci",
                    data_vencimento as "Data Vencimento",
                    status_recebimento as "Status"
                FROM notas_fiscais
                ORDER BY data_emissao DESC
            ''')

        elif tipo_relatorio == 'pendentes':
            hoje = datetime.now().strftime('%Y-%m-%d')
            cursor.execute('''
                SELECT 
                    numero_nf as "Nº NF",
                    data_emissao as "Data Emissão",
                    tipo as "Tipo",
                    COALESCE(valor_nominal_conferencia, valor_nominal_calculado) as "Valor a Receber",
                    data_vencimento as "Data Vencimento",
                    CASE 
                        WHEN date(data_vencimento) < date(?) THEN 'ATRASADO'
                        ELSE 'A RECEBER'
                    END as "Situação",
                    CAST(julianday(?) - julianday(data_vencimento) as INTEGER) as "Dias",
                    tomador as "Tomador",
                    localidade as "Localidade"
                FROM notas_fiscais
                WHERE status_recebimento != 'RECEBIDO'
                ORDER BY data_vencimento ASC
            ''', (hoje, hoje))

        elif tipo_relatorio == 'extrato':
            cursor.execute('''
                SELECT 
                    data_recebimento as "Data Recebimento",
                    valor_recebido as "Valor Recebido",
                    nfs_referentes as "NFs",
                    tipo_recebimento as "Tipo",
                    CASE WHEN foi_adiantado = 1 THEN 'SIM' ELSE 'NÃO' END as "Adiantado",
                    complemento as "Complemento"
                FROM extrato
                ORDER BY data_recebimento DESC
            ''')

        dados = [dict(row) for row in cursor.fetchall()]
        conn.close()

        return dados