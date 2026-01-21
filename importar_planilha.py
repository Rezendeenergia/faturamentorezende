import pandas as pd
from datetime import datetime
from database import Database
from excel_handler import ExcelHandler


class PlanilhaImporter:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.db = Database()

    def importar_tudo(self):
        """Importa todas as NFs e Extrato da planilha"""
        print("=" * 60)
        print("IMPORTANDO PLANILHA EXISTENTE")
        print("=" * 60)

        # Importa NFs
        print("\n1️⃣ Importando Notas Fiscais...")
        nfs_importadas = self.importar_notas_fiscais()
        print(f"   ✅ {nfs_importadas} notas fiscais importadas")

        # Importa Extrato
        print("\n2️⃣ Importando Extrato...")
        extratos_importados = self.importar_extrato()
        print(f"   ✅ {extratos_importados} lançamentos de extrato importados")

        # Dashboard
        print("\n3️⃣ Gerando Dashboard...")
        dashboard = self.db.dashboard_recebimentos()

        print(f"\n{'=' * 60}")
        print("RESUMO DA IMPORTAÇÃO")
        print(f"{'=' * 60}")
        print(f"📊 Total de NFs: {nfs_importadas}")
        print(f"💰 Total de Recebimentos: {extratos_importados}")
        print(f"\n🟢 RECEBIDO: {dashboard['recebido']['qtd']} NFs - R$ {dashboard['recebido']['total']:,.2f}")
        print(f"🟡 A RECEBER: {dashboard['a_receber']['qtd']} NFs - R$ {dashboard['a_receber']['total']:,.2f}")
        print(f"🔴 ATRASADO: {dashboard['atrasado']['qtd']} NFs - R$ {dashboard['atrasado']['total']:,.2f}")
        print(f"{'=' * 60}\n")

        return {
            'nfs': nfs_importadas,
            'extrato': extratos_importados,
            'dashboard': dashboard
        }

    def importar_notas_fiscais(self):
        """Importa todas as notas fiscais da aba NF'S"""
        df = pd.read_excel(self.excel_path, sheet_name="NF'S")

        contador = 0

        for idx, row in df.iterrows():
            try:
                # Pula linhas vazias
                if pd.isna(row['Nº NF']):
                    continue

                # Converte data
                data_emissao = self._converter_data(row['Data Emissão'])
                if not data_emissao:
                    continue

                # Dados básicos
                tipo = str(row['Tipo']).strip() if pd.notna(row['Tipo']) else 'CONSTRUCAO'
                valor_bruto = float(row['Valor Bruto']) if pd.notna(row['Valor Bruto']) else 0

                # Retenções - LÊ DIRETO DA PLANILHA
                inss = float(row['Retenções Federais (INSS)']) if pd.notna(row['Retenções Federais (INSS)']) else 0
                iss = float(row['ISS']) if pd.notna(row['ISS']) else 0
                retencao_equatorial = float(row['Retenção Equatorial']) if pd.notna(row['Retenção Equatorial']) else 0

                # PIS/COFINS/CSLL
                pis_cofins_csll = float(row['PIS/COFINS/CSLL']) if pd.notna(row['PIS/COFINS/CSLL']) else 0
                pis_cofins_retido = pis_cofins_csll > 0

                # Valor Nominal - USA O DA PLANILHA (coluna N)
                valor_nominal = float(row['Valor Nominal (Vinci)']) if pd.notna(row['Valor Nominal (Vinci)']) else 0

                # Valor Nominal Conferência (coluna M)
                valor_nominal_conferencia = float(row['Valor Nominal Conferência']) if pd.notna(
                    row['Valor Nominal Conferência']) else None

                # Valor Líquido Vinci (coluna O)
                valor_liquido_vinci = float(row['Valor Líquido Vinci']) if pd.notna(
                    row['Valor Líquido Vinci']) else None

                # Valor Retido Vinci (coluna R) - DIRETO DA PLANILHA
                valor_retido_vinci = float(row['Valor retido Vinci']) if pd.notna(row['Valor retido Vinci']) else None

                # Data do adiantamento (coluna P)
                data_adiantamento = self._converter_data(row['Data do adiantamento']) if pd.notna(
                    row.get('Data do adiantamento')) else None

                # Percentual de Adiantamento (coluna Q)
                percentual_adiantamento = float(row['% de Adiantamento']) * 100 if pd.notna(
                    row.get('% de Adiantamento')) else None

                # Foi adiantado se tem Valor Líquido Vinci OU Valor Retido Vinci
                foi_adiantado = (valor_liquido_vinci is not None and valor_liquido_vinci > 0) or \
                                (valor_retido_vinci is not None and valor_retido_vinci > 0)

                dados = {
                    'data_emissao': data_emissao,
                    'numero_nf': str(row['Nº NF']).strip(),
                    'tipo': tipo,
                    'valor_bruto': valor_bruto,
                    'localidade': str(row.get('Localidade', '')).strip() if pd.notna(row.get('Localidade')) else '',
                    'tomador': str(row.get('Tomador do Serviço', '')).strip() if pd.notna(
                        row.get('Tomador do Serviço')) else '',
                    'inss': inss,
                    'iss': iss,
                    'retencao_equatorial': retencao_equatorial,
                    'pis_cofins_retido': pis_cofins_retido,
                    'pis_cofins_csll': pis_cofins_csll,
                    'valor_nominal_conferencia': valor_nominal_conferencia,
                    'valor_nominal_calculado': valor_nominal,  # Usa o da planilha!
                    'valor_liquido_vinci': valor_liquido_vinci,
                    'foi_adiantado': foi_adiantado,
                    'data_adiantamento': data_adiantamento,
                    'valor_retido_vinci': valor_retido_vinci,
                    'percentual_adiantamento': percentual_adiantamento
                }

                self.db.inserir_nota(dados)
                contador += 1

            except Exception as e:
                print(f"   ⚠️  Erro na linha {idx + 2}: {str(e)}")
                continue

        return contador

    def importar_extrato(self):
        """Importa todos os lançamentos da aba Extrato"""
        df = pd.read_excel(self.excel_path, sheet_name="Extrato")

        contador = 0

        for idx, row in df.iterrows():
            try:
                # Pula linhas vazias
                if pd.isna(row['Data']):
                    continue

                data_recebimento = self._converter_data(row['Data'])
                if not data_recebimento:
                    continue

                # Valor
                valor_col = 'Valor            '  # Nome exato da coluna
                valor_recebido = float(row[valor_col]) if pd.notna(row[valor_col]) else 0

                if valor_recebido <= 0:
                    continue

                # NFs referentes
                nfs_col = "NF'S"
                nfs_referentes = str(row[nfs_col]).strip() if pd.notna(row[nfs_col]) else ''

                if not nfs_referentes:
                    continue

                # Tipo
                tipo_recebimento = str(row.get('Tipo', 'Integral')).strip() if pd.notna(row.get('Tipo')) else 'Integral'

                # Complemento
                complemento = str(row.get(
                    'Complemento                                                                                                                                           ',
                    '')).strip() if pd.notna(row.get(
                    'Complemento                                                                                                                                           ')) else ''

                dados = {
                    'data_recebimento': data_recebimento,
                    'valor_recebido': valor_recebido,
                    'nfs_referentes': nfs_referentes,
                    'tipo_recebimento': tipo_recebimento,
                    'complemento': complemento
                }

                self.db.inserir_recebimento(dados)
                contador += 1

            except Exception as e:
                print(f"   ⚠️  Erro na linha {idx + 2}: {str(e)}")
                continue

        return contador

    def _converter_data(self, data):
        """Converte data para formato YYYY-MM-DD"""
        if pd.isna(data):
            return None

        try:
            if isinstance(data, str):
                # Tenta diferentes formatos
                for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                    try:
                        dt = datetime.strptime(data, fmt)
                        return dt.strftime('%Y-%m-%d')
                    except:
                        continue
            elif isinstance(data, datetime):
                return data.strftime('%Y-%m-%d')
            elif hasattr(data, 'to_pydatetime'):
                return data.to_pydatetime().strftime('%Y-%m-%d')
        except:
            pass

        return None


if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Uso: python importar_planilha.py <caminho_planilha>")
        sys.exit(1)

    excel_path = sys.argv[1]
    importer = PlanilhaImporter(excel_path)
    importer.importar_tudo()