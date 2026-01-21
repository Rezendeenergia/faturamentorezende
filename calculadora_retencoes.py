"""
Calculadora de Retenções - Independente de arquivo Excel
Calcula INSS, ISS, Retenção Equatorial e PIS/COFINS/CSLL
"""


class CalculadoraRetencoes:
    """Calcula retenções sem depender de arquivo Excel"""

    @staticmethod
    def calcular_retencoes(tipo, valor_bruto, pis_cofins_retido=False):
        """
        Calcula todas as retenções baseado no tipo de serviço

        Args:
            tipo: Tipo de serviço (CONSTRUCAO, ENSAIO DIELETRICO, TRANSPORTE, TRANSPORTE_CTE)
            valor_bruto: Valor bruto da nota fiscal
            pis_cofins_retido: Se PIS/COFINS/CSLL foram retidos

        Returns:
            dict com todas as retenções calculadas
        """
        retencoes = {
            'inss': 0.0,
            'aliquota_inss': 0.0,
            'iss': 0.0,
            'aliquota_iss': 0.0,
            'retencao_equatorial': 0.0,
            'pis_cofins_csll': 0.0
        }

        if tipo == 'CONSTRUCAO':
            # INSS: 11% sobre 50% do valor (mão de obra)
            retencoes['inss'] = valor_bruto * 0.50 * 0.11
            retencoes['aliquota_inss'] = 0.055  # 5.5% sobre valor bruto
            # ISS: 5%
            retencoes['iss'] = valor_bruto * 0.05
            retencoes['aliquota_iss'] = 0.05
            # Retenção Equatorial: 5%
            retencoes['retencao_equatorial'] = valor_bruto * 0.05

        elif tipo == 'ENSAIO DIELETRICO':
            # INSS: 11% sobre valor total
            retencoes['inss'] = valor_bruto * 0.11
            retencoes['aliquota_inss'] = 0.11
            # ISS: 5%
            retencoes['iss'] = valor_bruto * 0.05
            retencoes['aliquota_iss'] = 0.05
            # Retenção Equatorial: 0%
            retencoes['retencao_equatorial'] = 0.0

        elif tipo == 'TRANSPORTE':
            # INSS: 0%
            retencoes['inss'] = 0.0
            retencoes['aliquota_inss'] = 0.0
            # ISS: 5%
            retencoes['iss'] = valor_bruto * 0.05
            retencoes['aliquota_iss'] = 0.05
            # Retenção Equatorial: 3%
            retencoes['retencao_equatorial'] = valor_bruto * 0.03

        elif tipo == 'TRANSPORTE_CTE':
            # CT-e tem regras diferentes
            # INSS: 0%
            retencoes['inss'] = 0.0
            retencoes['aliquota_inss'] = 0.0
            # ISS: 0%
            retencoes['iss'] = 0.0
            retencoes['aliquota_iss'] = 0.0
            # Retenção Equatorial: 3%
            retencoes['retencao_equatorial'] = valor_bruto * 0.03

        # PIS/COFINS/CSLL: 4.65% se retido
        if pis_cofins_retido:
            retencoes['pis_cofins_csll'] = valor_bruto * 0.0465

        return retencoes

    @staticmethod
    def calcular_valor_nominal(valor_bruto, retencoes):
        """
        Calcula valor nominal (valor após retenções)

        Fórmula: Valor Bruto - INSS - ISS - Ret.Equatorial - PIS/COFINS/CSLL

        Args:
            valor_bruto: Valor bruto da nota
            retencoes: Dict com retenções calculadas

        Returns:
            float: Valor nominal calculado
        """
        return (
                valor_bruto
                - retencoes['inss']
                - retencoes['iss']
                - retencoes['retencao_equatorial']
                - retencoes['pis_cofins_csll']
        )

    @staticmethod
    def calcular_completo(tipo, valor_bruto, pis_cofins_retido=False):
        """
        Calcula retenções E valor nominal de uma vez

        Args:
            tipo: Tipo de serviço
            valor_bruto: Valor bruto da NF
            pis_cofins_retido: Se PIS/COFINS foram retidos

        Returns:
            dict: {
                'retencoes': {...},
                'valor_nominal': float,
                'valor_bruto': float
            }
        """
        retencoes = CalculadoraRetencoes.calcular_retencoes(tipo, valor_bruto, pis_cofins_retido)
        valor_nominal = CalculadoraRetencoes.calcular_valor_nominal(valor_bruto, retencoes)

        return {
            'retencoes': retencoes,
            'valor_nominal': round(valor_nominal, 2),
            'valor_bruto': valor_bruto
        }