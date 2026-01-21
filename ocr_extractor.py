import pdfplumber
import re
from datetime import datetime


class NFExtractor:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.data = {}

    def extract(self):
        """Extrai dados do PDF"""
        with pdfplumber.open(self.pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

        # Identifica o tipo de nota
        self.data['tipo'] = self._identify_type(text)

        # Extrai dados básicos
        self.data['numero_nf'] = self._extract_nf_number(text)
        self.data['data_emissao'] = self._extract_date(text)
        self.data['valor_bruto'] = self._extract_valor_bruto(text)
        self.data['localidade'] = self._extract_localidade(text)
        self.data['tomador'] = self._extract_tomador(text)

        # Extrai retenções
        self.data['inss'] = self._extract_inss(text)
        self.data['iss'] = self._extract_iss(text)

        # Extrai dados específicos por tipo
        if self.data['tipo'] == 'CONSTRUCAO':
            self.data['contrato'] = self._extract_contrato(text)
            self.data['folhas_registro'] = self._extract_folhas(text)
        elif self.data['tipo'] == 'TRANSPORTE':
            self.data['stm'] = self._extract_stm(text)
            self.data['requisicao'] = self._extract_requisicao(text)

        return self.data

    def _identify_type(self, text):
        """Identifica o tipo de nota fiscal"""
        text_upper = text.upper()

        # CT-e tem prioridade
        if 'CT-E' in text_upper or 'DACTE' in text_upper or 'CONHECIMENTO DE TRANSPORTE' in text_upper:
            return 'TRANSPORTE_CTE'
        # Ensaio dielétrico
        elif 'ENSAIO' in text_upper and ('DIELETRIC' in text_upper or 'RIGIDEZ' in text_upper):
            return 'ENSAIO DIELETRICO'
        # Transporte NF
        elif 'TRANSPORTE' in text_upper and (
                'RODOVIARIO' in text_upper or 'MUNICIPAL' in text_upper) and 'CT-E' not in text_upper:
            return 'TRANSPORTE'
        # Construção (PLPT, obras, etc)
        elif 'CONSTRUCAO' in text_upper or 'CONSTRUÇÃO' in text_upper or 'PLPT' in text_upper or 'OBRAS' in text_upper:
            return 'CONSTRUCAO'

        return 'CONSTRUCAO'  # Default

    def _extract_nf_number(self, text):
        """Extrai número da NF"""
        patterns = [
            r'Nº\s*(\d+)',
            r'NÚMERO\s*(\d+)',
            r'NF-e.*?Nº\s*(\d+)',
            r'NOTA FISCAL.*?Nº\s*(\d+)',
            r'CT-E.*?Nº\s*DOCUMENTO:\s*(\d+)',
            r'NÚMERO.*?(\d+).*?SÉRIE'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group(1)
        return ""

    def _extract_date(self, text):
        """Extrai data de emissão"""
        patterns = [
            r'emitida em:\s*(\d{2}/\d{2}/\d{4})',
            r'Data emissão\s*(\d{2}/\d{2}/\d{4})',
            r'DATA E HORA DE EMISSÃO\s*(\d{2}/\d{2}/\d{4})',
            r'(\d{2}/\d{2}/\d{4})\s*\d{2}:\d{2}:\d{2}'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1)
        return ""

    def _extract_valor_bruto(self, text):
        """Extrai valor bruto da nota"""
        patterns = [
            r'Valor dos serviços\s*R?\$?\s*([\d.,]+)',
            r'Valor da nota\s*R?\$?\s*([\d.,]+)',
            r'VALOR TOTAL DO SERVIÇO\s*R?\$?\s*([\d.,]+)',
            r'VALOR TOTAL A RECEBER\s*R?\$?\s*([\d.,]+)',
            r'FRETE.*?VALOR.*?R?\$?\s*([\d.,]+)'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                valor = match.group(1).replace('.', '').replace(',', '.')
                try:
                    return float(valor)
                except:
                    continue
        return 0.0

    def _extract_localidade(self, text):
        """Extrai município/localidade"""
        patterns = [
            r'MUNICIPIO[:\s]+([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ\s]+)',
            r'Serviço prestado em\s*PA-([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ\s]+)',
            r'MUNICÍPIO:\s*([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ\s]+)'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    def _extract_tomador(self, text):
        """Extrai tomador do serviço"""
        patterns = [
            r'TOMADOR DE SERVIÇOS.*?Nome/Razão:\s*([^\n]+)',
            r'CENTRAIS ELETRICAS DO PARA',
            r'EQUATORIAL PARA',
            r'CONECTA EMPREENDIMENTOS'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                tomador = match.group(0) if 'CENTRAIS' in match.group(0) or 'EQUATORIAL' in match.group(
                    0) or 'CONECTA' in match.group(0) else match.group(1)
                if 'CELPA' in tomador.upper() or 'CENTRAIS ELETRICAS' in tomador.upper():
                    return 'CELPA'
                elif 'EQUATORIAL' in tomador.upper():
                    return 'EQUATORIAL'
                elif 'CONECTA' in tomador.upper():
                    return 'CONECTA'
                return tomador.strip()
        return ""

    def _extract_inss(self, text):
        """Extrai valor do INSS"""
        patterns = [
            r'INSS\s*R?\$?\s*([\d.,]+)',
            r'RETENÇÃO INSS[:\s]+R?\$?\s*([\d.,]+)'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                valor = match.group(1).replace('.', '').replace(',', '.')
                return float(valor)
        return 0.0

    def _extract_iss(self, text):
        """Extrai valor do ISS"""
        patterns = [
            r'Valor do imposto\(ISS\)\s*R?\$?\s*([\d.,]+)',
            r'VALOR DO ISS\s*([\d.,]+)',
            r'ISS Retido.*?R?\$?\s*([\d.,]+)'
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                valor = match.group(1).replace('.', '').replace(',', '.')
                return float(valor)
        return 0.0

    def _extract_contrato(self, text):
        """Extrai número do contrato"""
        match = re.search(r'CONTRATO N\.?º?\s*(\d+/\d+)', text, re.IGNORECASE)
        return match.group(1) if match else ""

    def _extract_folhas(self, text):
        """Extrai folhas de registro"""
        folhas = []
        pattern = r'FOLHA DE REGISTRO.*?(\d{10})'
        matches = re.findall(pattern, text, re.IGNORECASE)
        return ', '.join(matches) if matches else ""

    def _extract_stm(self, text):
        """Extrai número STM"""
        match = re.search(r'STM\s*(\d+)', text, re.IGNORECASE)
        return match.group(1) if match else ""

    def _extract_requisicao(self, text):
        """Extrai número da requisição"""
        match = re.search(r'REQUISIÇÃO[:\s]+(\d+)', text, re.IGNORECASE)
        return match.group(1) if match else ""