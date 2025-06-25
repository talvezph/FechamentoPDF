import unittest
import os
import pandas as pd
from unittest.mock import patch, MagicMock
from collections import defaultdict
from datetime import datetime
from script_fechamento import normalize, encontrar_nome_aproximado, extrair_dados_pdf, calcular_fechamento, diarios_info, VALOR_ENTREGA, BONUS_DIARIO, main

class TestFechamentoMotoristas(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Configurações para testes, se necessário
        # Mock da planilha de veículos para testes
        cls.mock_diarios_info = {
            "MOTORISTA A": {"diaria": 100.0, "tipo": "CARRO"},
            "MOTORISTA B": {"diaria": 120.0, "tipo": "MOTO"},
            "MOTORISTA C": {"diaria": 150.0, "tipo": "VAN"},
            "ELISIANE LUDMYLLA FERREIRA SANTOS": {"diaria": 100.0, "tipo": "CARRO"},
        }
        # Substitui a variável global diarios_info pelo mock
        global diarios_info
        diarios_info = cls.mock_diarios_info

    def test_normalize(self):
        self.assertEqual(normalize("Olá Mundo!"), "ola mundo!")
        self.assertEqual(normalize("Árvore Cão"), "arvore cao")
        self.assertEqual(normalize("TESTE COM ACENTOS"), "teste com acentos")
        self.assertEqual(normalize("123 Teste"), "123 teste")

    def test_encontrar_nome_aproximado(self):
        # Teste com cutoff 0.95 para corresponder ao script principal
        self.assertEqual(encontrar_nome_aproximado("MOTORISTA A"), "MOTORISTA A")
        self.assertIsNone(encontrar_nome_aproximado("MOTORISTA Z"))
        self.assertIsNone(encontrar_nome_aproximado("MOTORISTA AAAA"))

    @patch("pdfplumber.open")
    def test_extrair_dados_pdf_cenario1(self, mock_pdfplumber_open):
        # Mock do PDF para o Cenário 1
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_pdf.pages = [mock_page]
        mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

        # Cenário 1: PDF com nome de motorista, entregas e bônus
        mock_page.extract_text.return_value = (
            "Motorista: MOTORISTA A\n" +
            "01/01/2023 - Entrega - Sim\n" +
            "02/01/2023 - Entrega - Não\n" +
            "Remunerações Diárias\n" +
            "01/01/2023 - Bônus - R$ 30,00\n"
        )
        nome, entregas, acrescimos, bonus = extrair_dados_pdf("dummy.pdf")
        self.assertEqual(nome, "MOTORISTA A")
        self.assertEqual(entregas[datetime(2023, 1, 1).date()]["entregues"], 1)
        self.assertEqual(entregas[datetime(2023, 1, 2).date()]["insucessos"], 1)
        self.assertIn(datetime(2023, 1, 1).date(), bonus)

    @patch("pdfplumber.open")
    def test_extrair_dados_pdf_cenario2(self, mock_pdfplumber_open):
        # Mock do PDF para o Cenário 2
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_pdf.pages = [mock_page]
        mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

        # Cenário 2: PDF com acréscimo (linha a linha)
        mock_page.extract_text.return_value = (
            "Motorista: MOTORISTA B\n" +
            "03/01/2023 - Acréscimo - R$ 15,50\n"
        )
        nome, entregas, acrescimos, bonus = extrair_dados_pdf("dummy2.pdf")
        self.assertEqual(nome, "MOTORISTA B")
        self.assertEqual(acrescimos[datetime(2023, 1, 3).date()], 15.50)

    @patch("pdfplumber.open")
    def test_extrair_dados_pdf_cenario3(self, mock_pdfplumber_open):
        # Mock do PDF para o Cenário 3
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_pdf.pages = [mock_page]
        mock_pdfplumber_open.return_value.__enter__.return_value = mock_pdf

        # Cenário 3: PDF com acréscimo em tabela
        mock_page.extract_text.return_value = (
            "Motorista: MOTORISTA C\n" +
            "ACRÉSCIMOS\n" +
            "Data Valor\n" +
            "04/01/2023 R$ 25,00\n" +
            "05/01/2023 R$ 10,00\n"
        )
        mock_page.extract_tables.return_value = [
            [["Data", "Valor"], ["04/01/2023", "R$ 25,00"], ["05/01/2023", "R$ 10,00"]]
        ]
        nome, entregas, acrescimos, bonus = extrair_dados_pdf("dummy3.pdf")
        self.assertEqual(nome, "MOTORISTA C")
        self.assertAlmostEqual(acrescimos[datetime(2023, 1, 4).date()], 25.00)
        self.assertAlmostEqual(acrescimos[datetime(2023, 1, 5).date()], 10.00)

    def test_calcular_fechamento(self):
        # Dados de teste
        nome_motorista = "MOTORISTA A"
        entregas_por_dia = defaultdict(lambda: {"entregues": 0, "insucessos": 0})
        entregas_por_dia[datetime(2023, 1, 1).date()]["entregues"] = 5
        entregas_por_dia[datetime(2023, 1, 2).date()]["insucessos"] = 2
        acres_por_data = defaultdict(float)
        acres_por_data[datetime(2023, 1, 1).date()] = 10.0
        bonus_pago_dates = {datetime(2023, 1, 1).date()}

        nome_final, df = calcular_fechamento(nome_motorista, entregas_por_dia, acres_por_data, bonus_pago_dates)

        self.assertEqual(nome_final, "MOTORISTA A")
        self.assertFalse(df.empty)

        # Verifica os totais
        total_row = df[df["Data"] == "Total"].iloc[0]
        self.assertEqual(total_row["Entregues"], 5)
        self.assertEqual(total_row["Insucessos"], 2)
        self.assertAlmostEqual(total_row["Valor Entregas"], 5 * VALOR_ENTREGA)
        self.assertAlmostEqual(total_row["Descontos"], 2 * VALOR_ENTREGA)
        self.assertAlmostEqual(total_row["Acréscimo Pago"], 10.0)
        self.assertAlmostEqual(total_row["Bônus"], BONUS_DIARIO)

    @patch("os.listdir")
    @patch("os.path.exists")
    @patch("script_fechamento.extrair_dados_pdf")
    @patch("script_fechamento.calcular_fechamento")
    @patch("pandas.ExcelWriter")
    @patch("openpyxl.load_workbook")
    @patch("pandas.read_excel")
    def test_main_pdf_grouping_and_consolidation(self, mock_pd_read_excel, mock_load_workbook, mock_excel_writer, mock_calcular_fechamento, mock_extrair_dados_pdf, mock_exists, mock_listdir):
        mock_exists.return_value = True
        mock_listdir.return_value = ["ELISIANELUDMYLLAFERREIRASANTOS.pdf", "ELISIANELUDMYLLAFERREIRASANTOS2.pdf"]

        # Mock para extrair_dados_pdf para o primeiro PDF
        mock_extrair_dados_pdf.side_effect = [
            ("ELISIANE LUDMYLLA FERREIRA SANTOS", 
             {datetime(2023, 1, 1).date(): {"entregues": 1, "insucessos": 0}},
             {datetime(2023, 1, 1).date(): 5.0},
             set()),
            # Mock para extrair_dados_pdf para o segundo PDF
            ("ELISIANE LUDMYLLA FERREIRA SANTOS", 
             {datetime(2023, 1, 2).date(): {"entregues": 1, "insucessos": 0}},
             {datetime(2023, 1, 2).date(): 10.0},
             set())
        ]

        # Mock para calcular_fechamento
        mock_calcular_fechamento.return_value = (
            "ELISIANE LUDMYLLA FERREIRA SANTOS",
            pd.DataFrame({
                "Data": ["01/01/2023", "02/01/2023", "Total"],
                "Motorista": ["ELISIANE LUDMYLLA FERREIRA SANTOS"] * 3,
                "Entregues": [1, 1, 2],
                "Acréscimo Pago": [5.0, 10.0, 15.0]
            })
        )

        # Mock do ExcelWriter para evitar FileNotFoundError
        mock_writer_instance = MagicMock()
        mock_excel_writer.return_value.__enter__.return_value = mock_writer_instance
        mock_excel_writer.return_value.__exit__.return_value = None # Garante que __exit__ não levante exceção

        # Mock do load_workbook para evitar FileNotFoundError quando o arquivo não existe
        def mock_load_workbook_side_effect(filename):
            if not os.path.exists(filename):
                raise FileNotFoundError
            return MagicMock()
        mock_load_workbook.side_effect = mock_load_workbook_side_effect

        # Mock do pd.read_excel para retornar um DataFrame vazio se a aba não existe
        def mock_pd_read_excel_side_effect(filename, sheet_name):
            if not os.path.exists(filename):
                raise FileNotFoundError
            return pd.DataFrame() # Retorna um DataFrame vazio se o arquivo existe mas a aba não
        mock_pd_read_excel.side_effect = mock_pd_read_excel_side_effect

        main()

        # Verifica se extrair_dados_pdf foi chamado para ambos os PDFs
        self.assertEqual(mock_extrair_dados_pdf.call_count, 2)
        # Verifica se calcular_fechamento foi chamado uma vez para o motorista consolidado
        self.assertEqual(mock_calcular_fechamento.call_count, 1)
        # Verifica se o ExcelWriter foi chamado com a aba correta
        mock_writer_instance.to_excel.assert_called_once()
        args, kwargs = mock_writer_instance.to_excel.call_args
        self.assertEqual(kwargs["sheet_name"], "ELISIANE LUDMYLLA FERREIRA S")

if __name__ == "__main__":
    unittest.main()


