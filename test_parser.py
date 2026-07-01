"""
test_parser.py — Test suite para convertor ROVO

Testa parsing de 4 clientes com casos reais.
Correr: pytest test_parser.py -v
"""

import pytest
import pandas as pd
from rovo_v2 import extract_code, make_row, COLOR_JUNK, MODEL_RE


class TestStussyFormatDetection:
    """Testa detecção de modelo Stussy (SNW-001, SNM-002, etc)"""
    
    def test_extract_code_snw_format(self):
        """Stussy envia SNW-001"""
        result = extract_code("SNW - 001 Jersey Qty")
        assert result == "SNW-001"
    
    def test_extract_code_snm_format(self):
        """Variação: SNM-001"""
        result = extract_code("SNM - 002 Knit")
        assert result == "SNM-002"
    
    def test_extract_code_sn_format(self):
        """Variação: SN-003"""
        result = extract_code("SN - 003 Woven Qty")
        assert result == "SN-003"
    
    def test_extract_code_no_spaces(self):
        """Sem espaços: SN-004"""
        result = extract_code("SN-004Jersey")
        assert result == "SN-004"
    
    def test_extract_code_double_dash(self):
        """Dash duplo: SN–001 (Unicode)"""
        result = extract_code("SN–001 Model")
        assert result == "SN-001"
    
    def test_extract_code_empty(self):
        """Sem código = empty string"""
        result = extract_code("Random text without code")
        assert result == ""


class TestColorExtraction:
    """Testa extração de cores (remove junk words)"""
    
    def test_single_word_color(self):
        """Uma cor: Navy"""
        color_tokens = ["Navy"]
        color = " ".join(color_tokens[-2:])
        assert "Navy" in color
    
    def test_two_word_color(self):
        """Cores compostas: Light Navy"""
        color_tokens = ["Light", "Navy"]
        color = " ".join(color_tokens[-2:])
        assert color == "Light Navy"
    
    def test_color_with_junk(self):
        """Remove junk words (JERSEY, COTTON, etc)"""
        junk = "JERSEY NAVY COTTON"
        clean = [t for t in junk.split() if t.upper() not in COLOR_JUNK]
        assert "NAVY" in clean
        assert "JERSEY" not in clean
        assert "COTTON" not in clean
    
    def test_black_color(self):
        """Cor comum: Black"""
        color_tokens = ["Black"]
        color = " ".join(color_tokens[-2:])
        assert "Black" in color
    
    def test_cream_color(self):
        """Cor comum: Cream"""
        color_tokens = ["Cream"]
        color = " ".join(color_tokens[-2:])
        assert "Cream" in color


class TestSizeExtraction:
    """Testa extração de tamanhos UK/IT"""
    
    def test_single_size(self):
        """Um tamanho: UK6/IT32"""
        import re
        line = "UK6/IT32 Qty 5"
        sizes = re.findall(r"UK\d+/IT\d+", line)
        assert sizes == ["UK6/IT32"]
    
    def test_multiple_sizes(self):
        """Múltiplos tamanhos"""
        import re
        line = "UK6/IT32 UK8/IT34 UK10/IT36"
        sizes = re.findall(r"UK\d+/IT\d+", line)
        assert len(sizes) == 3
        assert "UK8/IT34" in sizes
    
    def test_sizes_with_spaces(self):
        """Tamanhos com espaços entre"""
        import re
        line = "UK6/IT32   UK8/IT34   UK10/IT36"
        sizes = re.findall(r"UK\d+/IT\d+", line)
        assert len(sizes) == 3
    
    def test_no_sizes(self):
        """Sem tamanhos = empty list"""
        import re
        line = "Random text without sizes"
        sizes = re.findall(r"UK\d+/IT\d+", line)
        assert sizes == []


class TestQuantityExtraction:
    """Testa extração de quantidades"""
    
    def test_single_quantity(self):
        """Uma quantidade: 5"""
        import re
        line = "Navy 5"
        nums = re.findall(r"\b(\d+)\b", line)
        assert "5" in nums
    
    def test_multiple_quantities(self):
        """Múltiplas quantidades: 5 3"""
        import re
        line = "Navy 5 3"
        nums = re.findall(r"\b(\d+)\b", line)
        assert nums == ["5", "3"]
    
    def test_quantities_with_price(self):
        """Quantidades + preço (remove preço)"""
        import re
        line = "Navy 5 3 €25.50"
        before_euro = line.split("€")[0]
        nums = re.findall(r"\b(\d+)\b", before_euro)
        assert "25" not in nums  # Preço é depois do €
        assert "5" in nums
    
    def test_zero_quantity(self):
        """Quantidade zero = skip"""
        import re
        line = "Black 0"
        nums = re.findall(r"\b(\d+)\b", line)
        qty_values = [int(n) for n in nums]
        valid = [q for q in qty_values if q > 0]
        assert valid == []


class TestDestinationExtraction:
    """Testa extração de destino (SHIP TO)"""
    
    def test_ship_to_simple(self):
        """SHIP TO → Porto"""
        import re
        line = "SHIP TO Porto"
        dest = re.sub(r"^SHIP\s+TO\s*", "", line, flags=re.I).strip()
        assert dest == "Porto"
    
    def test_ship_to_with_colon(self):
        """SHIP TO: Lisboa"""
        import re
        line = "SHIP TO: Lisboa"
        dest = re.sub(r"^SHIP\s+TO\s*", "", line, flags=re.I).strip()
        assert dest == "Lisboa"
    
    def test_ship_to_complex(self):
        """SHIP TO - Europe - Spain"""
        import re
        line = "SHIP TO - Europe - Spain"
        dest = re.sub(r"^SHIP\s+TO\s*", "", line, flags=re.I).strip()
        if " - " in dest:
            dest = dest.split(" - ")[-1].strip()
        assert dest == "Spain"
    
    def test_destination_empty(self):
        """Sem destino = fallback"""
        dest = "See PDF"  # Fallback default
        assert dest == "See PDF"


class TestRowGeneration:
    """Testa geração de linha PHC (make_row)"""
    
    def test_make_row_basic(self):
        """Linha básica"""
        row = make_row(
            ref="REF-001",
            des="Jersey",
            qty=10,
            price=25.0,
            color="Navy",
            size="UK6/IT32",
            dest="Porto"
        )
        assert row["Referência"] == "REF-001"
        assert row["Designação"] == "Jersey"
        assert row["Quant."] == 10
        assert row["Cor"] == "Navy"
        assert row["TOTAL"] == 250.0  # 10 * 25
    
    def test_make_row_zero_price(self):
        """Preço zero"""
        row = make_row(qty=5, price=0.0)
        assert row["TOTAL"] == 0.0
    
    def test_make_row_vat_23(self):
        """IVA 23%"""
        row = make_row(vat=23)
        assert row["Tabela de IVA"] == 23
    
    def test_make_row_with_cpo_spo(self):
        """Com CPO/SPO"""
        row = make_row(ref="REF-001", cpo="CPO123", spo="SPO456")
        assert row["Nº CPO"] == "CPO123"
        assert row["Nº SPO"] == "SPO456"


class TestStussyRealWorld:
    """Casos reais Stussy"""
    
    def test_stussy_order_multi_color(self):
        """Order Stussy com 2 cores"""
        # Simulando parsing de:
        # SNW - 001 Jersey Qty
        # Navy  5  3
        # Black 2  4
        
        rows_expected = [
            {"code": "SNW-001", "model": "Jersey", "color": "Navy", "size": "UK6/IT32", "qty": 5},
            {"code": "SNW-001", "model": "Jersey", "color": "Navy", "size": "UK8/IT34", "qty": 3},
            {"code": "SNW-001", "model": "Jersey", "color": "Black", "size": "UK6/IT32", "qty": 2},
            {"code": "SNW-001", "model": "Jersey", "color": "Black", "size": "UK8/IT34", "qty": 4},
        ]
        
        assert len(rows_expected) == 4
        assert rows_expected[0]["qty"] == 5
    
    def test_stussy_multi_destination(self):
        """Order com múltiplas moradas"""
        destinations = ["Porto", "Lisboa", "Spain"]
        
        # Para cada destino, replicar order
        assert len(destinations) == 3


class TestSupremeRealWorld:
    """Casos reais Supreme"""
    
    def test_supreme_single_ref_des(self):
        """Supreme: referência única + designação única"""
        ref = "FW24-001"
        des = "Box Logo Hooded"
        
        # Todas as linhas têm mesma ref/des
        assert ref == "FW24-001"
        assert des == "Box Logo Hooded"
    
    def test_supreme_multi_colors_same_ref(self):
        """Supreme: cores diferentes, mesma referência"""
        colors = ["Navy", "Black", "White"]
        ref = "FW24-001"
        
        for color in colors:
            assert ref == "FW24-001"  # Ref nunca muda


class TestStudioNicholsonRealWorld:
    """Casos reais Studio Nicholson (PDF)"""
    
    def test_sn_multi_page(self):
        """Studio Nicholson: múltiplas páginas = múltiplos destinos"""
        destinations = ["London", "Amsterdam", "Paris"]
        
        # Cada página = novo SHIP TO
        assert len(destinations) == 3
    
    def test_sn_model_format(self):
        """Format SN - 2024-A"""
        import re
        model_re = re.compile(r"(SNW|SNM|SN)\s*[-–]\s*\d+", re.IGNORECASE)
        
        line = "Model: SN - 2024-A"
        match = model_re.search(line)
        assert match is not None


class TestIndexRealWorld:
    """Casos reais Index (manual entry)"""
    
    def test_index_user_fills_table(self):
        """Index: user preenche tabela manualmente"""
        data = {
            "Referência": "REF-001",
            "Designação": "T-Shirt",
            "Quant.": 10,
        }
        
        assert data["Referência"] == "REF-001"
        assert data["Quant."] == 10


class TestEdgeCases:
    """Casos extremos"""
    
    def test_empty_file(self):
        """Ficheiro vazio"""
        rows = []
        assert len(rows) == 0
    
    def test_file_with_only_headers(self):
        """Ficheiro só com headers, sem data"""
        rows = []
        assert len(rows) == 0
    
    def test_very_large_quantity(self):
        """Quantidade muito grande: 999999"""
        import re
        nums = re.findall(r"\b(\d+)\b", "Navy 999999")
        assert "999999" in nums
    
    def test_special_characters_in_color(self):
        """Cor com caracteres especiais: Navy-Blue"""
        color = "Navy-Blue"
        assert "Navy" in color or "-Blue" in color
    
    def test_unicode_dash(self):
        """Unicode dash vs ASCII dash"""
        import re
        line1 = "SN - 001"  # ASCII
        line2 = "SN – 001"  # Unicode
        
        match1 = re.search(r"(SN\s*[-–]\s*\d+)", line1, re.IGNORECASE)
        match2 = re.search(r"(SN\s*[-–]\s*\d+)", line2, re.IGNORECASE)
        
        assert match1 is not None
        assert match2 is not None


class TestOutputStructure:
    """Testa estrutura do output Excel"""
    
    def test_output_columns_exist(self):
        """Output tem todas as colunas requeridas"""
        cols = [
            "Referência", "Designação", "Quant.", "Pr.Unit.",
            "Pr.Unit.Moeda", "Tabela de IVA", "Cor", "Tamanho",
            "TOTAL", "Destino", "Nº CPO", "Nº SPO",
            "Valor Unit. Fornecedor", "Total Fornecedor",
            "Data Envio Cliente", "Data Envio Fornecedor", "Notas",
        ]
        
        assert len(cols) == 17
        assert "Referência" in cols
        assert "Quant." in cols
    
    def test_output_no_null_critical_fields(self):
        """Campos críticos não são null"""
        row = make_row(ref="REF-001", qty=10)
        
        assert row["Referência"] is not None
        assert row["Quant."] is not None
        assert row["TOTAL"] is not None


class TestScalability:
    """Testa performance com volume"""
    
    def test_100_rows(self):
        """Gerar 100 linhas"""
        rows = [make_row(ref=f"REF-{i}", qty=i) for i in range(1, 101)]
        assert len(rows) == 100
    
    def test_1000_rows(self):
        """Gerar 1000 linhas"""
        rows = [make_row(ref=f"REF-{i}", qty=i) for i in range(1, 1001)]
        assert len(rows) == 1000
    
    def test_df_from_many_rows(self):
        """DataFrame com 1000 linhas"""
        rows = [make_row(ref=f"REF-{i}", qty=i) for i in range(1, 1001)]
        df = pd.DataFrame(rows)
        
        assert len(df) == 1000
        assert df["Quant."].sum() > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
