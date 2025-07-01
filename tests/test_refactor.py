import pytest
import pandas as pd
from io import BytesIO

from refactor import RefactorExcel


class TestRefactorExcel:
    @pytest.fixture(scope='function')
    def sample_file(self, tmp_path):
        """
        Создает временный Excel-файл с тестовыми данными.
        """
        df = pd.DataFrame({
            'ID Материала': ['ABI23', 'I45X', 'C67'],
            'Кол-во по заявке': ['10', 'I5', '3'],
            'Поступило всего': ['8', '15', '2']
        })
        path = tmp_path / 'input.xlsx'
        df.to_excel(path, index=False, engine='openpyxl')
        return path

    def test_load_and_normalize_id(self, sample_file):
        processor = RefactorExcel(sample_file)
        processor.load_file()
        processor.normalize_id_material()
        df = processor.df
        assert df.loc[0, 'ID Материала'] == '123'
        assert df.loc[1, 'ID Материала'] == '145'

    def test_filter_and_add_difference(self, sample_file):
        processor = RefactorExcel(sample_file)
        processor.load_file()
        processor.normalize_id_material()
        processor.filter_rows()
        processor.add_difference()
        df = processor.df

        assert len(df) == 2
        assert df['Расхождение заявка-приход'].iloc[0] == pytest.approx(2)
        assert df['Расхождение заявка-приход'].iloc[1] == pytest.approx(1)

    def test_run_and_save(self, sample_file, tmp_path):
        out_path = tmp_path / 'output.xlsx'
        processor = RefactorExcel(sample_file)
        processor.run(out_path)
        result = pd.read_excel(out_path, engine='openpyxl')

        assert 'Расхождение заявка-приход' in result.columns
        assert len(result) == 2  # Было 1, должно быть 2
        assert result['Расхождение заявка-приход'].iloc[0] == pytest.approx(2)
        assert result['Расхождение заявка-приход'].iloc[1] == pytest.approx(1)

    def test_save_to_buffer(self, sample_file):
        processor = RefactorExcel(sample_file)
        processor.load_file()
        processor.normalize_id_material()
        processor.filter_rows()
        processor.add_difference()
        buffer = BytesIO()
        processor.df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        df2 = pd.read_excel(buffer, engine='openpyxl')
        assert 'Расхождение заявка-приход' in df2.columns
        assert buffer.getvalue() != b''
