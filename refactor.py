import pandas as pd
from pathlib import Path

class RefactorExcel:
    """ Класс для модернизации xlsx файла """

    def __init__(self, path_to_file: str | Path, sheet_name: str | int = 0) -> None:
        self.path_to_file = path_to_file
        self.sheet_name = sheet_name
        self.df = None

    def load_file(self) -> None:
        """1. Чтение данных из файла"""
        ext = Path(self.path_to_file).suffix.lower()

        if ext == ".xlsb":
            self.df = pd.read_excel(self.path_to_file,
                                    sheet_name=self.sheet_name,
                                    dtype='str',
                                    engine='pyxlsb')
        else:
            self.df = pd.read_excel(self.path_to_file,
                                    sheet_name=self.sheet_name,
                                    dtype='str',
                                    engine='openpyxl')

    def normalize_id_material(self) -> None:
        """2. Нормализация колонки E и числовых колонок"""
        if 'ID Материала' in self.df.columns:
            self.df['ID Материала'] = (
                self.df['ID Материала']
                .astype(str)
                .str.replace('I', '1', regex=False)
                .str.replace(r'\D+', '', regex=True)
            )

        for col in ['Кол-во по заявке', 'Поступило всего']:
            if col in self.df.columns:
                self.df[col] = (
                    self.df[col]
                    .astype(str)
                    .str.replace('I', '1', regex=False)
                    .str.extract(r'(\d+)', expand=False)
                    .astype(float)
                )

    def filter_rows(self) -> None:
        """3. Оставить только строки, где заявка > приход"""
        req, rec = 'Кол-во по заявке', 'Поступило всего'
        if {req, rec}.issubset(self.df.columns):
            self.df = self.df[self.df[req] > self.df[rec]].copy()

    def add_difference(self) -> None:
        """4. Добавить колонку 'Расхождение заявка-приход'"""
        req, rec = 'Кол-во по заявке', 'Поступило всего'
        if {req, rec}.issubset(self.df.columns):
            print("Типы колонок:", self.df[req].dtype, self.df[rec].dtype)
            print(self.df[[req, rec]].head())

            self.df['Расхождение заявка-приход'] = self.df[req] - self.df[rec]


    def save(self, output_path: str | Path) -> None:
        """5. Сохранить результат в xlsx-файл"""
        self.df.to_excel(Path(output_path), index=False, engine='openpyxl')

    def run(self, output_path: str | Path) -> None:
        """Выполнить весь цикл: load, normalize, filter_rows, add_difference, save"""
        self.load_file()
        self.normalize_id_material()
        print("Колонки в файле:", list(self.df.columns))

        self.filter_rows()
        self.add_difference()
        self.save(output_path)