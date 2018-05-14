import pandas as pd
from pathlib import Path
import numpy as np


class Reader:
    def __init__(self, list_path):
        self.list_path = Path(list_path)
        self.list_file = self.read_list()

    def read_list(self):
        df = pd.read_excel(f'{self.list_path.resolve()}')
        # Shorten dataframe at real data
        df = df[df['#'].apply(lambda x: isinstance(x, (int, np.int64)))]

        # define delivery address
        df['address_1'] = df.apply(lambda x: self.delivery_address(x, 1), axis=1)
        df['address_2'] = df.apply(lambda x: self.delivery_address(x, 2), axis=1)
        df['address_3'] = df.apply(lambda x: self.delivery_address(x, 3), axis=1)
        return df

    @staticmethod
    def delivery_address(row, type):
        if row['Expédié avec'] == 'E':
            ship = 1
        else:
            ship = 2

        if type == 1:
            if ship == 1:
                return row['EXPÉDIBUS']
            else:
                return row['ORGANISME']

        elif type == 2:
            if ship == 1:
                return row['ADRESSE EXPÉDIBUS']
            else:
                return row['ADRESSE ORGANISME'].split('\n')[0]

        elif type == 3:
            if ship == 1:
                return row['VILLE EXPÉDIBUS']
            else:
                return row['ADRESSE ORGANISME'].split('\n')[1]
