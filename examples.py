from tango import Writer
from tango import Reader

from pathlib import Path

ROOT = Path('/home/romain/Downloads/data')
LIST = ROOT / '1_LISTE_ENVOI_TANGO_11Mai2018_v5.xlsx'
TEMPLATE_INVOICE = ROOT / 'modeleFFQ2018_v4.xlsx'
TEMPLATE_ETQ = ROOT / 'modeleETQ.xlsx'
OUTPUT = ROOT / 'output'
LOGO = ROOT / 'logo_2.png'

reader = Reader(list_path=LIST)
writer = Writer(
    template_invoice_path=TEMPLATE_INVOICE,
    template_etq_path=TEMPLATE_ETQ,
    list_file=reader.list_file,
    logo_path=LOGO,
    export_path=OUTPUT
)
