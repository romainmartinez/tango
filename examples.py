from tango import Writer
from tango import Reader

from pathlib import Path

ROOT = Path("/home/romain/Downloads/tango/2019")
LIST = ROOT / "input" / "LISTE_ENVOI_TANGO_3_05_2019.xlsx"
TEMPLATE_INVOICE = ROOT / "template" / "facture.xlsx"
TEMPLATE_ETQ = ROOT / "template" / "etiquette.xlsx"
LOGO = ROOT / "input" / "logo.png"
OUTPUT = ROOT / "output"

reader = Reader(list_path=LIST)
writer = Writer(
    template_invoice_path=TEMPLATE_INVOICE,
    template_etq_path=TEMPLATE_ETQ,
    list_file=reader.list_file,
    logo_path=LOGO,
    export_path=OUTPUT,
)
