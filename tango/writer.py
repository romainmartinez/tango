from pathlib import Path
import openpyxl


class Writer:
    def __init__(
        self,
        template_invoice_path,
        template_etq_path,
        list_file,
        export_path,
        logo_path,
    ):
        self.template_invoice_path = Path(template_invoice_path)
        self.template_etq_path = Path(template_etq_path)
        self.list_file = list_file

        self.export_path = export_path
        self.logo_path = logo_path

        self.main_loop()

    def main_loop(self):
        for i, irow in self.list_file.iterrows():
            print(f"\t{irow['#']}")

            invoice_filename = self.set_filename(irow, type="factures")
            self.write_invoice(irow, filename=invoice_filename)

            etq_filename = self.set_filename(irow, type="etiquettes")
            self.write_etq(irow, filename=etq_filename)

    def set_filename(self, row, type):
        date = row["DATE D'ACTIVITÉ"].date()
        filename = Path(
            self.export_path, type, f'{row["#"]}_{date}_{row["ORGANISME"][:10]}.xlsx'
        )
        return f"{filename}"

    def write_invoice(self, row, filename):
        workbook = openpyxl.load_workbook(f"{self.template_invoice_path}")
        worksheet = workbook["Bon livraison"]

        worksheet["B3"] = row["#"]
        worksheet["B4"] = row["FICHE PEH"]
        worksheet["B5"] = row["DATE D'ACTIVITÉ"].date()

        worksheet["A9"] = row["ORGANISME"]
        worksheet["B10"] = row["RESPONSABLE"]
        worksheet["B11"] = row["CELLULAIRE"]
        worksheet["B12"] = row["TÉLÉPHONE BUR."]
        worksheet["B13"] = row["TÉLÉPHONE RÉS"]
        worksheet["B14"] = row["COURRIEL"]

        worksheet["F9"] = row["address_1"]
        worksheet["F10"] = row["address_2"]
        worksheet["F11"] = row["address_3"]

        worksheet["F17"] = row["CANNE"]
        worksheet["F24"] = row["GUIDE ÉTÉ"]
        worksheet["F25"] = row["BANNIÈRE"]

        if row["PERMIS"]:
            worksheet["F18"] = row["PERMIS"]
            # worksheet["F19"] = row["DÉBUT # PERMIS"]
            # worksheet["G19"] = row["FIN # PERMIS"]

        img = openpyxl.drawing.image.Image(f"{self.logo_path}")
        worksheet.add_image(img, "H2")

        workbook.save(filename)

    def write_etq(self, row, filename):
        workbook = openpyxl.load_workbook(f"{self.template_etq_path}")
        worksheet = workbook["Etiquette"]
        date = row["DATE D'ACTIVITÉ"].date()
        info = f"#{row['#']} - {row['ORGANISME']}_n_" f"{date}"

        coord = (
            f"{row['RESPONSABLE']}_n_"
            f"Cellulaire : {row['CELLULAIRE']}_n_"
            f"bureau : {row['TÉLÉPHONE BUR.']}_n_"
            f"Maison : {row['TÉLÉPHONE RÉS']}"
        )

        address = (
            f"{row['address_1']}_n_" f"{row['address_2']}_n_" f"{row['address_3']}"
        )

        info_cell = ["1", "4", "7", "10"]
        coord_cell = ["2", "5", "8", "11"]
        address_cell = ["3", "6", "9", "12"]
        for l in ["A", "B"]:
            for info_c, coord_c, address_c in zip(info_cell, coord_cell, address_cell):
                worksheet[f"{l}{info_c}"] = info.replace("_n_", "\n")
                worksheet[f"{l}{coord_c}"] = coord.replace("_n_", "\n").replace(
                    "nan", ""
                )
                worksheet[f"{l}{address_c}"] = address.replace("_n_", "\n")

        workbook.save(filename)
