from tkinter import messagebox

import fitz  # PyMuPDF --> estrazione testo, immagini e tabelle


def rect_coverage(x0, y0, x1, y1, a0, b0, a1, b1):
    if b0 <= y0 and b1 >= y1 and a0 <= x0 and a1 >= x1:
        return True
    else:
        return False


class PDF:
    def __init__(self, document):
        self.doc = fitz.open(document)
        self.width = self.doc[0].rect.width
        self.height = self.doc[0].rect.height

    def __del__(self):
        if not self.is_closed():
            self.doc.close()

    def is_closed(self):
        return self.doc.is_closed

    def get_height(self):
        return self.height

    def get_width(self):
        return self.width

    def get_page_number(self):
        return self.doc.page_count - 1

    def exists_img(self, p):
        page = self.doc[p]
        images = page.get_images()
        if len(images) == 0:
            messagebox.showerror(title="Errore",
                                 message="Non esistono immagini all'interno della pagina selezionata")
            return False
        else:
            return True

    def exists_tab(self, p):
        page = self.doc[p]
        tables = page.find_tables().tables
        if len(tables) == 0:
            messagebox.showerror(title="Errore",
                                 message="Non esistono tabelle all'interno della pagina selezionata")
            return False
        else:
            return True

    def get_text(self, x0, y0, x1, y1, p):
        rect = (x0, y0, x1, y1)
        page = self.doc[p]
        return page.get_textbox(rect)

    def get_img(self, x0, y0, x1, y1, p):
        page = self.doc[p]
        info_images = page.get_image_info(True, True)
        data_img = None
        ext = None
        for img in info_images:
            bbox = img['bbox']
            if rect_coverage(bbox[0], bbox[1], bbox[2], bbox[3], x0, y0, x1, y1):
                base_img = self.doc.extract_image(img['xref'])
                data_img = base_img['image']
                ext = base_img['ext']
                break
        return data_img, ext

    def get_tab(self, x0, y0, x1, y1, p):
        page = self.doc[p]
        data = []
        df = None
        exists_table = False
        tables = page.find_tables().tables
        for t in tables:
            bbox = t.bbox
            if rect_coverage(bbox[0], bbox[1], bbox[2], bbox[3], x0, y0, x1, y1):
                df = t.to_pandas()
                exists_table = True
                break
        if exists_table:
            column = [*df]
            for index in range(len(df)):
                row = {}
                for c in column:
                    row[c] = df.loc[index, c]
                data.append(row)
        return data
