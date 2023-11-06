import os
import tkinter
import Const
from Extract import PDF
from Mongo import Database
import customtkinter as ctk
from tkinter import filedialog, StringVar, messagebox, Canvas
import io
import PIL  # --> libreria per lavorare su immagini
from PIL import Image, ImageTk  # --> resize immagine e visualizzazione in canvas
from pdf2image import convert_from_path  # --> convertire pdf in immagine così da visualizzare le pagine

poppler_path = r"C:\Poppler\Release-23.08.0-0\poppler-23.08.0\Library\bin"


def not_point(x0, y0, x1, y1):
    if x0 != x1 and y0 != y1:
        return True
    else:
        return False


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # inizzializzazione variabili
        self.image_canvas = None
        self.rule_name = StringVar()
        self.selected_rule = StringVar()
        self.rules = []
        self.all_rules = {}
        self.destination_directory = StringVar()
        self.obj = {}
        self.image_to_save = []
        self.menu = None
        self.type = StringVar()
        self.rect_id = None
        self.output_image = None
        self.coord = {}
        self.page = 0
        self.image_counter = 0
        self.topx, self.topy, self.botx, self.boty = 0, 0, 0, 0
        self.ask_add_coord = True
        self.tab = []
        self.img = []
        self.txt = []
        self.height = 0
        self.width = 0
        self.number_of_page = 0
        self.selected_file = StringVar()
        self.pdf_path = StringVar()
        self.extractor = None

        # setup
        ctk.set_appearance_mode('light')
        self.geometry('1000x630+0+0')
        self.title('PDF extractor')
        self.iconbitmap('./PDF.ico')
        self.protocol("WM_DELETE_WINDOW", self.close_root_window)

        # layout
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1, uniform='a')
        self.columnconfigure(1, weight=1, uniform='a')

        # widgets
        self.output_image = Canvas(self)
        self.output_image.grid(row=0, column=1, sticky='nsew')

        # menu comandi
        self.menu = ctk.CTkTabview(self)
        self.menu.grid(row=0, column=0, sticky='nsew', ipadx=10, ipady=10, padx=10, pady=10)
        self.menu.add('Comandi')
        self.menu.add('Regole')
        self.menu.add('Estrai')
        self.menu.add('Memorizza')

        # frame comandi
        self.commands_frame = ctk.CTkFrame(self.menu.tab('Comandi'), fg_color='transparent')
        self.commands_frame.pack(expand=True, fill=tkinter.BOTH)
        self.commands_frame.rowconfigure(0, weight=1)
        self.commands_frame.rowconfigure(1, weight=2)
        self.commands_frame.columnconfigure(0, weight=1)

        self.btn_open = ctk.CTkButton(self.commands_frame, text='Apri PDF', command=self.open_file)
        self.btn_open.grid(row=0, column=0, columnspan=3, ipadx=5, ipady=5, padx=5, pady=5)

        self.selected_file.set('Nessun file PDF selezionato')
        self.lbl_file = ctk.CTkLabel(self.commands_frame, textvariable=self.selected_file)
        self.lbl_file.grid(row=0, column=0, columnspan=3, ipadx=5, ipady=5, padx=5, pady=5, sticky=tkinter.S)

        self.btn_others_files = ctk.CTkButton(self.commands_frame, text='Seleziona più PDF',
                                              command=self.open_others_files)
        self.btn_others_files.grid(row=2, column=0, ipadx=5, ipady=5, padx=5, pady=5)

        # frame rules
        self.rules_frame = ctk.CTkFrame(self.menu.tab('Regole'), fg_color='transparent')
        self.rules_frame.pack(expand=True, fill=tkinter.BOTH)
        self.rules_frame.rowconfigure(0, weight=1)
        self.rules_frame.rowconfigure(1, weight=2)
        self.rules_frame.columnconfigure(0, weight=1)
        self.rules_frame.columnconfigure(1, weight=1)

        self.txt_rule_name = ctk.CTkEntry(self.rules_frame, placeholder_text='Inserire nome regola',
                                          textvariable=self.rule_name)
        self.txt_rule_name.grid(row=0, column=0, ipadx=5, ipady=5, padx=5, pady=5)
        self.btn_save_rule = ctk.CTkButton(self.rules_frame, text='Salva regola di estrazione',
                                           command=self.save_rule)
        self.btn_save_rule.grid(row=0, column=1, ipadx=5, ipady=5, padx=5, pady=5)

        self.lbl_rules = ctk.CTkLabel(self.rules_frame, text="Scegli una regola da \n utilizzare per le tue estrazioni")
        self.lbl_rules.grid(row=1, column=0, ipadx=5, ipady=5, padx=5, pady=5)

        self.combobox_rules = ctk.CTkComboBox(self.rules_frame, variable=self.selected_rule,
                                              command=self.set_rule, state="readonly")
        self.combobox_rules.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # frame memorizza
        self.save_frame = ctk.CTkFrame(self.menu.tab('Memorizza'), fg_color='transparent')
        self.save_frame.pack(expand=True, fill=tkinter.BOTH)
        self.save_frame.rowconfigure(0, weight=1)
        self.save_frame.rowconfigure(1, weight=1)
        self.save_frame.rowconfigure(2, weight=10)
        self.save_frame.columnconfigure(0, weight=1)
        self.save_frame.columnconfigure(1, weight=1)
        self.save_frame.columnconfigure(2, weight=1)

        self.btn_save = ctk.CTkButton(self.save_frame, text='Salva i dati estratti',
                                      command=self.save)
        self.btn_save.grid(row=0, column=0, columnspan=4, ipadx=5, ipady=5, padx=5, pady=5)

        self.lbl_dest_dir = ctk.CTkLabel(self.save_frame,
                                         text="Cartella di destinazione \n delle immagini")
        self.lbl_dest_dir.grid(row=1, column=0, ipadx=5, ipady=5, padx=5, pady=5, sticky=tkinter.EW)
        self.entry_dest_dir = ctk.CTkEntry(self.save_frame, textvariable=self.destination_directory,
                                           state='disabled', xscrollcommand=True)
        self.entry_dest_dir.grid(row=1, column=2, ipadx=5, ipady=5, padx=5, pady=5, sticky=tkinter.EW)
        self.btn_dest_dir = ctk.CTkButton(self.save_frame, text='Apri explorer', command=self.choose_dest_dir)
        self.btn_dest_dir.grid(row=1, column=3, ipadx=5, ipady=5, padx=5, pady=5)

        self.btn_clear = ctk.CTkButton(self.save_frame, text='Clear', command=self.clear)
        self.btn_clear.grid(row=2, column=0, columnspan=4, sticky=tkinter.S, ipadx=5, ipady=5, padx=5, pady=5)

        # frame estrai
        self.extraction_frame = ctk.CTkFrame(self.menu.tab('Estrai'), fg_color='transparent')
        self.extraction_frame.pack(expand=True, fill=tkinter.BOTH)

        self.type_frame = ctk.CTkFrame(self.extraction_frame, fg_color='transparent')
        self.type_frame.pack(fill=tkinter.X)
        self.type_frame.columnconfigure(0, weight=5)
        self.type_frame.columnconfigure(1, weight=5)
        self.type_frame.columnconfigure(2, weight=5)
        self.type_frame.rowconfigure(0, weight=1)

        self.lbl_type = ctk.CTkLabel(self.type_frame, text='Seleziona il tipo di dato da estrarre')
        self.lbl_type.grid(row=0, column=0, columnspan=3)

        self.btn_text = ctk.CTkButton(self.type_frame, text='Testo', state='disabled',
                                      command=self.set_text)
        self.btn_text.grid(row=1, column=0, ipadx=5, ipady=5, padx=5, pady=5)

        self.btn_image = ctk.CTkButton(self.type_frame, text='Immagini', state='disabled',
                                       command=self.set_image)

        self.btn_image.grid(row=1, column=1, ipadx=5, ipady=5, padx=5, pady=5)

        self.btn_table = ctk.CTkButton(self.type_frame, text='Tabelle', state='disabled',
                                       command=self.set_table)

        self.btn_table.grid(row=1, column=2, ipadx=5, ipady=5, padx=5, pady=5)

        self.pagination_frame = ctk.CTkFrame(self.extraction_frame, fg_color='transparent')
        self.pagination_frame.pack(expand=True, fill=tkinter.BOTH, side=tkinter.BOTTOM)
        self.pagination_frame.rowconfigure(0, weight=1)
        self.pagination_frame.columnconfigure(0, weight=5)
        self.pagination_frame.columnconfigure(1, weight=5)

        self.lbl_page = ctk.CTkLabel(self.pagination_frame, text='Numero pagina: ')
        self.lbl_page.grid(row=0, column=0, columnspan=2, sticky=tkinter.S)
        self.btn_minus = ctk.CTkButton(self.pagination_frame, text='<--', command=self.minus)
        self.btn_minus.grid(row=1, column=0, ipadx=5, ipady=5, padx=5, pady=5)
        self.btn_plus = ctk.CTkButton(self.pagination_frame, text='-->', command=self.plus)
        self.btn_plus.grid(row=1, column=1, ipadx=5, ipady=5, padx=5, pady=5)

        # run
        self.view_image()
        self.get_rules()
        self.mainloop()

    def open_file(self):
        path = filedialog.askopenfilename(title="Seleziona il PDF da aprire",
                                          filetypes=[('PDF', '*.pdf')])
        if path != '':
            self.pdf_path.set(path)
            file = path.split('/')
            self.selected_file.set(file[len(file) - 1])
            self.extractor = PDF(self.pdf_path.get())
            self.number_of_page = self.extractor.get_page_number()
            self.width = self.extractor.width
            self.height = self.extractor.height
            self.txt = []
            self.img = []
            self.tab = []
            self.ask_add_coord = True
            self.coord.clear()
            self.image_counter = 0
            self.page = 0
            self.lbl_page.configure(text='Numero pagina: ' + str(self.page + 1)
                                         + ' di ' + str(self.number_of_page + 1))
            self.image_canvas = convert_from_path(pdf_path=self.pdf_path.get(),
                                                  poppler_path=poppler_path)
            self.view_image()
            self.combobox_rules.set('')
            self.btn_text.configure(state='normal')
            self.btn_image.configure(state='normal')
            self.btn_table.configure(state='normal')

    def open_others_files(self):
        if len(self.coord) != 0:
            if ((len(self.txt) != 0 or len(self.img) != 0 or len(self.tab) != 0) and
               messagebox.askyesno(title="Salvare",
                                   message="Memorizzare i dati appena estratti?")):
                self.local_memorizzation()
            files = filedialog.askopenfilenames(title="Seleziona i PDF da aprire",
                                                filetypes=[('PDF', '*.pdf')])
            if len(files) != 0:
                self.ask_add_coord = False
                for f in files:
                    self.pdf_path.set(f)
                    file = f.split('/')
                    self.selected_file.set(file[len(file) - 1])
                    self.extractor = PDF(self.pdf_path.get())
                    self.image_counter = 0
                    for t in self.coord:
                        coordinates = self.coord[t]
                        for c in coordinates:
                            self.page = c[4]
                            if t == Const.TEXT:
                                self.get_t(c[0], c[1], c[2], c[3], self.page)
                            elif t == Const.IMG:
                                self.get_image(c[0], c[1], c[2], c[3], self.page)
                            elif t == Const.TAB:
                                self.get_table(c[0], c[1], c[2], c[3], self.page)
                            else:
                                raise Exception("Errore")
                    self.local_memorizzation()
                    self.extractor.__del__()
                messagebox.showinfo(title="OK",
                                    message="Estrazione avvenuta con successo !\n " +
                                    "Definire nuove regole, applicare le regole ad altri pdf " +
                                    "o salvare i dati estratti")
                self.selected_file.set('Nessun file PDF selezionato')
                self.pdf_path.set('')
                self.output_image.delete('all')
        else:
            messagebox.showerror(title="Errore", message="Nessuna regola di estrazione selezionata")

    def view_image(self):
        if self.pdf_path.get() != '':
            img = self.image_canvas[self.page]
            img.save(fp='./img.png', format='png')
            pic = Image.open("./img.png").resize((int(self.width), int(self.height)), Image.LANCZOS)
            output = ImageTk.PhotoImage(pic)
            self.output_image.img = output
            self.output_image.create_image(0, 0, image=output, anchor='nw')

            self.rect_id = self.output_image.create_rectangle(self.topx, self.topy, self.topx,
                                                              self.topy, fill='', outline='red', width=2)

            self.output_image.bind('<B1-Motion>', self.view_rect)
            self.output_image.bind('<Button-1>', self.get_init_coord)
            self.output_image.bind('<ButtonRelease-1>', self.get_rect)

        if self.page == 0:
            self.btn_minus.configure(state="disabled")

        if self.btn_minus.cget('state') == "disabled" and self.page > 0:
            self.btn_minus.configure(state="normal")

        if self.page == self.number_of_page:
            self.btn_plus.configure(state="disabled")

        if self.btn_plus.cget('state') == "disabled" and self.page < self.number_of_page:
            self.btn_plus.configure(state="normal")

    def save(self):
        self.local_memorizzation()
        if len(self.obj) != 0:
            ok = messagebox.askyesno(title="Salvare?", message="Confermi di salvare i dati estratti?")
            if ok:
                if self.destination_directory.get() == '' and len(self.image_to_save) != 0:
                    self.choose_dest_dir()
                    if self.destination_directory.get() != '':
                        self.insert_database()
                else:
                    self.insert_database()
        else:
            messagebox.showerror(title="Errore",
                                 message="Nessun dato memorizzato\n " +
                                 "Estrai dati prima di poterli salvare")

    def get_init_coord(self, event):
        self.topx = event.x
        self.topy = event.y

    def get_rect(self, event):
        self.botx = event.x
        self.boty = event.y

        if not_point(self.topx, self.topy, self.botx, self.boty) and not self.extractor.is_closed():
            if self.type.get() == Const.TEXT:
                self.get_t(self.topx, self.topy, self.botx, self.boty, self.page)
            elif self.type.get() == Const.IMG:
                self.get_image(self.topx, self.topy, self.botx, self.boty, self.page)
            elif self.type.get() == Const.TAB:
                self.get_table(self.topx, self.topy, self.botx, self.boty, self.page)
            else:
                messagebox.showerror(title="Errore", message="Selezionare il tipo da estrarre")

    def view_rect(self, event):
        self.botx = event.x
        self.boty = event.y
        self.output_image.coords(self.rect_id, self.topx, self.topy, self.botx, self.boty)

    def get_t(self, x0, y0, x1, y1, p):
        data = self.extractor.get_text(x0, y0, x1, y1, p)
        if data:
            if self.ask_add_coord:
                ok = self.add_coord(x0, y0, x1, y1, data, Const.TEXT)
                if ok:
                    self.txt.append(data)
            else:
                self.txt.append(data)
        else:
            if self.ask_add_coord:
                messagebox.showerror(title='Errore', message="Nessun elemento testuale selezionato")

    def get_image(self, x0, y0, x1, y1, p):
        img = self.extractor.get_img(x0, y0, x1, y1, p)
        data_img = img[0]
        ext = img[1]
        if data_img and ext:
            if self.ask_add_coord:
                ok = self.add_coord(x0, y0, x1, y1, "Aggiungere l'immagine selezionata?", Const.IMG)
                if ok:
                    name = self.save_img(data_img, ext)
                    self.img.append(name)
            else:
                name = self.save_img(data_img, ext)
                self.img.append(name)
        else:
            if self.ask_add_coord:
                messagebox.showerror(title="Errore", message="Nessun immagine selezionata")

    def get_table(self, x0, y0, x1, y1, p):
        data = self.extractor.get_tab(x0, y0, x1, y1, p)
        if len(data) != 0:
            if self.ask_add_coord:
                ok = self.add_coord(x0, y0, x1, y1, "Aggiungere la tabella selezionata?", Const.TAB)
                if ok:
                    self.tab.append(data)
            else:
                self.tab.append(data)
        else:
            if self.ask_add_coord:
                messagebox.showerror(title="Errore", message="Nessuna tabella selezionata")

    def set_rule(self, rule):
        self.ask_add_coord = False
        coordinates = self.all_rules[rule]
        for c in coordinates:
            self.coord[c] = coordinates[c]

    def get_rules(self):
        self.rules.clear()
        rules = Database().get_rules()
        for rule in rules:
            for name in rule:
                self.rules.append(name)
                self.all_rules[name] = rule[name]
        self.combobox_rules.configure(values=self.rules)

    def save_rule(self):
        if (len(self.txt) != 0 or len(self.img) != 0 or len(self.tab) != 0) and self.rule_name.get() != '':
            if messagebox.askyesno(title="Salvare?", message="Memorizzare i dati estratti?"):
                self.local_memorizzation()
            self.ask_add_coord = False
            self.output_image.delete('all')
            self.extractor.__del__()
            new_rule = {self.rule_name.get(): self.coord}
            insert = Database().insert_rule(new_rule)
            if insert:
                self.get_rules()
                messagebox.showinfo(title="Successo", message="Regola salvata con successo")
            else:
                messagebox.showerror(title="Errore", message="Errore nel salvataggio della regola")
        elif len(self.txt) == 0 and len(self.img) == 0 and len(self.tab) == 0 and self.pdf_path.get() == '':
            messagebox.showerror(title="Errore", message="Selezionare un PDF e definire la regola")
        elif self.rule_name.get() == '':
            messagebox.showerror(title="Errore", message="Associare un nome alla regola da memorizzare")
        else:
            messagebox.showerror(title="Errore", message="Nessuna regola memorizzata")

    def local_memorizzation(self):
        if len(self.txt) != 0 or len(self.img) != 0 or len(self.tab) != 0:
            d = {Const.TEXT: self.txt, Const.IMAGE: self.img, Const.TABLE: self.tab}
            self.obj[self.selected_file.get()] = d
            self.txt, self.img, self.tab = [], [], []

    def add_coord(self, x0, y0, x1, y1, data, tipo):
        ok = messagebox.askyesno(title="Aggiungere?", message=data)
        if ok:
            if self.coord.get(tipo):
                self.coord.get(tipo).append((x0, y0, x1, y1, self.page))
            else:
                self.coord[tipo] = []
                self.coord.get(tipo).append((x0, y0, x1, y1, self.page))
            return True
        else:
            return False

    def save_img(self, img_data, ext):
        self.image_counter += 1
        folder_name = self.selected_file.get().split('.')
        name = (folder_name[0] + '/image' + str(self.image_counter) +
                '.' + str(ext))
        file_name = 'image' + str(self.image_counter) + '.' + str(ext)
        self.image_to_save.append((file_name, folder_name[0], img_data))
        return name

    def insert_database(self):
        for t in self.obj:
            json = {t: self.obj.get(t)}
            temp = []
            image_arr = json.get(t).get(Const.IMAGE)
            for i in image_arr:
                temp.append(self.destination_directory.get() + '/' + i)
            json.get(t)[Const.IMAGE] = temp
            insert = Database().insert_data(json)
            if not insert:
                messagebox.showerror(title="Errore", message="Errore nel salvataggio dei dati")
        for img in self.image_to_save:
            self.save_images(img[0], img[1], img[2])
        messagebox.showinfo(title="Successo", message="Salvataggio avvenuto con successo")
        self.obj.clear()
        self.image_to_save.clear()

    def choose_dest_dir(self):
        directory = filedialog.askdirectory(title="Seleziona cartella",
                                            initialdir="C:/Users/galli/OneDrive/Desktop")
        if directory != '':
            self.destination_directory.set(directory)

    def save_images(self, name, directory, img_data):
        os.chdir(self.destination_directory.get())
        if not os.path.exists('./' + directory):
            os.mkdir('./' + directory)
        os.chdir('./' + directory)
        img_save = PIL.Image.open(io.BytesIO(img_data))
        img_save.save(open(name, 'wb'))
        os.chdir('../../')

    def plus(self):
        if self.page < self.number_of_page:
            self.page += 1
            self.lbl_page.configure(text='Numero pagina: ' + str(self.page + 1)
                                         + ' di ' + str(self.number_of_page + 1))
            self.view_image()

    def minus(self):
        if self.page > 0:
            self.page -= 1
            self.lbl_page.configure(text='Numero pagina: ' + str(self.page + 1)
                                         + ' di ' + str(self.number_of_page + 1))
            self.view_image()

    def clear(self):
        self.txt.clear()
        self.img.clear()
        self.tab.clear()
        self.obj.clear()
        self.coord.clear()
        self.ask_add_coord = True
        self.image_counter = 0
        self.page = 0
        self.type.set('')
        self.image_to_save.clear()
        self.output_image.delete('all')
        self.destination_directory.set('')
        self.height = 0
        self.width = 0
        self.number_of_page = 0
        self.selected_file.set('Nessun file PDF selezionato')
        self.pdf_path.set('')
        self.lbl_page.configure(text='Numero pagina: ')
        self.combobox_rules.set('')
        self.rule_name.set('')
        self.selected_rule.set('')
        self.rules.clear()
        self.get_rules()
        if self.extractor is not None:
            self.extractor.__del__()

    def set_text(self):
        self.type.set('text')
        messagebox.showinfo(title="Selezione", message="TESTO selezionato !")

    def set_image(self):
        if not self.extractor.is_closed() and self.extractor.exists_img(self.page):
            self.type.set('img')
            messagebox.showinfo(title="Selezione", message="IMMAGINI selezionato !")
        else:
            self.type.set('')

    def set_table(self):
        if not self.extractor.is_closed() and self.extractor.exists_tab(self.page):
            self.type.set('tab')
            messagebox.showinfo(title="Selezione", message="TABELLE selezionato !")
        else:
            self.type.set('')

    def close_root_window(self):
        if os.path.exists('./img.png'):
            os.remove('./img.png')
        self.destroy()
