import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook
import pathlib

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()
        self.appearance()
        self.todo_sistema()

    def layout_config(self):
        self.title("Sistema de Gestão de Clientes")
        self.geometry("700x500")

    def appearance(self):
        self.lb_apm = ctk.CTkLabel(self, text="Tema", bg_color="transparent", text_color=['#000', "#fff"])
        self.lb_apm.place(x=50, y=430)

        self.opt_apm = ctk.CTkOptionMenu(self, values=["Light", "Dark", "System"], command=self.change_apm)
        self.opt_apm.place(x=50, y=460)

    def todo_sistema(self):
        frame = ctk.CTkFrame(self, width=700, height=50, corner_radius=0, bg_color="teal", fg_color="teal")
        frame.place(x=0, y=10)

        title = ctk.CTkLabel(frame, text="Sistema de Gestão de Clientes", font=("Century Gothic bold", 24), text_color="#fff")
        title.place(x=190, y=10)

        span = ctk.CTkLabel(self, text="Por favor, preencha todos os campos do formulário!", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        span.place(x=50, y=70)

        ficheiro_path = pathlib.Path("Clientes.xlsx")

        if ficheiro_path.exists():
            pass
        else:
            workbook = Workbook()
            folha = workbook.active
            folha['A1'] = "Nome completo"
            folha['B1'] = "Contacto"
            folha["C1"] = "Idade"
            folha["D1"] = "Genero"
            folha["E1"] = "Endereço"
            folha["F1"] = "Observações"

            workbook.save(ficheiro_path)

        def submit():
            # Pegar dados do entry
            name = name_value.get()
            contact = Contact_value.get()
            age = age_value.get()
            gender = gender_combobox.get()
            address = address_value.get()
            obs = obs_entry.get(1.0, END)

            ficheiro = openpyxl.load_workbook('Clientes.xlsx')
            folha = ficheiro.active
            folha.append([name, contact, age, gender, address, obs])

            ficheiro.save("Clientes.xlsx")
            messagebox.showinfo("Sistema", "Dados salvos com sucesso!")

        def clear():
            name_value.set("")
            Contact_value.set("")
            age_value.set("")
            address_value.set("")
            obs_entry.delete(1.0, END)

        # Texts variables
        name_value = StringVar()
        Contact_value = StringVar()
        age_value = StringVar()
        address_value = StringVar()

        # Entries
        name_entry = ctk.CTkEntry(self, width=350, textvariable=name_value, font=("Century Gothic bold", 16), fg_color="transparent")
        Contact_entry = ctk.CTkEntry(self, width=200, textvariable=Contact_value, font=("Century Gothic bold", 16), fg_color="transparent")
        age_entry = ctk.CTkEntry(self, width=150, textvariable=age_value, font=("Century Gothic bold", 16), fg_color="transparent")
        address_entry = ctk.CTkEntry(self, width=200, textvariable=address_value, font=("Century Gothic bold", 16), fg_color="transparent")

        # Combobox
        gender_combobox = ctk.CTkComboBox(self, values=["Masculino", "Feminino"], font=("Century Gothic bold", 14), width=150)
        gender_combobox.set("Masculino")

        # Entrada de obs
        obs_entry = ctk.CTkTextbox(self, height=120, width=500, font=("arial", 18), border_color="#aaa", border_width=2, fg_color="transparent")

        # Labels
        lb_name = ctk.CTkLabel(self, text="Nome Completo", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        lb_contact = ctk.CTkLabel(self, text="Contato", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        lb_age = ctk.CTkLabel(self, text="Idade", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        lb_gender = ctk.CTkLabel(self, text="Gênero", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        lb_address = ctk.CTkLabel(self, text="Endereço", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])
        lb_obs = ctk.CTkLabel(self, text="Observações", font=("Century Gothic bold", 16), text_color=["#000", "#fff"])

        btn_submit = ctk.CTkButton(self, text="Salvar dados".upper(), command=submit, fg_color="#151", hover_color="#131")
        btn_submit.place(x=300, y=420)

        btn_clear = ctk.CTkButton(self, text="Limpar Campos".upper(), command=clear, fg_color="#555", hover_color="#333")
        btn_clear.place(x=500, y=420)

        # Posicionando os elementos na janela
        lb_name.place(x=50, y=120)
        name_entry.place(x=50, y=150)

        lb_contact.place(x=450, y=120)
        Contact_entry.place(x=450, y=150)

        lb_age.place(x=300, y=190)
        age_entry.place(x=300, y=220)

        lb_gender.place(x=500, y=190)
        gender_combobox.place(x=500, y=220)

        lb_address.place(x=50, y=190)
        address_entry.place(x=50, y=220)

        lb_obs.place(x=50, y=260)
        obs_entry.place(x=150, y=260)

    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)

if __name__ == "__main__":
    app = App()
    app.mainloop()
