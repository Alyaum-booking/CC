import tkinter as tk
from tkinter import messagebox, ttk
import hashlib
import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
import arabic_reshaper
from bidi.algorithm import get_display

DATA_FILE = "cards.xlsx"
OUTPUT_FILE = "powerbi_data.xlsx"
CARD_INFO_FILE = "CDB.xlsx"

plt.rcParams["font.family"] = "Tahoma"


def ar(text):
    text = str(text)
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


database = {
    "balyaqub": hash_password("Alyaum@123"),
    "aalomari": hash_password("Svmg#152")
}

attempts = 3


def login():
    global attempts

    username = entry_user.get()
    password = hash_password(entry_pass.get())

    if username in database and database[username] == password:

        messagebox.showinfo("Success", "Login Successful")

        login_window.destroy()

        root = tk.Tk()
        app = CardApp(root)
        root.mainloop()

    else:
        attempts -= 1
        messagebox.showerror("Error", f"Wrong data\nAttempts left: {attempts}")

        if attempts == 0:
            login_btn.config(state="disabled")
            messagebox.showwarning("Locked", "System Locked")

DATA_FILE = "cards.xlsx"
OUTPUT_FILE = "powerbi_data.xlsx"
CARD_INFO_FILE = "CDB.xlsx"

plt.rcParams["font.family"] = "Tahoma"


def ar(text):
    text = str(text)
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


class CardApp:

    def __init__(self, root):

        self.root = root
        self.root.title("Card Selection System")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        self.style = ttk.Style()
        self.style.theme_use("clam")

        if not os.path.exists(DATA_FILE):
            messagebox.showerror("Error", "cards.xlsx not found")
            root.destroy()
            return

        self.data = self.load_data()
        self.build_ui()

    def load_data(self):
        df = pd.read_excel(DATA_FILE)
        data = {}

        for _, row in df.iterrows():
            letter = str(row["letter"])
            number = str(row["number"])

            if letter not in data:
                data[letter] = []

            data[letter].append(number)

        for letter in data:
            data[letter] = sorted(list(set(data[letter])))

        return data

    def build_ui(self):

        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        title = tk.Label(frame, text="Card Selection", font=("Arial", 18, "bold"))
        title.pack(pady=10)

        tk.Label(frame, text="Select Card Letter").pack()

        self.letter_var = tk.StringVar()

        self.letter_combo = ttk.Combobox(
            frame,
            textvariable=self.letter_var,
            values=list(self.data.keys()),
            state="readonly"
        )

        self.letter_combo.pack(pady=5)
        self.letter_combo.bind("<<ComboboxSelected>>", self.update_numbers)

        tk.Label(frame, text="Select Card Number").pack()

        self.number_var = tk.StringVar()

        self.number_combo = ttk.Combobox(
            frame,
            textvariable=self.number_var,
            state="readonly"
        )

        self.number_combo.pack(pady=5)

        ttk.Button(frame, text="Submit", command=self.submit).pack(pady=10)
        ttk.Button(frame, text="Reset", command=self.reset).pack()

    def update_numbers(self, event):

        letter = self.letter_var.get()
        numbers = self.data.get(letter, [])

        self.number_combo["values"] = numbers
        self.number_var.set("")

    def save_to_excel(self, letter, number):

        new_data = pd.DataFrame([{
            "letter": letter,
            "number": number,
            "timestamp": datetime.now()
        }])

        if os.path.exists(OUTPUT_FILE):

            df = pd.read_excel(OUTPUT_FILE)
            df = pd.concat([df, new_data], ignore_index=True)

        else:

            df = new_data

        df.to_excel(OUTPUT_FILE, index=False)

    def submit(self):

        letter = self.letter_var.get()
        number = self.number_var.get()

        if not letter or not number:
            messagebox.showwarning("Warning", "Please select letter and number")
            return

        self.save_to_excel(letter, number)

        messagebox.showinfo(
            "Success",
            f"Saved\nLetter: {letter}\nNumber: {number}"
        )

        self.show_card_charts(letter, number)
        self.reset()

    def show_card_charts(self, letter, number):

        if not os.path.exists(CARD_INFO_FILE):
            messagebox.showerror("Error", f"{CARD_INFO_FILE} not found")
            return

        df = pd.read_excel(CARD_INFO_FILE)

        card_data = df[
            (df["letter"] == letter) &
            (df["number"] == int(number))
        ]

        if card_data.empty:
            messagebox.showwarning("Not Found", "Card not found in analysis file")
            return

        count = len(card_data)

        dept_counts = card_data["endins"].value_counts()
        cost_counts = card_data["costins"].value_counts()
        status_counts = card_data["status"].value_counts()
        state_counts = card_data["insstate"].value_counts()
        dep_counts = card_data["dep"].value_counts()
        fee_counts = card_data["feeofnewlen"].value_counts()
        fin_counts = card_data["fineofnolen"].value_counts()
        div_counts = card_data["driver"].value_counts()
        lin_counts = card_data["endlin"].value_counts()
        brandcar=card_data["brand"].value_counts()
        carnamee=card_data["carname"].value_counts()

        cards = []

        cards.append(("عدد السيارات", "حسب القيمة المختارة", count))

        for dept, val in dept_counts.items():
            cards.append(("نهاية الفحص", dept, val))

        for st, val in cost_counts.items():
            cards.append(("تكلفة الفحص", st, val))

        for su, val  in status_counts.items():
            cards.append(("الحالة", su, val))

        for sa, val  in state_counts.items():
            cards.append(("حالة الفحص", sa, val))

        for sd, val  in dep_counts.items():
            cards.append(("القسم", sd, val))

        for sd, val in fee_counts.items():
            cards.append(("رسوم تجديد الاستمارة", sd, val))

        for sd, val in fin_counts.items():
            cards.append(("غرامة عدم تجديد الاستمارة", sd, val))

        for sd, val in div_counts.items():
            cards.append(("اسم السائق", sd, val))

        for sd, val in lin_counts.items():
            cards.append(("تاريخ انتهاء الاستمارة", sd, val))

        for brandcar, val in brandcar.items():
            cards.append(("الماركة", brandcar, val))

        for carnamee, val in carnamee.items():
            cards.append(("الطراز", carnamee, val))

        plt.figure(figsize=(10, 5))
        plt.axis("off")

        cols = 3
        card_w = 1 / cols
        card_h = 0.25

        for i, (title, name, value) in enumerate(cards):

            row = i // cols
            col = i % cols

            x = col * card_w
            y = 1 - (row + 1) * card_h

            rect = plt.Rectangle(
                (x, y),
                card_w - 0.01,
                card_h - 0.01,
                edgecolor="black",
                facecolor="lightblue"
            )

            plt.gca().add_patch(rect)

            plt.text(
                x + 0.02,
                y + card_h - 0.08,
                ar(title),
                fontsize=10,
                weight="bold"
            )

            plt.text(
                x + 0.02,
                y + card_h - 0.15,
                ar(name),
                fontsize=9
            )

            plt.text(
                x + 0.02,
                y + 0.05,
                str(value),
                fontsize=14
            )

        plt.title(ar("لوحة معلومات السيارة"))
        plt.show()

    def reset(self):

        self.letter_var.set("")
        self.number_var.set("")
        self.number_combo["values"] = []


if __name__ == "__main__":

    login_window = tk.Tk()

    login_window.title("Secure Login System")
    login_window.geometry("400x300")

    title = tk.Label(login_window, text="Secure Login", font=("Arial", 18))
    title.pack(pady=20)

    tk.Label(login_window, text="Username", font=("Arial", 14)).pack()

    entry_user = tk.Entry(login_window)
    entry_user.pack(pady=5)

    tk.Label(login_window, text="Password", font=("Arial", 14)).pack()

    entry_pass = tk.Entry(login_window, show="*")
    entry_pass.pack(pady=5)

    login_btn = tk.Button(login_window, text="Login", command=login, font=("Arial", 14))
    login_btn.pack(pady=20)

    login_window.mainloop()