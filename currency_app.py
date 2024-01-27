import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, filedialog, messagebox 
import webbrowser
import pandas as pd

class CurrencyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Currency Rates")
        self.root.geometry("1550x1000")

        self.create_widgets()
        self.fetch_and_update()

    def create_widgets(self):
        self.tree = ttk.Treeview(self.root, columns=("Currency", "Last Update", "Current Rate", "Min Rate", "Max Rate"), show="headings")
        self.tree.heading("Currency", text="Currency")
        self.tree.heading("Last Update", text="Last Update")
        self.tree.heading("Current Rate", text="Current Rate")
        self.tree.heading("Min Rate", text="Min Rate")
        self.tree.heading("Max Rate", text="Max Rate")

        self.tree.column("Currency", width=300)
        self.tree.column("Last Update", width=300)
        self.tree.column("Current Rate", width=300)
        self.tree.column("Min Rate", width=300)
        self.tree.column("Max Rate", width=300)

        self.tree.tag_configure("even", background="#f0f0f0")  # Light gray
        self.tree.tag_configure("odd", background="#ffffff")   # White

        self.tree.grid(row=0, column=0, sticky="nsew")

        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        save_button = tk.Button(self.root, text="Save to Excel", command=self.save_to_excel)
        save_button.grid(row=1, column=0, sticky="w")

        self.tree.bind("<Double-1>", self.open_main_link)

    def open_main_link(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            selected_currency = self.tree.item(selected_item, "values")[0]
            main_link = f"https://www.tgju.org/profile/{selected_currency}"
            webbrowser.open(main_link)

    def save_to_excel(self):
        currency_rates = self.fetch_currency_rates()
        if not currency_rates:
            return

        df = pd.DataFrame(currency_rates)
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        df.to_excel(file_path, index=False)
        messagebox.showinfo("Save to Excel", f"Data saved to {file_path}")

    def fetch_currency_rates(self):
        try:
            response = requests.get("https://www.tgju.org/currency")
            response.raise_for_status()

            soup = BeautifulSoup(response.content, "html.parser")
            currency_rows = soup.select("tbody tr")

            currency_rates = []
            for row in currency_rows:
                currency_name_slug = row.get("data-market-nameslug", "N/A")
                currency_data = [td.text for td in row.select("td")]

                if len(currency_data) >= 5:
                    today = currency_data[4]
                    current_rate = currency_data[0]
                    min_rate = currency_data[2]
                    max_rate = currency_data[3]
                    currency_rates.append({
                        "currency_name_slug": currency_name_slug,
                        "today": today,
                        "current_rate": current_rate,
                        "min_rate": min_rate,
                        "max_rate": max_rate
                    })

            return currency_rates

        except requests.RequestException as e:
            print(f"Error fetching data: {e}")
            return []

    def update_gui(self, currency_rates):
        for row in self.tree.get_children():
            self.tree.delete(row)

        for i, currency in enumerate(currency_rates):
            currency_name_slug = currency["currency_name_slug"]
            today = currency["today"]
            current_rate = currency["current_rate"]
            min_rate = currency["min_rate"]
            max_rate = currency["max_rate"]

            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(currency_name_slug, today, current_rate, min_rate, max_rate), tags=(tag,))

        tree_height = len(currency_rates)
        self.tree.configure(height=tree_height)

    def fetch_and_update(self):
        currency_rates = self.fetch_currency_rates()
        self.update_gui(currency_rates)

        self.root.after(600000, self.fetch_and_update)

def main():
    root = tk.Tk()
    app = CurrencyApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
