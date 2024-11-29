import requests
from openpyxl import Workbook
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

class WeatherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta e Exportação de Clima")
        self.root.geometry("400x250")

        self.create_widgets()

    def create_widgets(self):
        self.city_label = tk.Label(self.root, text="Digite o nome da cidade:")
        self.city_label.pack(pady=10)
        
        self.city_entry = tk.Entry(self.root, width=30)
        self.city_entry.pack(pady=5)
        
        self.submit_button = tk.Button(self.root, text="Consultar e Exportar", command=self.submit)
        self.submit_button.pack(pady=10)
        
        self.progress = ttk.Progressbar(self.root, orient='horizontal', mode='indeterminate')
        self.result_label = tk.Label(self.root, text="", fg="green")
        self.result_label.pack(pady=10)
        
        self.progress.pack(pady=10)

    def get_weather(self, city):
        API_KEY = "77e851e373a0b4f6c9d73012e01e1d8a"
        url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_KEY}&units=metric"

        try:
            response = requests.get(url)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"An error occurred while fetching weather data for {city}: {e}")
            return None

    def export_to_excel(self, city):
        weather_data = self.get_weather(city)
        if weather_data:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(['Cidade', 'Descrição', 'Temperatura (°C)'])

            city_name = city
            description = weather_data["weather"][0]["description"]
            temperature = weather_data["main"]["temp"]
            sheet.append([city_name, description, temperature])

            filename = f"clima_{city_name}.xlsx"
            workbook.save(filename)
            return filename
        return None

    def submit(self):
        city = self.city_entry.get()
        if city:
            self.progress.start()
            filename = self.export_to_excel(city)
            self.progress.stop()
            if filename:
                self.result_label.config(text=f"Dados do clima salvos em {filename}", fg="green")
                messagebox.showinfo("Sucesso", f"Dados do clima salvos em {filename}")
            else:
                self.result_label.config(text="Não foi possível encontrar dados para a cidade informada.", fg="red")
                messagebox.showerror("Erro", "Não foi possível encontrar dados para a cidade informada.")
        else:
            self.result_label.config(text="Por favor, digite o nome da cidade.", fg="orange")
            messagebox.showwarning("Aviso", "Por favor, digite o nome da cidade.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WeatherApp(root)
    root.mainloop()
