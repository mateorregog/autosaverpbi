import os
import time
import pyautogui
import schedule
import pygetwindow as gw
import psutil
import datetime
import tkinter as tk
from tkinter import messagebox
welcome_message = """
        ¡Bienvenido a AutoSaver for Power BI!
        
        Programa desarrollado por Matt. O.
        Cualquier duda o comentario escribir al correo mateorrego@gmail.com.
        Espero sea útil para su trabajo con PowerBI :)
        
        A continuación tendrá que ingresar el nombre del archivo .pbix que quiere autoguardar de \nforma automática y el intervalo de ejecución.
        
        Tenga en cuenta las siguientes recomendaciones:
        - Este ejecutable debe estar ubicado en la misma carpeta del archivo que quiere autoguardar
        - Asegúrese de que el nombre ingresado sea igual al archivo que quiere autoguardar
        - No incluya en el nombre la extensión .pbix
        - No incluya comillas
        - No cierre esta ventana si quiere que el autoguardado siga ejecutándose
        - Es posible que deba ejecutar como administrador el .exe para que pueda guardar con éxito los cambios
        __________________________________________________________________________
        """

class AutoSaverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoSaver for Power BI")

        self.frame = tk.Frame(root)
        self.frame.pack(padx=10, pady=10)

        self.entry_label = tk.Label(self.frame, text="Nombre del archivo .pbix:")
        self.entry_label.grid(row=0, column=0, padx=5, pady=5)
        self.entry = tk.Entry(self.frame, width=40)
        self.entry.grid(row=0, column=1, padx=5, pady=5)

        self.time_label = tk.Label(self.frame, text="Intervalo de tiempo (minutos):")
        self.time_label.grid(row=1, column=0, padx=5, pady=5)
        self.time_entry = tk.Entry(self.frame, width=40)
        self.time_entry.grid(row=1, column=1, padx=5, pady=5)

        self.submit_button = tk.Button(self.frame, text="Submit", command=self.submit)
        self.submit_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

        self.log_text = tk.Text(self.frame, width=100, height=10, state='disabled')
        self.log_text.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

        self.log(welcome_message)
        #messagebox.showinfo("Bienvenido", welcome_message)
        

    def submit(self):
        
        nombre_archivo = self.entry.get()
        tiempo_tracking = int(self.time_entry.get())
        messagebox.showinfo("Confirmación", f"Nombre del archivo: {nombre_archivo}\nIntervalo de tiempo: {tiempo_tracking} minutos")
        self.guardar_archivo_power_bi(nombre_archivo)
        schedule.every(tiempo_tracking).minutes.do(self.programar_guardado, nombre_archivo)
        self.check_schedule()

    def guardar_archivo_power_bi(self, nombre_archivo):
        archivos_pbix = [archivo for archivo in os.listdir() if archivo.endswith('.pbix')]

        if archivos_pbix:
            archivo_power_bi = archivos_pbix[0]
            proceso_power_bi = [proceso for proceso in psutil.process_iter() if "PBIDesktop" in proceso.name()]
            if proceso_power_bi:
                ventana_power_bi = gw.getWindowsWithTitle(nombre_archivo)
                if ventana_power_bi:
                    try:
                        ventana_power_bi[0].activate()
                        time.sleep(2)
                        pyautogui.hotkey('ctrl', 's')
                        time.sleep(2)
                        self.log(f"Archivo guardado correctamente a las {datetime.datetime.now()}")
                    except Exception as e:
                        self.log(f"Error al activar la ventana: {e}")
                else:
                    self.log("No se pudo encontrar la ventana de Power BI.")
            else:
                self.log("Power BI no está en ejecución.")
        else:
            self.log("No se encontraron archivos .pbix en el directorio actual.")

    def programar_guardado(self, nombre_archivo):
        self.guardar_archivo_power_bi(nombre_archivo)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.config(state='disabled')

    def check_schedule(self):
        schedule.run_pending()
        self.root.after(1000, self.check_schedule)

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoSaverApp(root)
    root.mainloop()
    
