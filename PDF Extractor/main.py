import tkinter as tk
from tkinter import filedialog, messagebox
import os
import tempfile
from pdf_extractor import extract_all_text_from_pdf
from data_processor import organize_pdf_content

class PDFProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Processador de PDF")
        master.geometry("550x200")

        # Caminho inicial padrão
        self.default_path = r"Informe o caminho de onde os arquivos PDF costumam ser salvos"

        # Variáveis para armazenar os caminhos
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()

        # Configuração do grid
        master.grid_columnconfigure(1, weight=1)

        # Widgets para o PDF de entrada
        tk.Label(master, text="Arquivo PDF:", font=('Arial', 10)).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.entry_pdf = tk.Entry(master, textvariable=self.pdf_path, width=60, font=('Arial', 10))
        self.entry_pdf.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        tk.Button(master, text="Procurar", command=self.browse_pdf, width=10).grid(row=0, column=2, padx=5, pady=5)

        # Widgets para a pasta de saída
        tk.Label(master, text="Pasta de Saída:", font=('Arial', 10)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_output = tk.Entry(master, textvariable=self.output_dir, width=60, font=('Arial', 10))
        self.entry_output.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        tk.Button(master, text="Procurar", command=self.browse_output, width=10).grid(row=1, column=2, padx=5, pady=5)

        # Botão de execução
        self.btn_run = tk.Button(master, text="Executar Processamento", command=self.run_processing, 
                               bg="#4CAF50", fg="white", font=('Arial', 10, 'bold'), width=20)
        self.btn_run.grid(row=2, column=1, pady=20)

        # Status
        self.status_label = tk.Label(master, text="Pronto para iniciar", fg="gray", font=('Arial', 10))
        self.status_label.grid(row=3, column=1)

    def browse_pdf(self):
        """Abre o diálogo de seleção de arquivo com o caminho pré-definido"""
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo PDF",
            initialdir=self.default_path,  # Define o diretório inicial
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)

    def browse_output(self):
        """Abre o diálogo de seleção de pasta"""
        dir_path = filedialog.askdirectory(
            title="Selecione a pasta de saída"
        )
        if dir_path:
            self.output_dir.set(dir_path)

    def run_processing(self):
        """Executa o processamento completo"""
        if not self.pdf_path.get():
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo PDF!")
            return
            
        if not self.output_dir.get():
            messagebox.showwarning("Aviso", "Por favor, selecione uma pasta de saída!")
            return

        # Cria arquivo temporário
        with tempfile.NamedTemporaryFile(
            suffix=".xlsx", 
            delete=False,
            mode='w+b'
        ) as tmpfile:
            temp_excel = tmpfile.name

        try:
            self.status_label.config(text="Processando PDF...", fg="blue")
            self.master.update_idletasks()

            # Etapa 1: Extração do texto do PDF
            if not extract_all_text_from_pdf(self.pdf_path.get(), temp_excel):
                messagebox.showerror("Erro", "Falha na extração do conteúdo do PDF")
                return

            # Etapa 2: Processamento dos dados
            base_name = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]
            output_filename = os.path.join(
                self.output_dir.get(),
                f"{base_name}-FORMATADO.xlsx"
            )

            if not organize_pdf_content(temp_excel, output_filename):
                messagebox.showerror("Erro", "Falha no processamento dos dados")
                return

            self.status_label.config(text="Processamento concluído!", fg="green")
            messagebox.showinfo(
                "Sucesso", 
                f"Arquivo gerado com sucesso!\n\nLocal: {output_filename}",
                parent=self.master
            )

        except Exception as e:
            messagebox.showerror(
                "Erro", 
                f"Ocorreu um erro durante o processamento:\n\n{str(e)}",
                parent=self.master
            )
            self.status_label.config(text="Erro no processamento", fg="red")
        
        finally:
            # Limpeza do arquivo temporário
            if os.path.exists(temp_excel):
                try:
                    os.remove(temp_excel)
                except Exception as e:
                    print(f"Erro ao remover arquivo temporário: {str(e)}")
            
            self.status_label.config(text="Pronto para novo processamento", fg="gray")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()