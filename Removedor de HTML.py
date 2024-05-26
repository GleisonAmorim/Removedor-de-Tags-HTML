# Importar as bibliotecas necessárias
import pandas as pd  # Para manipulação de dados em formato de tabela
from bs4 import BeautifulSoup  # Para lidar com HTML e XML de forma fácil
import tkinter as tk  # Para criar interfaces gráficas
from tkinter import filedialog, messagebox  # Para caixas de diálogo de seleção de arquivos e pastas, e mensagens

# Função para remover tags HTML
def remove_html_tags(text):
    # Verificar se o texto é uma string
    if isinstance(text, str):
        # Criar um objeto BeautifulSoup para analisar o texto HTML
        soup = BeautifulSoup(text, "html.parser")
        # Extrair o texto limpo sem as tags HTML
        return soup.get_text()
    else:
        # Se não for uma string, retornar o próprio valor
        return text

# Função para selecionar o arquivo Excel
def select_excel_file():
    # Abrir uma janela de seleção de arquivos Excel
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
    # Verificar se um arquivo foi selecionado
    if file_path:
        # Limpar e preencher o campo de entrada com o caminho do arquivo selecionado
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

# Função para selecionar a pasta de destino
def select_output_folder():
    # Abrir uma janela de seleção de pasta
    output_folder = filedialog.askdirectory()
    # Verificar se uma pasta foi selecionada
    if output_folder:
        # Limpar e preencher o campo de entrada com o caminho da pasta selecionada
        entry_output_folder.delete(0, tk.END)
        entry_output_folder.insert(0, output_folder)

# Função para processar a remoção de HTML
def process_html_removal():
    # Obter os caminhos do arquivo Excel e da pasta de destino
    file_path = entry_file_path.get()
    output_folder = entry_output_folder.get()

    # Verificar se um arquivo Excel foi selecionado
    if not file_path:
        # Exibir uma mensagem de erro se nenhum arquivo foi selecionado
        messagebox.showerror("Erro", "Selecione um arquivo Excel.")
        return
    # Verificar se uma pasta de destino foi selecionada
    if not output_folder:
        # Exibir uma mensagem de erro se nenhuma pasta foi selecionada
        messagebox.showerror("Erro", "Selecione uma pasta de destino.")
        return

    # Carregar a planilha do Excel
    df = pd.read_excel(file_path)

    # Aplicar a função para remover as tags HTML em todas as células do DataFrame
    # Verificar se o valor de cada célula é uma string antes de aplicar a função
    df = df.apply(lambda x: x.map(remove_html_tags) if x.dtype == "object" else x)

    # Salvar o DataFrame resultante em um novo arquivo Excel na pasta selecionada
    output_file_path = output_folder + "/texto_limpo.xlsx"
    df.to_excel(output_file_path, index=False)
    # Exibir uma mensagem de conclusão
    print(f"Texto limpo salvo em: {output_file_path}")
    messagebox.showinfo("Concluído", "Texto limpo salvo com sucesso!")

# FRONT END
root = tk.Tk()
root.title("Removedor de HTML")

# Botão para selecionar arquivo Excel
button_select_file = tk.Button(root, text="Selecionar Arquivo Excel", command=select_excel_file)
button_select_file.pack(pady=10)

# Entrada de texto para exibir o caminho do arquivo selecionado
entry_file_path = tk.Entry(root, width=50)
entry_file_path.pack(pady=5)

# Botão para selecionar a pasta de destino
button_select_folder = tk.Button(root, text="Selecionar Pasta de Destino", command=select_output_folder)
button_select_folder.pack(pady=10)

# Entrada de texto para exibir o caminho da pasta selecionada
entry_output_folder = tk.Entry(root, width=50)
entry_output_folder.pack(pady=5)

# Botão para remover HTML
button_remove_html = tk.Button(root, text="Remover HTML", command=process_html_removal)
button_remove_html.pack(pady=20)

# Executar o loop principal da janela
root.mainloop()
