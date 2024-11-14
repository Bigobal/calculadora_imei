import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Função para calcular o dígito verificador Luhn
def final_imei_luhn(imei_base):
    """Algoritmo de Luhn para calcular o dígito verificador do IMEI."""
    imei_digits = [int(d) for d in imei_base]
    odd_sum = 0
    even_sum = 0

    # Passo 1: Somar dígitos nas posições ímpares (1º, 3º, etc.)
    for i in range(0, len(imei_digits), 2):
        odd_sum += imei_digits[i]

    # Passo 2: Dobrar os dígitos nas posições pares e somar os dígitos individuais
    for i in range(1, len(imei_digits), 2):
        doubled = imei_digits[i] * 2
        even_sum += sum(int(digit) for digit in str(doubled))  # Somar os dígitos do resultado dobrado

    # Passo 3: Somar os resultados ímpares e pares
    total_sum = odd_sum + even_sum

    # Passo 4: O dígito verificador é o necessário para tornar a soma divisível por 10
    check_digit = (10 - (total_sum % 10)) % 10
    return check_digit

# Função para processar o arquivo de IMEIs selecionado
def process_imei_file(file_path):
    # Verificar se o arquivo de entrada existe
    if not os.path.exists(file_path):
        messagebox.showerror("Erro", f"Arquivo de entrada não encontrado: {file_path}")
        return

    # Carregar o arquivo de entrada
    try:
        imei_df = pd.read_excel(file_path)
        imei_df['IMEI'] = imei_df['IMEI'].astype(str)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {str(e)}")
        return
    
    # Aplicar o cálculo do dígito verificador para IMEIs
    imei_df['IMEI_Calculado'] = imei_df['IMEI'].apply(lambda imei: imei[:-1] + str(final_imei_luhn(imei[:-1])))

    # Caminho para salvar o arquivo de saída na mesma pasta
    output_file_path = os.path.splitext(file_path)[0] + '_calculado.xlsx'

    # Salvar o arquivo com os IMEIs calculados
    try:
        imei_df[['IMEI_Calculado']].to_excel(output_file_path, index=False)
        messagebox.showinfo("Sucesso", f"Arquivo ajustado salvo em: {output_file_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {str(e)}")

# Função para abrir o diálogo de seleção de arquivo
def select_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo IMEI",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if file_path:
        process_imei_file(file_path)

# Função para iniciar a interface gráfica
def create_gui():
    # Criar a janela principal
    root = tk.Tk()
    root.title("Calculadora de IMEI")

    # Definir tamanho mínimo da janela
    root.geometry("300x150")

    # Texto explicativo
    label = tk.Label(root, text="Selecione o arquivo IMEI para calcular", pady=20)
    label.pack()

    # Botão para selecionar o arquivo
    button = tk.Button(root, text="Selecionar arquivo", command=select_file, padx=20, pady=10)
    button.pack()

    # Executar a interface
    root.mainloop()

# Executar a interface gráfica
if __name__ == "__main__":
    create_gui()
