import pandas as pd
from tkinter import Tk, Label, Text, Button, END, messagebox
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows




# Função para adicionar feedback e sugestão de melhoria
def add_feedback():
    global df  # Declarar df como global
    feedback = feedback_entry.get("1.0", END).strip()
    suggestion = suggestion_entry.get("1.0", END).strip()
   
    # Limitar o número de caracteres do feedback
    max_feedback_length = 500
    if len(feedback) > max_feedback_length:
        messagebox.showwarning("Feedback", f"O feedback deve ter no máximo {max_feedback_length} caracteres.")
        return
   
    if feedback:
        new_row = {
            'date': datetime.now().strftime('%Y-%m-%d'),
            'feedback': feedback,
            'suggestion': suggestion
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        # Usar um nome de arquivo temporário para evitar problemas de permissão
        temp_filename = 'feedbacks_temp.xlsx'
        df.to_excel(temp_filename, index=False)
       
        # Aplicar formatação ao arquivo Excel
        format_excel(temp_filename, 'feedbacks.xlsx')
       
        feedback_entry.delete("1.0", END)
        suggestion_entry.delete("1.0", END)
        messagebox.showinfo("Feedback", "Feedback e sugestão adicionados com sucesso!")
    else:
        messagebox.showwarning("Feedback", "Por favor, insira um feedback.")




def format_excel(input_filename, output_filename):
    # Carregar o arquivo temporário
    wb = load_workbook(input_filename)
    ws = wb.active




    # Definir estilo do cabeçalho
    header_fill = PatternFill(start_color='0000FF', end_color='0000FF', fill_type='solid')
    header_font = Font(color='FFFFFF', bold=True)
    border = Border(left=Side(border_style='thin'),
                    right=Side(border_style='thin'),
                    top=Side(border_style='thin'),
                    bottom=Side(border_style='thin'))




    # Aplicar estilo ao cabeçalho
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border




    # Aplicar bordas e definir largura das colunas
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border




    # Ajustar largura das colunas
    column_widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                column_widths[cell.column] = max((column_widths.get(cell.column, 0), len(str(cell.value))))




    for col, width in column_widths.items():
        ws.column_dimensions[chr(64 + col)].width = width + 2  # +2 para algum padding




    # Salvar o arquivo formatado
    wb.save(output_filename)




# Inicializar DataFrame
try:
    df = pd.read_excel('feedbacks.xlsx')
except FileNotFoundError:
    df = pd.DataFrame(columns=['date', 'feedback', 'suggestion'])




# Interface gráfica com tkinter
root = Tk()
root.title("Coleta de Feedbacks")




Label(root, text="Digite seu feedback (máximo de 500 caracteres):").pack()
feedback_entry = Text(root, height=10, width=50)
feedback_entry.pack()




Label(root, text="Digite sua sugestão de melhoria (opcional):").pack()
suggestion_entry = Text(root, height=5, width=50)
suggestion_entry.pack()




Button(root, text="Adicionar Feedback e Sugestão", command=add_feedback).pack()




root.mainloop()
