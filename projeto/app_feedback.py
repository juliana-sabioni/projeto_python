import pandas as pd
from tkinter import Tk, Label, Text, Button, END, messagebox
from datetime import datetime


# Função para adicionar feedback
def add_feedback():
    global df  # Declarar df como global
    feedback = feedback_entry.get("1.0", END).strip()
    if feedback:
        new_row = {
            'Projeto': len(df) + 1,
            'Data': datetime.now().strftime('%Y-%m-%d'),
            'Feedback': feedback
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel('feedbacks.xlsx', index=False)  # Salvar imediatamente em Excel
        feedback_entry.delete("1.0", END)
        messagebox.showinfo("Feedback", "Feedback adicionado com sucesso!")
    else:
        messagebox.showwarning("Feedback", "Por favor, insira um feedback.")


# Inicializar DataFrame
try:
    df = pd.read_excel('feedbacks.xlsx')
except FileNotFoundError:
    df = pd.DataFrame(columns=['Projeto', 'Data', 'Feedback'])


# Interface gráfica com tkinter
root = Tk()
root.title("Coleta de Feedbacks")


Label(root, text="Digite seu feedback:").pack()
feedback_entry = Text(root, height=10, width=50)
feedback_entry.pack()


Button(root, text="Adicionar Feedback", command=add_feedback).pack()


root.mainloop()


