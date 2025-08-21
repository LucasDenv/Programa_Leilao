import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import pandas as pd
import os
import subprocess
from datetime import datetime
import shutil

ARQUIVO = 'lotes.xlsx'
ABA_LOTES = 'Lotes'
ABA_HISTORICO = 'Historico'
ABA_RESUMO = 'Resumo'

COLUNAS = ['Nome', 'Preço', 'Descrição', 'Lote', 'Data Cadastro']

# --- Funções auxiliares para planilha ---

def carregar_planilha():
    if os.path.exists(ARQUIVO):
        try:
            xls = pd.ExcelFile(ARQUIVO)
            if ABA_LOTES in xls.sheet_names:
                df_lotes = pd.read_excel(xls, sheet_name=ABA_LOTES)
            else:
                df_lotes = pd.DataFrame(columns=COLUNAS)
            if ABA_HISTORICO in xls.sheet_names:
                df_hist = pd.read_excel(xls, sheet_name=ABA_HISTORICO)
            else:
                df_hist = pd.DataFrame(columns=['Timestamp', 'Lote', 'Ação', 'Campo', 'Valor Antigo', 'Valor Novo'])
            if ABA_RESUMO in xls.sheet_names:
                df_resumo = pd.read_excel(xls, sheet_name=ABA_RESUMO)
            else:
                df_resumo = pd.DataFrame()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
            df_lotes = pd.DataFrame(columns=COLUNAS)
            df_hist = pd.DataFrame(columns=['Timestamp', 'Lote', 'Ação', 'Campo', 'Valor Antigo', 'Valor Novo'])
            df_resumo = pd.DataFrame()
    else:
        df_lotes = pd.DataFrame(columns=COLUNAS)
        df_hist = pd.DataFrame(columns=['Timestamp', 'Lote', 'Ação', 'Campo', 'Valor Antigo', 'Valor Novo'])
        df_resumo = pd.DataFrame()
    return df_lotes, df_hist, df_resumo

def salvar_planilha(df_lotes, df_hist, df_resumo):
    try:
        with pd.ExcelWriter(ARQUIVO, engine='openpyxl') as writer:
            df_lotes.to_excel(writer, sheet_name=ABA_LOTES, index=False)
            df_hist.to_excel(writer, sheet_name=ABA_HISTORICO, index=False)
            df_resumo.to_excel(writer, sheet_name=ABA_RESUMO, index=False)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar a planilha: {e}")

def criar_backup():
    if os.path.exists(ARQUIVO):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        nome_backup = f"lotes_backup_{timestamp}.xlsx"
        try:
            shutil.copy(ARQUIVO, nome_backup)
        except Exception as e:
            messagebox.showwarning("Backup", f"Não foi possível criar backup: {e}")

def gerar_novo_codigo(df_lotes):
    if df_lotes.empty:
        return "L001"
    ultimos_codigos = df_lotes['Lote'].dropna().astype(str)
    numeros = [int(codigo[1:]) for codigo in ultimos_codigos if codigo.startswith('L') and codigo[1:].isdigit()]
    proximo = max(numeros) + 1 if numeros else 1
    return f"L{proximo:03d}"

def atualizar_resumo(df_lotes):
    total_lotes = len(df_lotes)
    soma_precos = df_lotes['Preço'].sum() if not df_lotes.empty else 0
    ultimo_lote = df_lotes['Lote'].iloc[-1] if not df_lotes.empty else ''
    ultima_data = df_lotes['Data Cadastro'].max() if not df_lotes.empty else ''
    df_resumo = pd.DataFrame({
        'Total de Lotes': [total_lotes],
        'Soma dos Preços': [soma_precos],
        'Último Lote Cadastrado': [ultimo_lote],
        'Última Data de Cadastro': [ultima_data]
    })
    return df_resumo

# --- Funções da interface ---

class InventarioApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Inventário de Lotes")
        self.master.geometry("600x650")
        self.master.resizable(False, False)

        # Dados carregados
        self.df_lotes, self.df_hist, self.df_resumo = carregar_planilha()

        # Variáveis Tkinter
        self.var_nome = tk.StringVar()
        self.var_preco = tk.StringVar()
        self.var_lote = tk.StringVar()
        self.text_descricao = None

        # Busca
        self.var_busca = tk.StringVar()

        self._criar_widgets()
        self._atualizar_codigo_lote()

        # Evento fechamento
        self.master.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _criar_widgets(self):
        # Campos entrada
        tk.Label(self.master, text="Nome do Produto *").pack(anchor='w', padx=10, pady=(10,0))
        self.entry_nome = tk.Entry(self.master, textvariable=self.var_nome, width=70)
        self.entry_nome.pack(padx=10)

        tk.Label(self.master, text="Preço *").pack(anchor='w', padx=10, pady=(10,0))
        self.entry_preco = tk.Entry(self.master, textvariable=self.var_preco, width=70)
        self.entry_preco.pack(padx=10)

        tk.Label(self.master, text="Descrição").pack(anchor='w', padx=10, pady=(10,0))
        self.text_descricao = tk.Text(self.master, height=5, width=70)
        self.text_descricao.pack(padx=10)

        tk.Label(self.master, text="Código do Lote (gerado automaticamente)").pack(anchor='w', padx=10, pady=(10,0))
        self.entry_lote = tk.Entry(self.master, textvariable=self.var_lote, width=70, state='readonly')
        self.entry_lote.pack(padx=10)

        # Botões principais
        tk.Button(self.master, text="➕ Adicionar Lote", bg="#4CAF50", fg="white", height=2, command=self.adicionar_lote).pack(pady=15)

        # Busca
        frame_busca = tk.Frame(self.master)
        frame_busca.pack(padx=10, pady=(5,10), fill='x')
        tk.Label(frame_busca, text="Buscar por Nome ou Lote:").pack(side='left')
        self.entry_busca = tk.Entry(frame_busca, textvariable=self.var_busca)
        self.entry_busca.pack(side='left', expand=True, fill='x', padx=(5,5))
        tk.Button(frame_busca, text="🔍 Buscar", command=self.buscar_lotes).pack(side='left')

        # Botões auxiliares
        frame_botoes = tk.Frame(self.master)
        frame_botoes.pack(pady=5)

        tk.Button(frame_botoes, text="✏️ Editar Lote", command=self.editar_lote).pack(side='left', padx=5)
        tk.Button(frame_botoes, text="🗑️ Excluir Lote", bg="#e53935", fg="white", command=self.excluir_lote).pack(side='left', padx=5)
        tk.Button(frame_botoes, text="📄 Duplicar Lote", command=self.duplicar_lote).pack(side='left', padx=5)
        tk.Button(frame_botoes, text="📂 Abrir Planilha", command=self.abrir_planilha).pack(side='left', padx=5)
        tk.Button(frame_botoes, text="📊 Resumo", command=self.exibir_resumo).pack(side='left', padx=5)

    def _atualizar_codigo_lote(self):
        novo_codigo = gerar_novo_codigo(self.df_lotes)
        self.var_lote.set(novo_codigo)

    def _validar_campos(self):
        erro = False
        # Reset cores
        self.entry_nome.config(bg="white")
        self.entry_preco.config(bg="white")

        nome = self.var_nome.get().strip()
        preco = self.var_preco.get().replace(",", ".").strip()

        if nome == "":
            self.entry_nome.config(bg="#ffcccc")
            erro = True
        try:
            preco_float = float(preco)
            if preco_float < 0:
                self.entry_preco.config(bg="#ffcccc")
                erro = True
        except:
            self.entry_preco.config(bg="#ffcccc")
            erro = True

        return not erro

    def adicionar_lote(self):
        if not self._validar_campos():
            messagebox.showerror("Erro", "Preencha os campos obrigatórios corretamente (Nome e Preço positivo).")
            return

        nome = self.var_nome.get().strip()
        preco = float(self.var_preco.get().replace(",", "."))
        descricao = self.text_descricao.get("1.0", tk.END).strip()
        lote = self.var_lote.get()

        # Verifica duplicidade lote e nome
        if lote in self.df_lotes['Lote'].astype(str).values:
            messagebox.showerror("Erro", f"Lote {lote} já cadastrado.")
            return
        if nome in self.df_lotes['Nome'].astype(str).values:
            if not messagebox.askyesno("Duplicidade", f"Já existe um produto com o nome '{nome}'. Deseja continuar?"):
                return

        data_cadastro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        novo = pd.DataFrame([[nome, preco, descricao, lote, data_cadastro]], columns=COLUNAS)
        self.df_lotes = pd.concat([self.df_lotes, novo], ignore_index=True)

        # Histórico
        self._adicionar_historico(lote, 'Inserção', '', '')

        salvar_planilha(self.df_lotes, self.df_hist, self.df_resumo)
        criar_backup()

        # Limpa e atualiza
        self.var_nome.set('')
        self.var_preco.set('')
        self.text_descricao.delete("1.0", tk.END)
        self._atualizar_codigo_lote()
        messagebox.showinfo("Sucesso", f"Lote {lote} cadastrado com sucesso!")

    def _adicionar_historico(self, lote, acao, campo, valor_antigo, valor_novo=''):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        novo_log = pd.DataFrame([[timestamp, lote, acao, campo, valor_antigo, valor_novo]],
                               columns=['Timestamp', 'Lote', 'Ação', 'Campo', 'Valor Antigo', 'Valor Novo'])
        self.df_hist = pd.concat([self.df_hist, novo_log], ignore_index=True)

    def buscar_lotes(self):
        termo = self.var_busca.get().strip().lower()
        if termo == '':
            dados = self.df_lotes.copy()
        else:
            mask = self.df_lotes['Nome'].str.lower().str.contains(termo) | self.df_lotes['Lote'].str.lower().str.contains(termo)
            dados = self.df_lotes[mask]

        self._mostrar_resultados(dados, f"Resultados da busca por '{termo}'")

    def _mostrar_resultados(self, dados, titulo):
        if dados.empty:
            messagebox.showinfo("Busca", "Nenhum lote encontrado.")
            return
        win = tk.Toplevel(self.master)
        win.title(titulo)
        win.geometry("700x400")

        tree = ttk.Treeview(win, columns=COLUNAS, show='headings')
        for c in COLUNAS:
            tree.heading(c, text=c)
            tree.column(c, width=130, anchor='center')
        for _, row in dados.iterrows():
            tree.insert('', 'end', values=list(row))
        tree.pack(fill='both', expand=True)

        tk.Button(win, text="Fechar", command=win.destroy).pack(pady=5)

    def editar_lote(self):
        codigo = simpledialog.askstring("Editar Lote", "Digite o código do lote (ex: L005):")
        if not codigo:
            return
        dados = self.df_lotes[self.df_lotes['Lote'] == codigo]
        if dados.empty:
            messagebox.showerror("Erro", f"Lote {codigo} não encontrado.")
            return
        linha = dados.iloc[0]

        # Janela de edição
        edit_win = tk.Toplevel(self.master)
        edit_win.title(f"Editar Lote {codigo}")
        edit_win.geometry("500x400")
        edit_win.resizable(False, False)

        # Variáveis
        var_nome = tk.StringVar(value=linha['Nome'])
        var_preco = tk.StringVar(value=str(linha['Preço']))
        var_lote = tk.StringVar(value=linha['Lote'])
        text_desc = tk.Text(edit_win, height=5, width=60)
        text_desc.insert('1.0', linha['Descrição'])

        tk.Label(edit_win, text="Nome *").pack(anchor='w', padx=10, pady=(10,0))
        entry_nome = tk.Entry(edit_win, textvariable=var_nome, width=60)
        entry_nome.pack(padx=10)

        tk.Label(edit_win, text="Preço *").pack(anchor='w', padx=10, pady=(10,0))
        entry_preco = tk.Entry(edit_win, textvariable=var_preco, width=60)
        entry_preco.pack(padx=10)

        tk.Label(edit_win, text="Descrição").pack(anchor='w', padx=10, pady=(10,0))
        text_desc.pack(padx=10)

        tk.Label(edit_win, text="Código do Lote").pack(anchor='w', padx=10, pady=(10,0))
        entry_lote = tk.Entry(edit_win, textvariable=var_lote, width=60, state='readonly')
        entry_lote.pack(padx=10)

        def salvar_edicao():
            novo_nome = var_nome.get().strip()
            novo_preco = var_preco.get().replace(',', '.').strip()
            nova_desc = text_desc.get("1.0", tk.END).strip()

            if novo_nome == '':
                messagebox.showerror("Erro", "Nome não pode ser vazio.")
                return
            try:
                preco_float = float(novo_preco)
                if preco_float < 0:
                    raise ValueError
            except:
                messagebox.showerror("Erro", "Preço inválido.")
                return

            # Atualizar dataframe e histórico
            idx = self.df_lotes.index[self.df_lotes['Lote'] == codigo][0]

            # Campos para comparar e logar
            campos = ['Nome', 'Preço', 'Descrição']
            novos_valores = [novo_nome, preco_float, nova_desc]
            antigos_valores = self.df_lotes.loc[idx, campos]

            for campo, antigo, novo in zip(campos, antigos_valores, novos_valores):
                if antigo != novo:
                    self._adicionar_historico(codigo, 'Edição', campo, antigo, novo)
                    self.df_lotes.at[idx, campo] = novo

            salvar_planilha(self.df_lotes, self.df_hist, self.df_resumo)
            criar_backup()
            messagebox.showinfo("Sucesso", f"Lote {codigo} atualizado.")
            edit_win.destroy()

        tk.Button(edit_win, text="Salvar", bg="#4CAF50", fg="white", command=salvar_edicao).pack(pady=15)

    def excluir_lote(self):
        codigo = simpledialog.askstring("Excluir Lote", "Digite o código do lote para excluir (ex: L003):")
        if not codigo:
            return
        if codigo not in self.df_lotes['Lote'].astype(str).values:
            messagebox.showerror("Erro", f"Lote {codigo} não encontrado.")
            return
        confirm = messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o lote {codigo}?")
        if confirm:
            idx = self.df_lotes.index[self.df_lotes['Lote'] == codigo][0]
            # Log antigo
            antigo = self.df_lotes.loc[idx].to_dict()
            self._adicionar_historico(codigo, 'Exclusão', 'Todos', str(antigo), '')
            self.df_lotes = self.df_lotes.drop(idx).reset_index(drop=True)
            salvar_planilha(self.df_lotes, self.df_hist, self.df_resumo)
            criar_backup()
            messagebox.showinfo("Sucesso", f"Lote {codigo} excluído.")

            self._atualizar_codigo_lote()

    def duplicar_lote(self):
        codigo = simpledialog.askstring("Duplicar Lote", "Digite o código do lote para duplicar (ex: L003):")
        if not codigo:
            return
        dados = self.df_lotes[self.df_lotes['Lote'] == codigo]
        if dados.empty:
            messagebox.showerror("Erro", f"Lote {codigo} não encontrado.")
            return
        linha = dados.iloc[0]

        # Abrir janela para novo cadastro com dados pré-preenchidos
        dup_win = tk.Toplevel(self.master)
        dup_win.title(f"Duplicar Lote {codigo}")
        dup_win.geometry("600x450")
        dup_win.resizable(False, False)

        var_nome = tk.StringVar(value=linha['Nome'])
        var_preco = tk.StringVar(value=str(linha['Preço']))
        var_lote = tk.StringVar()
        text_desc = tk.Text(dup_win, height=5, width=70)
        text_desc.insert('1.0', linha['Descrição'])

        tk.Label(dup_win, text="Nome do Produto *").pack(anchor='w', padx=10, pady=(10,0))
        entry_nome = tk.Entry(dup_win, textvariable=var_nome, width=70)
        entry_nome.pack(padx=10)

        tk.Label(dup_win, text="Preço *").pack(anchor='w', padx=10, pady=(10,0))
        entry_preco = tk.Entry(dup_win, textvariable=var_preco, width=70)
        entry_preco.pack(padx=10)

        tk.Label(dup_win, text="Descrição").pack(anchor='w', padx=10, pady=(10,0))
        text_desc.pack(padx=10)

        tk.Label(dup_win, text="Código do Lote (gerado automaticamente)").pack(anchor='w', padx=10, pady=(10,0))
        entry_lote = tk.Entry(dup_win, textvariable=var_lote, width=70, state='readonly')
        entry_lote.pack(padx=10)

        def atualizar_codigo_dup():
            novo_codigo = gerar_novo_codigo(self.df_lotes)
            var_lote.set(novo_codigo)
        atualizar_codigo_dup()

        def salvar_duplicacao():
            novo_nome = var_nome.get().strip()
            novo_preco = var_preco.get().replace(",", ".").strip()
            nova_desc = text_desc.get("1.0", tk.END).strip()
            novo_lote = var_lote.get()

            if novo_nome == "":
                messagebox.showerror("Erro", "Nome não pode ser vazio.")
                return
            try:
                preco_float = float(novo_preco)
                if preco_float < 0:
                    raise ValueError
            except:
                messagebox.showerror("Erro", "Preço inválido.")
                return
            if novo_lote in self.df_lotes['Lote'].astype(str).values:
                messagebox.showerror("Erro", f"Lote {novo_lote} já existe.")
                return

            data_cadastro = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            novo = pd.DataFrame([[novo_nome, preco_float, nova_desc, novo_lote, data_cadastro]], columns=COLUNAS)
            self.df_lotes = pd.concat([self.df_lotes, novo], ignore_index=True)
            self._adicionar_historico(novo_lote, 'Inserção (Duplicado)', '', '')

            salvar_planilha(self.df_lotes, self.df_hist, self.df_resumo)
            criar_backup()
            messagebox.showinfo("Sucesso", f"Lote {novo_lote} duplicado cadastrado.")
            dup_win.destroy()
            self._atualizar_codigo_lote()

        tk.Button(dup_win, text="Salvar", bg="#4CAF50", fg="white", command=salvar_duplicacao).pack(pady=15)

    def abrir_planilha(self):
        if os.path.exists(ARQUIVO):
            try:
                # Windows
                if os.name == 'nt':
                    os.startfile(ARQUIVO)
                # macOS
                elif sys.platform == "darwin":
                    subprocess.Popen(['open', ARQUIVO])
                # Linux
                else:
                    subprocess.Popen(['xdg-open', ARQUIVO])
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir a planilha:\n{e}")
        else:
            messagebox.showwarning("Aviso", "Planilha ainda não existe.")

    def exibir_resumo(self):
        self.df_resumo = atualizar_resumo(self.df_lotes)
        texto = (f"Total de Lotes: {self.df_resumo['Total de Lotes'][0]}\n"
                 f"Soma dos Preços: R$ {self.df_resumo['Soma dos Preços'][0]:.2f}\n"
                 f"Último Lote Cadastrado: {self.df_resumo['Último Lote Cadastrado'][0]}\n"
                 f"Última Data de Cadastro: {self.df_resumo['Última Data de Cadastro'][0]}")
        messagebox.showinfo("Resumo do Inventário", texto)

    def _on_closing(self):
        salvar_planilha(self.df_lotes, self.df_hist, self.df_resumo)
        criar_backup()
        self.master.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    app = InventarioApp(root)
    root.mainloop()
