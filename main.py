# ============================================================
# SISTEMA INTEGRADO DE CONTAS A PAGAR E FORNECEDORES
# ============================================================

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime  # <-- Adicionado para corrigir o erro de 'datetime'
from tkcalendar import DateEntry  # <-- Componente de Calendário

# Adicionado para corrigir o erro NameError: 'EXCEL_DISPONIVEL'
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    EXCEL_DISPONIVEL = True
except ImportError:
    EXCEL_DISPONIVEL = False


# ============================================================
# BANCO DE DADOS UNIFICADO
# ============================================================

def conectar():
    """Conecta ao banco de dados unificado do sistema."""
    return sqlite3.connect("sistema_financeiro.db")


def criar_tabelas():
    """Cria as tabelas necessárias se elas não existirem."""
    conn = conectar()
    cursor = conn.cursor()

    # Tabela de fornecedores atualizada com os novos campos do print
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS fornecedores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            cnpj TEXT NOT NULL,
            telefone TEXT,
            email TEXT
        )
    """)

    # Tabela de contas a pagar relacionada com fornecedores
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS contas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descricao TEXT NOT NULL,
            valor REAL NOT NULL,
            vencimento TEXT NOT NULL,
            status TEXT DEFAULT 'Pendente',
            fornecedor_id INTEGER,
            FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id)
        )
    """)

    conn.commit()
    conn.close()


# ============================================================
# JANELA: GERENCIAR FORNECEDORES
# ============================================================

class JanelaFornecedores(tk.Toplevel):
    def __init__(self, parent, callback_atualizar=None):
        super().__init__(parent)
        self.title("Gerenciar Fornecedores")
        self.geometry("550x500")
        self.resizable(False, False)
        self.grab_set()
        self.callback_atualizar = callback_atualizar
        self.criar_widgets()
        self.carregar_fornecedores()

    def criar_widgets(self):
        # Formulário de Cadastro
        frame_cad = tk.LabelFrame(self, text="Novo Fornecedor", padx=10, pady=10)
        frame_cad.pack(fill="x", padx=10, pady=10)

        tk.Label(frame_cad, text="Nome:").grid(row=0, column=0, sticky="w", pady=2)
        self.entry_nome = tk.Entry(frame_cad, width=45)
        self.entry_nome.grid(row=0, column=1, padx=5, pady=2)

        tk.Label(frame_cad, text="CNPJ:").grid(row=1, column=0, sticky="w", pady=2)
        self.entry_cnpj = tk.Entry(frame_cad, width=45)
        self.entry_cnpj.grid(row=1, column=1, padx=5, pady=2)

        tk.Label(frame_cad, text="Telefone:").grid(row=2, column=0, sticky="w", pady=2)
        self.entry_telefone = tk.Entry(frame_cad, width=45)
        self.entry_telefone.grid(row=2, column=1, padx=5, pady=2)

        tk.Label(frame_cad, text="Email:").grid(row=3, column=0, sticky="w", pady=2)
        self.entry_email = tk.Entry(frame_cad, width=45)
        self.entry_email.grid(row=3, column=1, padx=5, pady=2)

        tk.Button(frame_cad, text="Adicionar Fornecedor", command=self.adicionar, bg="#4CAF50", fg="white",
                  width=20).grid(row=4, column=0, columnspan=2, pady=10)

        # Visualização da Tabela
        frame_lista = tk.LabelFrame(self, text="Fornecedores Cadastrados", padx=10, pady=10)
        frame_lista.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(frame_lista, columns=("Nome", "CNPJ"), show="headings", height=6)
        self.tree.heading("Nome", text="Nome")
        self.tree.heading("CNPJ", text="CNPJ")
        self.tree.column("Nome", width=300)
        self.tree.column("CNPJ", width=180)
        self.tree.pack(fill="both", expand=True)

        tk.Button(self, text="Excluir Selecionado", command=self.excluir, bg="#f44336", fg="white").pack(pady=5)

    def carregar_fornecedores(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, cnpj FROM fornecedores ORDER BY nome")
        self.fornecedores = cursor.fetchall()
        conn.close()

        for f in self.fornecedores:
            self.tree.insert("", tk.END, iid=f[0], values=(f[1], f[2]))

    def adicionar(self):
        nome = self.entry_nome.get().strip()
        cnpj = self.entry_cnpj.get().strip()
        telefone = self.entry_telefone.get().strip()
        email = self.entry_email.get().strip()

        if not nome or not cnpj:
            messagebox.showwarning("Aviso", "Nome e CNPJ são obrigatórios!", parent=self)
            return

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO fornecedores (nome, cnpj, telefone, email) VALUES (?, ?, ?, ?)",
            (nome, cnpj, telefone, email)
        )
        conn.commit()
        conn.close()

        self.entry_nome.delete(0, tk.END)
        self.entry_cnpj.delete(0, tk.END)
        self.entry_telefone.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)

        self.carregar_fornecedores()
        if self.callback_atualizar:
            self.callback_atualizar()
        messagebox.showinfo("Sucesso", "Fornecedor cadastrado!", parent=self)

    def excluir(self):
        selecionado = self.tree.selection()
        if not id(selecionado):
            if not selecionado:
                messagebox.showwarning("Atenção", "Selecione um fornecedor para excluir.", parent=self)
                return

        fornecedor_id = selecionado[0]
        if messagebox.askyesno("Confirmar", "Deseja realmente excluir este fornecedor?", parent=self):
            conn = conectar()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM fornecedores WHERE id = ?", (fornecedor_id,))
            conn.commit()
            conn.close()
            self.carregar_fornecedores()
            if self.callback_atualizar:
                self.callback_atualizar()


# ============================================================
# JANELA: CADASTRAR / EDITAR CONTA
# ============================================================

class JanelaConta(tk.Toplevel):
    def __init__(self, parent, callback_atualizar, conta_id=None):
        super().__init__(parent)
        self.callback_atualizar = callback_atualizar
        self.conta_id = conta_id
        self.title("Editar Conta" if conta_id else "Nova Conta")
        self.geometry("400x320")
        self.resizable(False, False)
        self.grab_set()
        self.criar_widgets()
        self.carregar_fornecedores()
        if conta_id:
            self.preencher_dados()

    def criar_widgets(self):
        frame = tk.Frame(self, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Valor (R$):").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_valor = tk.Entry(frame, width=15)
        self.entry_valor.grid(row=1, column=1, sticky="w", pady=5)

        tk.Label(frame, text="Vencimento (DD/MM/AAAA):").grid(row=2, column=0, sticky="w", pady=5)
        self.entry_vencimento = tk.Entry(frame, width=15)
        self.entry_vencimento.grid(row=2, column=1, sticky="w", pady=5)

        tk.Label(frame, text="Fornecedor:").grid(row=3, column=0, sticky="w", pady=5)

        # Mudamos state de "readonly" para "normal" para permitir que você digite o nome!
        self.combo_fornecedor = ttk.Combobox(frame, width=27, state="normal")
        self.combo_fornecedor.grid(row=3, column=1, sticky="w", pady=5)

        # Evento que roda a função de filtro sempre que você solta uma tecla digitada
        self.combo_fornecedor.bind("<KeyRelease>", self.filtrar_fornecedores)

        tk.Label(frame, text="Status:").grid(row=4, column=0, sticky="w", pady=5)
        self.combo_status = ttk.Combobox(frame, values=["Pendente", "Pago"], width=12, state="readonly")
        self.combo_status.set("Pendente")
        self.combo_status.grid(row=4, column=1, sticky="w", pady=5)

        tk.Button(frame, text="💾 Salvar Conta", command=self.salvar, bg="#2196F3", fg="white", width=15).grid(row=5,
                                                                                                              column=0,
                                                                                                              columnspan=2,
                                                                                                              pady=15)

    def carregar_fornecedores(self):
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome FROM fornecedores ORDER BY nome")
        self.lista_fornecedores = cursor.fetchall()
        conn.close()

        # Guardamos apenas a lista de nomes limpos para o filtro visual funcionar
        self.nomes_fornecedores = ["(Nenhum)"] + [f[1] for f in self.lista_fornecedores]
        self.combo_fornecedor["values"] = self.nomes_fornecedores
        self.combo_fornecedor.set("(Nenhum)")

    def filtrar_fornecedores(self, event):
        """Filtra as opções do Combobox conforme o usuário digita."""
        texto_digitado = self.combo_fornecedor.get().strip().lower()

        if texto_digitado == "":
            # Se apagar tudo, volta a exibir a lista completa
            self.combo_fornecedor["values"] = self.nomes_fornecedores
        else:
            # Cria uma lista apenas com os nomes que contêm o texto digitado
            lista_filtrada = [nome for nome in self.nomes_fornecedores if texto_digitado in nome.lower()]
            self.combo_fornecedor["values"] = lista_filtrada

        # Abre o menu suspenso de escolhas automaticamente enquanto digita
        self.combo_fornecedor.event_generate("<Down>")

    def preencher_dados(self):
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT valor, vencimento, status, fornecedor_id FROM contas WHERE id = ?",
                       (self.conta_id,))
        conta = cursor.fetchone()
        conn.close()

        if conta:
            val, venc, status, f_id = conta
            self.entry_valor.insert(0, str(val))
            self.entry_vencimento.insert(0, venc)
            self.combo_status.set(status)

            if f_id:
                for idx, f in enumerate(self.lista_fornecedores):
                    if f[0] == f_id:
                        self.combo_fornecedor.set(f[1])
                        break

    def salvar(self):
        val_str = self.entry_valor.get().strip()
        venc = self.entry_vencimento.get().strip()
        status = self.combo_status.get()
        nome_selecionado = self.combo_fornecedor.get().strip()

        if not val_str or not venc:
            messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios!", parent=self)
            return

        try:
            val = float(val_str.replace(",", "."))
        except ValueError:
            messagebox.showerror("Erro", "Valor numérico inválido.", parent=self)
            return

        # Busca o ID correto cruzando o nome escrito/selecionado com nossa lista interna
        f_id = None
        if nome_selecionado and nome_selecionado != "(Nenhum)":
            for f in self.lista_fornecedores:
                if f[1].lower() == nome_selecionado.lower():
                    f_id = f[0]
                    break

            # Validação caso digitem um nome que não existe no banco
            if f_id is None:
                messagebox.showwarning("Aviso",
                                       "Fornecedor não encontrado! Cadastre-o primeiro ou selecione um válido.",
                                       parent=self)
                return

        conn = conectar()
        cursor = conn.cursor()
        if self.conta_id:
            cursor.execute("""
                UPDATE contas SET descricao=?, valor=?, vencimento=?, status=?, fornecedor_id=? WHERE id=?
            """, ("", val, venc, status, f_id, self.conta_id))  # Envia "" para a descrição
        else:
            cursor.execute("""
                INSERT INTO contas (descricao, valor, vencimento, status, fornecedor_id) VALUES (?, ?, ?, ?, ?)
            """, ("", val, venc, status, f_id))  # Envia "" para a descrição
        conn.commit()
        conn.close()

        self.callback_atualizar()
        messagebox.showinfo("Sucesso", "Dados salvos com sucesso!", parent=self)
        self.destroy()


# ============================================================
# JANELA PRINCIPAL (ROOT)
# ============================================================

class JanelaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle Financeiro")
        self.root.geometry("700x450")
        self.criar_widgets()
        self.carregar_contas()

    def criar_widgets(self):
        # --- Bloco de Botões Superiores ---
        frame_botoes = tk.Frame(self.root, pady=10)
        frame_botoes.pack(fill="x", padx=10)

        tk.Button(frame_botoes, text="➕ Nova Conta", command=self.nova_conta, bg="#2196F3", fg="white", width=12).pack(
            side="left", padx=3)
        tk.Button(frame_botoes, text="✏️ Editar Conta", command=self.editar_conta, bg="#FF9800", fg="white",
                  width=12).pack(side="left", padx=3)
        tk.Button(frame_botoes, text="❌ Excluir Conta", command=self.excluir_conta, bg="#f44336", fg="white",
                  width=12).pack(side="left", padx=3)
        tk.Button(frame_botoes, text="👥 Fornecedores", command=self.abrir_fornecedores, bg="#9C27B0", fg="white",
                  width=12).pack(side="right", padx=3)
        tk.Button(
            frame_botoes,
            text="✔ Dar Baixa",
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            command=self.dar_baixa
        ).pack(side="left", padx=5)

        # --- Bloco de Filtros por Período com Calendário Dinâmico ---
        frame_filtros = tk.LabelFrame(self.root, text="Filtro de Relatório por Vencimento", pady=5, padx=10)
        frame_filtros.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_filtros, text="De:").pack(side="left", padx=2)
        # Componente DateEntry: cria um campo de data formatado em PT-BR com calendário integrado
        self.entry_data_ini = DateEntry(frame_filtros, width=12, background='darkblue',
                                        foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_data_ini.pack(side="left", padx=5)

        tk.Label(frame_filtros, text="Até:").pack(side="left", padx=2)
        self.entry_data_fim = DateEntry(frame_filtros, width=12, background='darkblue',
                                        foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        self.entry_data_fim.pack(side="left", padx=5)

        tk.Button(frame_filtros, text="🔍 Filtrar", command=self.carregar_contas, bg="#009688", fg="white",
                  width=10).pack(side="left", padx=10)
        tk.Button(frame_filtros, text="📊 Exportar Excel", command=self.exportar_excel, bg="#4CAF50", fg="white",
                  width=15).pack(side="right", padx=5)

        # --- Grid da Tabela ---
        frame_grid = tk.Frame(self.root)
        frame_grid.pack(fill="both", expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(frame_grid, columns=("ID", "Fornecedor", "Valor", "Vencimento", "Status"), show="headings", selectmode="extended")
        self.tree.heading("ID", text="ID")
        self.tree.heading("Fornecedor", text="Fornecedor")
        self.tree.heading("Valor", text="Valor")
        self.tree.heading("Vencimento", text="Vencimento")
        self.tree.heading("Status", text="Status")

        self.tree["displaycolumns"] = (
    "Fornecedor",
    "Valor",
    "Vencimento",
    "Status"
)

        self.tree.column("ID", width=0, stretch=False)
        self.tree.column("Fornecedor", width=250, anchor="center")
        self.tree.column("Valor", width=120, anchor="center")
        self.tree.column("Vencimento", width=120, anchor="center")
        self.tree.column("Status", width=100, anchor="center")
        self.tree.pack(fill="both", expand=True)
        self.tree.tag_configure('vencido', background='#FFCCCC', foreground='black')

    def dar_baixa(self):

        itens_selecionados = self.tree.selection()

        if not itens_selecionados:
            messagebox.showwarning(
                "Aviso",
                "Selecione ao menos uma conta."
            )
            return

        resposta = messagebox.askyesno(
            "Confirmar",
            f"Deseja dar baixa em {len(itens_selecionados)} boleto(s)?"
        )

        if not resposta:
            return

        try:

            conn = conectar()
            cursor = conn.cursor()

            for item_id in itens_selecionados:
                item = self.tree.item(item_id)

                valores = item['values']
                print(valores)

                id_conta = valores[0]

                status = valores[3]

                if status == "Pago":
                    continue

                cursor.execute("""
                    UPDATE contas
                    SET status = 'Pago'
                    WHERE id = ?
                """, (id_conta,))

            conn.commit()
            conn.close()

            messagebox.showinfo(
                "Sucesso",
                "Baixa realizada com sucesso!"
            )

            self.carregar_contas()

        except Exception as e:
            messagebox.showerror(
                "Erro",
                f"Erro ao atualizar status:\n{e}"
            )

    def carregar_contas(self):
        """Carrega contas: filtra por data ou traz todas as vencidas por padrão."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Resgata os textos dos seletores DateEntry
        data_ini = self.entry_data_ini.get().strip()
        data_fim = self.entry_data_fim.get().strip()

        # Captura a data de hoje para o Python e formata para o banco (AAAA-MM-DD)
        data_hoje_obj = datetime.now().date()
        data_hoje_iso = data_hoje_obj.strftime("%Y-%m-%d")

        query = """
            SELECT c.id, f.nome, c.valor, c.vencimento, c.status 
            FROM contas c LEFT JOIN fornecedores f ON c.fornecedor_id = f.id
        """
        params = []

        if data_ini and data_fim:
            try:
                # Converte DD/MM/AAAA da interface para AAAA-MM-DD
                d_ini_iso = datetime.strptime(data_ini, "%d/%m/%Y").strftime("%Y-%m-%d")
                d_fim_iso = datetime.strptime(data_fim, "%d/%m/%Y").strftime("%Y-%m-%d")

                # Filtro por período selecionado pelo usuário
                query += " WHERE substr(c.vencimento,7,4)||'-'||substr(c.vencimento,4,2)||'-'||substr(c.vencimento,1,2) BETWEEN ? AND ?"
                params.extend([d_ini_iso, d_fim_iso])
            except ValueError:
                messagebox.showerror("Erro", "Formato interno de data inconsistente!")
                return
        else:
            # SE NÃO HOUVER FILTRO DE DATA: Mostra todos os vencidos e não pagos automaticamente
            query += " WHERE substr(c.vencimento,7,4)||'-'||substr(c.vencimento,4,2)||'-'||substr(c.vencimento,1,2) < ? AND LOWER(TRIM(c.status)) != 'pago'"
            params.append(data_hoje_iso)

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute(query, params)
        self.dados_carregados = cursor.fetchall()

        for linha in self.dados_carregados:
            c_id, forn_nome, valor, vencimento, status = linha

            nome_fornecedor = forn_nome if forn_nome else "(Nenhum)"

            try:
                valor_formatado = f"R$ {float(valor):.2f}"
            except (ValueError, TypeError):
                valor_formatado = f"R$ {valor}"

            # --- Lógica de Validação de Vencimento para as Tags ---
            tags_linha = ()
            try:
                data_venc_obj = datetime.strptime(vencimento, "%d/%m/%Y").date()
                if data_venc_obj < data_hoje_obj and str(status).strip().lower() != "pago":
                    tags_linha = ('vencido',)
            except (ValueError, TypeError):
                pass

            self.tree.insert(
                "",
                tk.END,
                iid=c_id,
                values=(
                    c_id,
                    nome_fornecedor,
                    valor_formatado,
                    vencimento,
                    status
                ),
                tags=tags_linha
            )
        conn.close()


    def exportar_excel(self):
        """Exporta os dados diretamente do Banco de Dados para garantir o alinhamento perfeito."""
        if not EXCEL_DISPONIVEL:
            messagebox.showerror("Erro", "A biblioteca 'openpyxl' não está instalada.")
            return

        if not self.tree.get_children():
            messagebox.showwarning("Aviso", "Não há dados na tabela para exportar.")
            return

        arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            title="Salvar Relatório Financeiro"
        )

        if not arquivo:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Relatório Financeiro"

        # 1. Configuração de Página para Impressão A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5

        # 2. Escreve o Cabeçalho
        headers = ["Fornecedor", "Valor (R$)", "Vencimento", "Status"]
        ws.append(headers)

        # 3. Busca os dados direto do Banco
        data_ini = self.entry_data_ini.get().strip()
        data_fim = self.entry_data_fim.get().strip()

        query = """
            SELECT f.nome, c.valor, c.vencimento, c.status 
            FROM contas c LEFT JOIN fornecedores f ON c.fornecedor_id = f.id
        """
        params = []

        if data_ini and data_fim:
            try:
                d_ini_iso = datetime.strptime(data_ini, "%d/%m/%Y").strftime("%Y-%m-%d")
                d_fim_iso = datetime.strptime(data_fim, "%d/%m/%Y").strftime("%Y-%m-%d")
                query += " WHERE substr(c.vencimento,7,4)||'-'||substr(c.vencimento,4,2)||'-'||substr(c.vencimento,1,2) BETWEEN ? AND ?"
                params.extend([d_ini_iso, d_fim_iso])
            except ValueError:
                pass

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute(query, params)
        dados_banco = cursor.fetchall()
        conn.close()

        # 4. Preenche as linhas do Excel com os dados limpos do banco
        valor_total = 0.0
        for linha in dados_banco:
            forn_nome, valor, vencimento, status = linha
            nome_fornecedor = forn_nome if forn_nome else "(Nenhum)"

            try:
                v_num = float(valor)
            except:
                v_num = 0.0

            valor_total += v_num

            try:
                data_excel = datetime.strptime(vencimento, "%d/%m/%Y")
            except:
                data_excel = vencimento

            ws.append([nome_fornecedor, v_num, data_excel, status])

        # Linha em branco
        ws.append([])

        # Linha total
        linha_total_index = ws.max_row + 1

        ws.cell(row=linha_total_index, column=1, value="TOTAL GERAL:")
        ws.cell(row=linha_total_index, column=2, value=valor_total)

        # 6. Estilização Visual Unificada e Centralização Total das Contas
        font_cabecalho = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_dados = Font(name="Arial", size=10, bold=False)
        font_total = Font(name="Arial", size=10, bold=True)

        fill_cabecalho = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

        align_centro = Alignment(horizontal="center", vertical="center")
        align_esquerda = Alignment(horizontal="left", vertical="center")

        borda_fina = Border(
            left=Side(style='thin', color='BFBFBF'),
            right=Side(style='thin', color='BFBFBF'),
            top=Side(style='thin', color='BFBFBF'),
            bottom=Side(style='thin', color='BFBFBF')
        )

        # Formata o Cabeçalho (Linha 1)
        # Cabeçalho
        for cell in ws[1]:
            cell.font = font_cabecalho
            cell.fill = fill_cabecalho
            cell.alignment = align_centro
            cell.border = borda_fina

        # Dados
        for row in ws.iter_rows(
                min_row=2,
                max_row=linha_total_index - 1,
                min_col=1,
                max_col=4
        ):
            for cell in row:
                cell.font = font_dados
                cell.border = borda_fina

                if cell.column == 1:
                    cell.alignment = align_esquerda
                else:
                    cell.alignment = align_centro

                if cell.column == 2:
                    cell.number_format = '"R$" #,##0.00'

                if cell.column == 3:
                    cell.number_format = 'DD/MM/YYYY'

        # Total
        for col in range(1, 5):
            cell = ws.cell(row=linha_total_index, column=col)

            cell.font = font_total
            cell.fill = fill_total
            cell.border = borda_fina

            if col == 1:
                cell.alignment = align_esquerda
            else:
                cell.alignment = align_centro

            if col == 2:
                cell.number_format = '"R$" #,##0.00'

        # Largura fixa
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 18

        # Altura padrão
        for row in range(1, ws.max_row + 1):
            ws.row_dimensions[row].height = 22

        # 7. Salva o arquivo e exibe a mensagem de sucesso
        try:
            wb.save(arquivo)
            messagebox.showinfo("Sucesso", "Relatório gerado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo:\n{e}")

    def nova_conta(self):
        JanelaConta(self.root, self.carregar_contas)

    def editar_conta(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma conta para editar.")
            return
        conta_id = self.tree.item(sel[0])['values'][0]
        JanelaConta(self.root, self.carregar_contas, conta_id)

    def excluir_conta(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma conta para excluir.")
            return

        # Correção: Pega o primeiro item selecionado da tupla do Treeview
        conta_id = sel[0]

        if messagebox.askyesno("Confirmar", "Deseja realmente excluir esta conta?"):
            conn = conectar()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM contas WHERE id = ?", (conta_id,))
            conn.commit()
            conn.close()

            # Recarrega a tabela na tela atualizando os dados
            self.carregar_contas()
            messagebox.showinfo("Sucesso", "Conta excluída com sucesso!")

    def abrir_fornecedores(self):
        JanelaFornecedores(self.root, self.carregar_contas)


if __name__ == "__main__":
    criar_tabelas()
    root = tk.Tk()
    app = JanelaPrincipal(root)
    root.mainloop()
