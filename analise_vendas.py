# analise_vendas.py
import pandas as pd
import openpyxl
from openpyxl.chart import BarChart, Reference, PieChart
import matplotlib.pyplot as plt
import os
import sqlite3
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog

# Configuração do Banco de Dados SQLite
def conectar_banco():
    """Conecta ao banco SQLite e cria tabelas se não existirem"""
    conn = sqlite3.connect('database/vendas.db')
    cursor = conn.cursor()
    
    # Criação das tabelas
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS vendas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data TEXT NOT NULL,
        produto TEXT NOT NULL,
        quantidade INTEGER NOT NULL,
        valor_unitario REAL NOT NULL,
        valor_total REAL NOT NULL,
        regiao TEXT NOT NULL,
        data_registro TEXT DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS metadados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ultima_atualizacao TEXT,
        total_registros INTEGER
    )
    ''')
    
    conn.commit()
    return conn

def inserir_vendas(conn, df_vendas):
    """Insere dados de vendas no banco de dados"""
    cursor = conn.cursor()
    
    # Limpar tabela existente (opcional - comentar se quiser manter histórico)
    # cursor.execute('DELETE FROM vendas')
    
    # Inserir novos registros
    for _, row in df_vendas.iterrows():
        cursor.execute('''
        INSERT INTO vendas (data, produto, quantidade, valor_unitario, valor_total, regiao)
        VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            row['Data'].strftime('%Y-%m-%d'),
            row['Produto'],
            row['Quantidade'],
            row['Valor Unitário'],
            row['Valor Total'],
            row['Região']
        ))
    
    # Atualizar metadados
    cursor.execute('''
    INSERT INTO metadados (ultima_atualizacao, total_registros)
    VALUES (datetime('now'), (SELECT COUNT(*) FROM vendas))
    ''')
    
    conn.commit()
    return cursor.rowcount

def consultar_vendas(conn, query, params=None):
    """Executa consulta no banco de dados e retorna DataFrame"""
    try:
        if params:
            df = pd.read_sql_query(query, conn, params=params)
        else:
            df = pd.read_sql_query(query, conn)
        return df
    except Exception as e:
        print(f"Erro na consulta: {e}")
        return None

# Processamento de Dados
def processar_dados_vendas(arquivo_excel=None):
    """Processa dados de vendas e gera relatórios"""
    try:
        # Criar diretórios se não existirem
        os.makedirs('database', exist_ok=True)
        os.makedirs('data', exist_ok=True)
        os.makedirs('images', exist_ok=True)
        
        # Conectar ao banco de dados
        conn = conectar_banco()
        
        # Caminhos dos arquivos
        caminho_vendas = arquivo_excel if arquivo_excel else os.path.join('data', 'vendas.xlsx')
        caminho_dashboard = os.path.join('data', 'dashboard.xlsx')
        
        # Ler dados de vendas
        try:
            df_vendas = pd.read_excel(caminho_vendas)
            print("Dados de vendas carregados com sucesso!")
            
            # Calcular valor total se não existir
            if 'Valor Total' not in df_vendas.columns:
                df_vendas['Valor Total'] = df_vendas['Quantidade'] * df_vendas['Valor Unitário']
        except FileNotFoundError:
            print("Arquivo de vendas não encontrado. Criando dados de exemplo...")
            df_vendas = gerar_dados_exemplo()
            df_vendas.to_excel(caminho_vendas, index=False)
        
        # Inserir dados no banco
        registros_inseridos = inserir_vendas(conn, df_vendas)
        print(f"{registros_inseridos} registros inseridos no banco de dados.")
        
        # Consultas para análise
        resumo_produto = consultar_vendas(conn, '''
            SELECT produto, 
                   SUM(quantidade) as total_quantidade,
                   SUM(valor_total) as total_vendas,
                   ROUND(AVG(valor_unitario), 2) as preco_medio
            FROM vendas
            GROUP BY produto
            ORDER BY total_vendas DESC
        ''')
        
        resumo_regiao = consultar_vendas(conn, '''
            SELECT regiao, 
                   SUM(valor_total) as total_vendas,
                   COUNT(*) as quantidade_vendas
            FROM vendas
            GROUP BY regiao
            ORDER BY total_vendas DESC
        ''')
        
        resumo_mensal = consultar_vendas(conn, '''
            SELECT strftime('%Y-%m', data) as mes,
                   SUM(valor_total) as total_vendas,
                   COUNT(*) as quantidade_vendas
            FROM vendas
            GROUP BY mes
            ORDER BY mes
        ''')
        
        # Criar dashboard no Excel
        criar_dashboard_excel(caminho_dashboard, df_vendas, resumo_produto, resumo_regiao, resumo_mensal)
        
        # Gerar gráficos de imagem
        gerar_graficos_imagem(resumo_produto, resumo_regiao)
        
        # Fechar conexão com o banco
        conn.close()
        
        return True
        
    except Exception as e:
        print(f"Erro durante o processamento: {e}")
        if 'conn' in locals():
            conn.close()
        return False

def gerar_dados_exemplo():
    """Gera dados de exemplo para demonstração"""
    dados_exemplo = {
        'Data': pd.date_range(start='2023-01-01', periods=90, freq='D'),
        'Produto': ['A', 'B', 'C'] * 30,
        'Quantidade': [10, 15, 8, 20, 5, 12, 18, 7, 13, 9] * 9,
        'Valor Unitário': [100, 150, 80, 90, 200, 110, 85, 190, 120, 95] * 9,
        'Região': ['Norte', 'Sul', 'Leste', 'Oeste'] * 22 + ['Norte', 'Sul']
    }
    df = pd.DataFrame(dados_exemplo)
    df['Valor Total'] = df['Quantidade'] * df['Valor Unitário']
    return df

def criar_dashboard_excel(caminho, df_vendas, resumo_produto, resumo_regiao, resumo_mensal):
    """Cria arquivo Excel com dashboard de vendas"""
    with pd.ExcelWriter(caminho, engine='openpyxl') as writer:
        # Página com dados brutos
        df_vendas.to_excel(writer, sheet_name='Dados Brutos', index=False)
        
        # Páginas com resumos
        resumo_produto.to_excel(writer, sheet_name='Resumo por Produto', index=False)
        resumo_regiao.to_excel(writer, sheet_name='Resumo por Região', index=False)
        resumo_mensal.to_excel(writer, sheet_name='Vendas Mensais', index=False)
        
        # Acessar a planilha e o workbook para adicionar gráficos
        workbook = writer.book
        
        # Gráfico de barras para produtos
        sheet_produto = workbook['Resumo por Produto']
        chart_produto = BarChart()
        data = Reference(sheet_produto, min_col=3, max_col=4, min_row=1, max_row=resumo_produto.shape[0]+1)
        categories = Reference(sheet_produto, min_col=1, max_col=1, min_row=2, max_row=resumo_produto.shape[0]+1)
        
        chart_produto.add_data(data, titles_from_data=True)
        chart_produto.set_categories(categories)
        chart_produto.title = "Vendas por Produto"
        chart_produto.style = 10
        chart_produto.x_axis.title = "Produto"
        chart_produto.y_axis.title = "Valor"
        sheet_produto.add_chart(chart_produto, "F2")
        
        # Gráfico de pizza para regiões
        sheet_regiao = workbook['Resumo por Região']
        chart_regiao = PieChart()
        labels = Reference(sheet_regiao, min_col=1, max_col=1, min_row=2, max_row=resumo_regiao.shape[0]+1)
        data = Reference(sheet_regiao, min_col=2, max_col=2, min_row=1, max_row=resumo_regiao.shape[0]+1)
        
        chart_regiao.add_data(data, titles_from_data=True)
        chart_regiao.set_categories(labels)
        chart_regiao.title = "Distribuição por Região"
        sheet_regiao.add_chart(chart_regiao, "F2")
    
    print(f"Dashboard criado com sucesso em: {caminho}")

def gerar_graficos_imagem(resumo_produto, resumo_regiao):
    """Gera gráficos em formato de imagem para relatórios"""
    try:
        # Gráfico de barras - Vendas por produto
        plt.figure(figsize=(12, 6))
        plt.bar(resumo_produto['produto'], resumo_produto['total_vendas'])
        plt.title('Vendas por Produto (R$)')
        plt.xlabel('Produto')
        plt.ylabel('Valor Total')
        plt.grid(axis='y', linestyle='--')
        plt.savefig(os.path.join('images', 'vendas_produto.png'), bbox_inches='tight')
        plt.close()
        
        # Gráfico de pizza - Distribuição por região
        plt.figure(figsize=(8, 8))
        plt.pie(resumo_regiao['total_vendas'], 
                labels=resumo_regiao['regiao'], 
                autopct='%1.1f%%',
                startangle=90,
                shadow=True)
        plt.title('Distribuição de Vendas por Região')
        plt.savefig(os.path.join('images', 'vendas_regiao.png'), bbox_inches='tight')
        plt.close()
        
        print("Gráficos de imagem gerados com sucesso")
    except Exception as e:
        print(f"Erro ao gerar gráficos: {e}")

# Interface Gráfica
class AplicativoAnaliseVendas:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Análise de Vendas")
        self.root.geometry("500x300")
        
        self.criar_widgets()
    
    def criar_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Título
        tk.Label(main_frame, text="Análise de Vendas", font=('Arial', 16, 'bold')).pack(pady=10)
        
        # Seção de arquivo
        file_frame = tk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(file_frame, text="Arquivo de Vendas:").pack(side=tk.LEFT)
        
        self.entry_arquivo = tk.Entry(file_frame, width=30)
        self.entry_arquivo.pack(side=tk.LEFT, padx=5)
        
        tk.Button(file_frame, text="Procurar", command=self.selecionar_arquivo).pack(side=tk.LEFT)
       
        # Botão de processamento
        tk.Button(main_frame, 
                 text="Processar Dados e Gerar Dashboard", 
                 command=self.processar_dados,
                 bg='#4CAF50',
                 fg='white',
                 font=('Arial', 10, 'bold')).pack(pady=20)
       
        # Status
        self.status_label = tk.Label(main_frame, text="", fg='blue')
        self.status_label.pack()
        
        # Botão de saída
        tk.Button(main_frame, text="Sair", command=self.root.quit).pack(side=tk.BOTTOM, pady=10)
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo de vendas",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if arquivo:
            self.entry_arquivo.delete(0, tk.END)
            self.entry_arquivo.insert(0, arquivo)
    
    def processar_dados(self):
        arquivo = self.entry_arquivo.get()
        
        if not arquivo and not os.path.exists(os.path.join('data', 'vendas.xlsx')):
            resposta = messagebox.askyesno(
                "Arquivo não selecionado",
                "Nenhum arquivo foi selecionado. Deseja gerar dados de exemplo?"
            )
            if not resposta:
                return
        
        self.status_label.config(text="Processando dados...", fg='blue')
        self.root.update()
        
        try:
            sucesso = processar_dados_vendas(arquivo if arquivo else None)
            if sucesso:
                self.status_label.config(text="Dashboard gerado com sucesso!", fg='green')
                messagebox.showinfo(
                    "Sucesso",
                    "Dashboard gerado com sucesso!\n\n" +
                    f"Arquivo: data/dashboard.xlsx\n" +
                    f"Gráficos: pasta images/"
                )
            else:
                self.status_label.config(text="Erro ao processar dados", fg='red')
        except Exception as e:
            self.status_label.config(text=f"Erro: {str(e)}", fg='red')
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

# Função principal
def main():
    root = tk.Tk()
    app = AplicativoAnaliseVendas(root)
    root.mainloop()

if __name__ == "__main__":
    main()
    
    