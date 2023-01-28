import sqlite3

# Conectar ao banco de dados
conn = sqlite3.connect("equipe.db")

# Crie um cursor para executar comandos SQL
cursor = conn.cursor()

# Criar a tabela "equipe"
cursor.execute("""CREATE TABLE equipe (
                    id INTEGER PRIMARY KEY,
                    nome TEXT,
                    email TEXT,
                    celular TEXT
                )""")

# Salvar alterações
conn.commit()

# Fechar conexão
conn.close()