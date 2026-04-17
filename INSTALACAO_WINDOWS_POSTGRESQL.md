# Passo a Passo de Instalacao

## 1. Preparar a maquina

1. Escolha a maquina Windows que vai funcionar como servidor.
2. Confirme que ela fica ligada durante o horario de uso.
3. Confirme que ela tem Excel instalado, se voce vai importar arquivos `.xlsx` ou `.xls`.
4. Crie uma pasta fixa para o projeto, por exemplo:

```powershell
D:\SISTEMAS\BANCO_CLIENTES_INSS
```

5. Copie a pasta `BANCO_CLIENTES_INSS` para esse local.

## 2. Instalar PostgreSQL

1. Baixe o instalador do PostgreSQL no site oficial:
   [https://www.postgresql.org/download/windows/](https://www.postgresql.org/download/windows/)
2. Instale com as opcoes padrao.
3. Defina uma senha para o usuario `postgres` e guarde essa senha.
4. Use a porta padrao `5432`.

## 3. Criar o banco

Abra o `SQL Shell (psql)` e execute:

```sql
CREATE DATABASE inss_clientes;
```

## 4. Instalar Python e dependencias

Abra PowerShell na pasta do projeto:

```powershell
cd D:\SISTEMAS\BANCO_CLIENTES_INSS
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 5. Configurar a conexao do sistema

No mesmo PowerShell, rode:

```powershell
$env:DATABASE_URL="postgresql+psycopg://postgres:SUA_SENHA@localhost:5432/inss_clientes"
```

Se quiser deixar permanente no Windows:

```powershell
setx DATABASE_URL "postgresql+psycopg://postgres:SUA_SENHA@localhost:5432/inss_clientes"
```

Depois feche e abra o PowerShell novamente.

## 6. Subir a aplicacao

Com o ambiente virtual ativo:

```powershell
python run_inss_app.py
```

Ou use:

```powershell
.\start_app.bat
```

Abra no navegador:

```text
http://localhost:8000
```

Se outra maquina da rede for acessar, descubra o IP do servidor:

```powershell
ipconfig
```

E abra:

```text
http://IP_DO_SERVIDOR:8000
```

## 7. Liberar acesso na rede

Se outras maquinas nao conseguirem abrir o sistema:

1. Abra o Firewall do Windows.
2. Libere a porta `8000` para a rede interna.
3. Confirme que o PostgreSQL esta local no servidor e a aplicacao esta rodando.

## 8. Primeiro uso

1. Abra `Importar lote`.
2. Informe `ano` e `mes`.
3. Preencha `Pasta no servidor` com a pasta onde estao as planilhas do mes.
4. Clique em `Importar arquivos`.
5. Confira os dados em `Clientes`.
6. Em `Gerar base`, prefira `Excel (.xlsx)` para evitar que o Excel altere moeda e data ao abrir.
7. Use `CSV` apenas quando precisar enviar para outro sistema.

## 9. Backup recomendado

Faca backup de:

```text
BANCO_CLIENTES_INSS\data\archive
BANCO_CLIENTES_INSS\data\exports
```

E do banco PostgreSQL com:

```powershell
pg_dump -U postgres -d inss_clientes -f D:\backup\inss_clientes.sql
```

## 10. Atualizacao futura

1. Pare a aplicacao.
2. Faca backup do banco.
3. Substitua os arquivos da pasta do projeto.
4. Ative o ambiente virtual.
5. Rode `pip install -r requirements.txt` novamente.
6. Suba a aplicacao de novo.
