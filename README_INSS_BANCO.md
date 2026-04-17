# Banco de Clientes INSS

Projeto local para importar planilhas de clientes INSS, guardar historico por arquivo, pesquisar clientes e gerar novas bases CSV filtradas.

## Estrutura da pasta

- `inss_db_app/`: aplicacao FastAPI
- `tests/`: testes automatizados
- `data/uploads/`: uploads recebidos pela interface
- `data/archive/`: copia dos arquivos importados
- `data/exports/`: CSVs gerados
- `run_inss_app.py`: inicializacao da aplicacao
- `start_app.bat`: atalho para subir o sistema no Windows
- `requirements.txt`: dependencias Python
- `inss_clientes.db`: banco local SQLite usado apenas como fallback de desenvolvimento

## Banco recomendado

Use `PostgreSQL` no servidor.

Se `DATABASE_URL` nao for definido, o sistema usa `SQLite` local automaticamente. Isso serve para teste rapido, nao para producao.

## Executar no servidor

1. Abra o PowerShell dentro da pasta `BANCO_CLIENTES_INSS`.
2. Ative o ambiente virtual.
3. Defina a variavel `DATABASE_URL` ou use `PGHOST/PGPORT/PGUSER/PGPASSWORD/PGDATABASE`.
4. Rode `python run_inss_app.py`.
5. Abra `http://localhost:8000`.

## Exemplo de conexao PostgreSQL

```powershell
$env:DATABASE_URL="postgresql+psycopg://postgres:senha@localhost:5432/inss_clientes"
```

## Criar banco PostgreSQL novo

Se quiser recriar do zero e reimportar as planilhas:

1. Rode `.\setup_postgres.ps1` no PowerShell.
2. Informe host, porta, usuario admin, senha e nome do banco.
3. O script cria o banco novo e deixa `PGHOST`, `PGPORT`, `PGUSER`, `PGPASSWORD` e `PGDATABASE` definidos na sessao atual.
4. Ainda na mesma janela, rode `.\.venv\Scripts\python.exe .\run_inss_app.py`.
5. Reimporte os arquivos pela interface.

Se preferir usar um arquivo de exemplo, copie `.env.example` e preencha os valores antes de exportar as variaveis no PowerShell.

## Fluxo operacional

1. Entre em `Importar lote`.
2. Informe `ano` e `mes`.
3. Use `Pasta no servidor` para importar uma pasta inteira ou envie varios arquivos.
4. Consulte em `Clientes`.
5. Gere `.xlsx` em `Gerar base` para abrir no Excel sem distorcer moeda e data.
6. Use `.csv` apenas quando precisar do arquivo bruto para outro sistema.

## XLSX e XLS

Para importar `.xlsx` e `.xls`, o Windows do servidor precisa ter Excel instalado, porque a conversao usa automacao COM do Excel.
