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
4. Em producao, defina tambem os segredos: `APP_SESSION_SECRET`, `APP_DEFAULT_ADMIN_PASSWORD`, `C6_USERNAME`, `C6_PASSWORD`.
4. Rode `python run_inss_app.py`.
5. Abra `http://localhost:8000`.

## Fila de jobs (Fase 3)

O sistema suporta dois modos de fila:

- `APP_QUEUE_BACKEND=thread` (padrao): processa jobs em thread local.
- `APP_QUEUE_BACKEND=rq`: processa jobs no Redis com worker dedicado.

Variaveis principais:

```powershell
$env:APP_QUEUE_BACKEND="rq"
$env:APP_QUEUE_REDIS_URL="redis://127.0.0.1:6379/0"
$env:APP_QUEUE_EXPORTS_NAME="inss_exports"
$env:APP_QUEUE_IMPORTS_NAME="inss_imports"
$env:APP_QUEUE_DEFAULT_NAME="inss_default"
```

### Subir worker RQ

```powershell
python run_rq_worker.py
```

Importante: o worker RQ precisa de Redis ativo em `APP_QUEUE_REDIS_URL`.
Se o Redis estiver fora, o sistema segue funcionando em fallback `thread`.

Atalhos prontos:

- API: `.\start_producao.ps1`
- Worker: `.\start_worker_producao.ps1`
- API + Worker: `.\start_producao_com_worker.ps1`

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

## Validacoes obrigatorias de producao

Quando `APP_ENV=production`, o sistema bloqueia a inicializacao se detectar configuracao insegura.

Regras obrigatorias:

- `APP_SESSION_SECRET` nao pode ser padrao/fraco e deve ter no minimo 32 caracteres.
- `APP_DEFAULT_ADMIN_PASSWORD` nao pode ser padrao/fraca e deve ter no minimo 12 caracteres.

Se alguma regra falhar, o app encerra com erro de configuracao para evitar subida insegura em producao.

## Fluxo operacional

1. Entre em `Importar lote`.
2. Informe `ano` e `mes`.
3. Use `Pasta no servidor` para importar uma pasta inteira ou envie varios arquivos.
4. Consulte em `Clientes`.
5. Gere `.xlsx` em `Gerar base` para abrir no Excel sem distorcer moeda e data.
6. Use `.csv` apenas quando precisar do arquivo bruto para outro sistema.

## XLSX e XLS

Para importar `.xlsx` e `.xls`, o Windows do servidor precisa ter Excel instalado, porque a conversao usa automacao COM do Excel.

## Retencao automatica (archive e logs)

Existe limpeza automatica para reduzir uso de disco em producao:

- Script: `housekeeping_cleanup.ps1`
- Escopo: `data/archive` e arquivos `*.log`
- Retencao padrao: `90` dias
- Log de execucao: `logs/housekeeping_cleanup.log`

### Testar sem apagar arquivos

```powershell
.\housekeeping_cleanup.ps1 -RetentionDays 90 -DryRun
```

### Agendar no Task Scheduler (Windows)

```powershell
.\setup_housekeeping_task.ps1 -TaskName "INSS_Housekeeping_Cleanup" -RetentionDays 90 -DailyAt "02:30"
```

### Verificar se a task executou

```powershell
Get-ScheduledTask -TaskName "INSS_Housekeeping_Cleanup" | Get-ScheduledTaskInfo
Get-Content .\logs\housekeeping_cleanup.log -Tail 30
```

## Indices recomendados (PostgreSQL)

Para melhorar desempenho de filtros e exportacoes em volume alto:

1. Rode o script `sql/phase1_postgres_indexes.sql` no banco PostgreSQL.
2. Execute primeiro no ambiente de teste.
3. Em producao, rode em janela de menor carga.

Atalho PowerShell:

```powershell
.\apply_phase1_indexes.ps1
```

## Ajuste de lote da exportacao

Para ajustar o tamanho de lote da exportacao sem alterar codigo:

```powershell
$env:APP_EXPORT_FETCH_SIZE="1000"
```

Faixa recomendada: `500` a `5000` (padrao `1000`).
