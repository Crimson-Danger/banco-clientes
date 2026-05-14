-- Fase 1: indices para consultas e exportacoes mais frequentes.
-- Execute em PostgreSQL no banco de teste primeiro.
-- Em producao, prefira executar fora de transacao devido ao CONCURRENTLY.

CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_occurrence_segment_uf_city
ON client_occurrences (base_segment, uf, cidade);

CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_occurrence_segment_margin
ON client_occurrences (base_segment, vl_margem);

CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_occurrence_segment_cpf
ON client_occurrences (base_segment, cpf);

CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_occurrence_segment_matricula
ON client_occurrences (base_segment, matricula);

CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_batches_ref
ON import_batches (ano_referencia DESC, mes_referencia DESC);

CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_source_file_batch_week
ON source_files (batch_id, semana_label);

ANALYZE client_occurrences;
ANALYZE import_batches;
ANALYZE source_files;
