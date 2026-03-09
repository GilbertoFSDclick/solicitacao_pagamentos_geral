-- =============================================================================
-- Script: criar_tabela_banco_controle.sql
-- Fluxo Prévio II - Banco de Controle da Automação
-- Conforme Detalhamento_RPA_Original_Lançamento_Solicitação_Pagamento_Geral
--
-- Como executar:
--   SQLite: sqlite3 controle_rpa.db < scripts/criar_tabela_banco_controle.sql
--   Ou: copie apenas a seção SQLite para seu cliente (DBeaver, etc.)
-- =============================================================================

-- -----------------------------------------------------------------------------
-- SQLite (url = sqlite:///controle_rpa.db)
-- -----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS rpa_controle_solicitacao_pagamento (
    id_processo            TEXT NOT NULL,
    Empresa_holmes          TEXT,
    Filial_holmes           TEXT,
    Cpf_holmes              TEXT,
    Num_nf_holmes           TEXT,
    Serie_nf_holmes         TEXT,
    "Emissão_nf_holmes"     TEXT,
    Entrada_nf_holmes       TEXT,
    Vlr_nf_holmes           TEXT,
    Protocolo_holmes        TEXT,
    Centro_custo_uab_holmes TEXT,
    Dt_Vencimento_Holmes    TEXT,
    Codigo_UAB_Contabil     TEXT,
    data_registro           TEXT DEFAULT (datetime('now', 'localtime')),
    PRIMARY KEY (id_processo)
);

-- Índices úteis para consultas
CREATE INDEX IF NOT EXISTS idx_rpa_controle_protocolo ON rpa_controle_solicitacao_pagamento(Protocolo_holmes);
CREATE INDEX IF NOT EXISTS idx_rpa_controle_data ON rpa_controle_solicitacao_pagamento(data_registro);


-- -----------------------------------------------------------------------------
-- SQL Server (url = mssql+pyodbc://user:pass@server/db?driver=ODBC+Driver+17+for+SQL+Server)
-- Descomente e adapte se usar SQL Server
-- -----------------------------------------------------------------------------
/*
CREATE TABLE rpa_controle_solicitacao_pagamento (
    id_processo            NVARCHAR(100) NOT NULL,
    Empresa_holmes         NVARCHAR(50),
    Filial_holmes           NVARCHAR(50),
    Cpf_holmes             NVARCHAR(20),
    Num_nf_holmes           NVARCHAR(50),
    Serie_nf_holmes         NVARCHAR(10),
    [Emissão_nf_holmes]     NVARCHAR(20),
    Entrada_nf_holmes       NVARCHAR(20),
    Vlr_nf_holmes           NVARCHAR(20),
    Protocolo_holmes        NVARCHAR(100),
    Centro_custo_uab_holmes NVARCHAR(50),
    Dt_Vencimento_Holmes    NVARCHAR(20),
    Codigo_UAB_Contabil     NVARCHAR(20),
    data_registro           DATETIME2 DEFAULT GETDATE(),
    PRIMARY KEY (id_processo)
);

CREATE INDEX idx_rpa_controle_protocolo ON rpa_controle_solicitacao_pagamento(Protocolo_holmes);
CREATE INDEX idx_rpa_controle_data ON rpa_controle_solicitacao_pagamento(data_registro);
*/
