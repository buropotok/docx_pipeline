-- БД translate_factory.sqlite

-- country определение

CREATE TABLE country (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    title_ru TEXT NOT NULL,
    title_en TEXT,
    comment TEXT,
    is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now'))
);


-- document_type определение

CREATE TABLE document_type (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    title TEXT NOT NULL UNIQUE,
    comment TEXT,
    cluster_key TEXT,
    is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now'))
);


-- font определение

CREATE TABLE font (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    title             TEXT NOT NULL,
    normalized_title  TEXT NOT NULL,
    created_at        TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at        TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE UNIQUE INDEX ux_font_normalized_title
    ON font(normalized_title);


-- font_scan_run определение

CREATE TABLE font_scan_run (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at            TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    finished_at           TEXT,
    status                TEXT NOT NULL,
    scanned_documents     INTEGER NOT NULL DEFAULT 0,
    inserted_fonts        INTEGER NOT NULL DEFAULT 0,
    inserted_font_usages  INTEGER NOT NULL DEFAULT 0,
    inserted_links        INTEGER NOT NULL DEFAULT 0,
    error_message         TEXT
);


-- "language" определение

CREATE TABLE language (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    title TEXT NOT NULL UNIQUE,
    comment TEXT,
    is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now'))
);


-- tag определение

CREATE TABLE tag (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE,
    title TEXT NOT NULL,
    comment TEXT,
    example_value TEXT,
    scope TEXT NOT NULL DEFAULT 'generic',
    is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now'))
);


-- document определение

CREATE TABLE document (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    uid TEXT NOT NULL UNIQUE,
    mail_uid TEXT NOT NULL,
    attachment_index INTEGER NOT NULL,
    source_filename TEXT,
    source_ext TEXT,
    source_abs_path TEXT,
    artifacts_abs_path TEXT,
    country_id INTEGER,
    document_type_id INTEGER,
    processing_status TEXT NOT NULL DEFAULT 'discovered',
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (country_id) REFERENCES country(id),
    FOREIGN KEY (document_type_id) REFERENCES document_type(id),
    UNIQUE (mail_uid, attachment_index)
);

CREATE INDEX idx_document_mail_uid ON document(mail_uid);


-- document_original_language определение

CREATE TABLE document_original_language (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    language_id INTEGER NOT NULL,
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (document_id) REFERENCES document(id) ON DELETE CASCADE,
    FOREIGN KEY (language_id) REFERENCES language(id),
    UNIQUE (document_id, language_id)
);

CREATE INDEX idx_document_original_language_document_id
    ON document_original_language(document_id);
CREATE INDEX idx_document_original_language_language_id
    ON document_original_language(language_id);


-- document_tag определение

CREATE TABLE document_tag (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    tag_id INTEGER NOT NULL,
    source TEXT NOT NULL DEFAULT 'ai',
    confidence REAL,
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (document_id) REFERENCES document(id) ON DELETE CASCADE,
    FOREIGN KEY (tag_id) REFERENCES tag(id),
    UNIQUE (document_id, tag_id)
);

CREATE INDEX idx_document_tag_document_id ON document_tag(document_id);
CREATE INDEX idx_document_tag_tag_id ON document_tag(tag_id);


-- document_translate_language определение

CREATE TABLE document_translate_language (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    language_id INTEGER NOT NULL,
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (document_id) REFERENCES document(id) ON DELETE CASCADE,
    FOREIGN KEY (language_id) REFERENCES language(id),
    UNIQUE (document_id, language_id)
);

CREATE INDEX idx_document_translate_language_document_id
    ON document_translate_language(document_id);
CREATE INDEX idx_document_translate_language_language_id
    ON document_translate_language(language_id);


-- document_type_tag определение

CREATE TABLE document_type_tag (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_type_id INTEGER NOT NULL,
    tag_id INTEGER NOT NULL,
    is_required INTEGER NOT NULL DEFAULT 0 CHECK (is_required IN (0, 1)),
    is_recommended INTEGER NOT NULL DEFAULT 1 CHECK (is_recommended IN (0, 1)),
    comment TEXT,
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (document_type_id) REFERENCES document_type(id) ON DELETE CASCADE,
    FOREIGN KEY (tag_id) REFERENCES tag(id),
    UNIQUE (document_type_id, tag_id)
);

CREATE INDEX idx_document_type_tag_document_type_id
    ON document_type_tag(document_type_id);
CREATE INDEX idx_document_type_tag_tag_id
    ON document_type_tag(tag_id);


-- fingerprint определение

CREATE TABLE fingerprint (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    fingerprint_version TEXT NOT NULL,
    artifact_path TEXT NOT NULL,
    tags_artifact_path TEXT,
    fingerprint_hash TEXT,
    status TEXT NOT NULL DEFAULT 'created',
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (document_id) REFERENCES document(id) ON DELETE CASCADE
);

CREATE INDEX idx_fingerprint_document_id ON fingerprint(document_id);


-- font_metric определение

CREATE TABLE font_metric (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    font_id           INTEGER NOT NULL,
    base_size_pt      REAL NOT NULL DEFAULT 12.0,
    char_code         INTEGER NOT NULL,
    char_text         TEXT NOT NULL,
    width_units       REAL NOT NULL,
    created_at        TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at        TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, font_title VARCHAR(50), source_font VARCHAR(50),
    FOREIGN KEY (font_id) REFERENCES font(id)
);


-- font_usage определение

CREATE TABLE font_usage (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    font_id             INTEGER NOT NULL,
    size_pt             REAL,
    bold                INTEGER,
    italic              INTEGER,
    underline           INTEGER,
    all_caps            INTEGER,
    small_caps          INTEGER,
    strike              INTEGER,
    double_strike       INTEGER,
    outline             INTEGER,
    shadow              INTEGER,
    emboss              INTEGER,
    imprint             INTEGER,
    rtl                 INTEGER,
    lang                TEXT,
    created_at          TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at          TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,

    FOREIGN KEY (font_id) REFERENCES font(id)
);

CREATE UNIQUE INDEX ux_font_usage_signature
    ON font_usage (
        font_id,
        COALESCE(size_pt, -1),
        COALESCE(bold, -1),
        COALESCE(italic, -1),
        COALESCE(underline, -1),
        COALESCE(all_caps, -1),
        COALESCE(small_caps, -1),
        COALESCE(strike, -1),
        COALESCE(double_strike, -1),
        COALESCE(outline, -1),
        COALESCE(shadow, -1),
        COALESCE(emboss, -1),
        COALESCE(imprint, -1),
        COALESCE(rtl, -1),
        COALESCE(lang, '')
    );


-- classification_run определение

CREATE TABLE classification_run (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL,
    model_name TEXT,
    prompt_version TEXT,
    input_artifact_path TEXT,
    response_artifact_path TEXT,
    result_artifact_path TEXT,
    status TEXT NOT NULL DEFAULT 'created',
    error_message TEXT,
    tokens_input INTEGER,
    tokens_output INTEGER,
    cost_estimate REAL,
    duration_ms INTEGER,
    create_date TEXT NOT NULL DEFAULT (datetime('now')),
    update_date TEXT NOT NULL DEFAULT (datetime('now')),

    FOREIGN KEY (document_id) REFERENCES document(id) ON DELETE CASCADE
);

CREATE INDEX idx_classification_run_document_id ON classification_run(document_id);


-- document_font_profile определение

CREATE TABLE document_font_profile (
    document_id              INTEGER PRIMARY KEY,
    primary_font_usage_id    INTEGER,
    total_runs_count         INTEGER NOT NULL DEFAULT 0,
    total_chars_count        INTEGER NOT NULL DEFAULT 0,
    unique_font_usages_count INTEGER NOT NULL DEFAULT 0,
    unique_fonts_count       INTEGER NOT NULL DEFAULT 0,
    created_at               TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at               TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,

    FOREIGN KEY (document_id) REFERENCES document(id),
    FOREIGN KEY (primary_font_usage_id) REFERENCES font_usage(id)
);


-- document_font_usage определение

CREATE TABLE document_font_usage (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id         INTEGER NOT NULL,
    font_usage_id       INTEGER NOT NULL,
    runs_count          INTEGER NOT NULL DEFAULT 0,
    chars_count         INTEGER NOT NULL DEFAULT 0,
    paragraphs_count    INTEGER NOT NULL DEFAULT 0,
    tables_count        INTEGER NOT NULL DEFAULT 0,
    first_seen_path     TEXT,
    created_at          TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at          TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, scan_run_id INTEGER,

    FOREIGN KEY (document_id) REFERENCES document(id),
    FOREIGN KEY (font_usage_id) REFERENCES font_usage(id)
);

CREATE UNIQUE INDEX ux_document_font_usage
    ON document_font_usage(document_id, font_usage_id);
CREATE INDEX ix_document_font_usage_document_id
    ON document_font_usage(document_id);
CREATE INDEX ix_document_font_usage_font_usage_id
    ON document_font_usage(font_usage_id);