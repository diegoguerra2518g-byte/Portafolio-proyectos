PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS datasets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre_original TEXT NOT NULL,
    ruta_original TEXT,
    tipo_archivo TEXT NOT NULL CHECK (tipo_archivo IN ('CSV','XLSX','XLS','SQLITE')),
    tamano_bytes INTEGER CHECK (tamano_bytes IS NULL OR tamano_bytes >= 0),
    hash_sha256 TEXT,
    hoja_origen TEXT,
    encoding TEXT,
    separador TEXT,
    filas INTEGER CHECK (filas IS NULL OR filas >= 0),
    columnas INTEGER CHECK (columnas IS NULL OR columnas >= 0),
    estado TEXT NOT NULL DEFAULT 'REGISTRADO' CHECK (estado IN ('REGISTRADO','VALIDADO','RECHAZADO','ARCHIVADO')),
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE UNIQUE INDEX IF NOT EXISTS ux_datasets_hash_sha256
ON datasets(hash_sha256)
WHERE hash_sha256 IS NOT NULL;

CREATE TABLE IF NOT EXISTS columnas_dataset (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dataset_id INTEGER NOT NULL,
    posicion INTEGER NOT NULL CHECK (posicion >= 0),
    nombre_original TEXT NOT NULL,
    nombre_canonico TEXT,
    rol_detectado TEXT,
    tipo_detectado TEXT,
    tipo_objetivo TEXT,
    total_nulos INTEGER NOT NULL DEFAULT 0 CHECK (total_nulos >= 0),
    total_unicos INTEGER CHECK (total_unicos IS NULL OR total_unicos >= 0),
    minimo_texto TEXT,
    maximo_texto TEXT,
    ejemplos_json TEXT,
    estadisticas_json TEXT,
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (dataset_id) REFERENCES datasets(id) ON DELETE CASCADE,
    UNIQUE(dataset_id, posicion),
    UNIQUE(dataset_id, nombre_original)
);

CREATE TABLE IF NOT EXISTS ejecuciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dataset_id INTEGER NOT NULL,
    tipo TEXT NOT NULL DEFAULT 'ANALISIS' CHECK (tipo IN ('ANALISIS','LIMPIEZA','REANALISIS')),
    estado TEXT NOT NULL DEFAULT 'INICIADA' CHECK (estado IN ('INICIADA','ANALIZADA','PENDIENTE_REVISION','FINALIZADA','ERROR')),
    score_inicial REAL CHECK (score_inicial IS NULL OR (score_inicial >= 0 AND score_inicial <= 100)),
    score_final REAL CHECK (score_final IS NULL OR (score_final >= 0 AND score_final <= 100)),
    total_incidencias INTEGER NOT NULL DEFAULT 0 CHECK (total_incidencias >= 0),
    total_automaticas INTEGER NOT NULL DEFAULT 0 CHECK (total_automaticas >= 0),
    total_revision INTEGER NOT NULL DEFAULT 0 CHECK (total_revision >= 0),
    total_cambios INTEGER NOT NULL DEFAULT 0 CHECK (total_cambios >= 0),
    configuracion_json TEXT,
    ruta_resultado TEXT,
    mensaje_error TEXT,
    iniciada_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    finalizada_en TEXT,
    FOREIGN KEY (dataset_id) REFERENCES datasets(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS reglas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT NOT NULL UNIQUE,
    nombre TEXT NOT NULL,
    categoria TEXT NOT NULL,
    descripcion TEXT NOT NULL,
    alcance TEXT NOT NULL DEFAULT 'COLUMNA' CHECK (alcance IN ('DATASET','COLUMNA','FILA','CELDA')),
    severidad_predeterminada TEXT NOT NULL CHECK (severidad_predeterminada IN ('INFO','ADVERTENCIA','CRITICA')),
    tratamiento_predeterminado TEXT NOT NULL CHECK (tratamiento_predeterminado IN ('AUTOMATICA_SEGURA','REVISION_HUMANA')),
    configuracion_json TEXT,
    activa INTEGER NOT NULL DEFAULT 1 CHECK (activa IN (0,1)),
    es_sistema INTEGER NOT NULL DEFAULT 1 CHECK (es_sistema IN (0,1)),
    creada_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    actualizada_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS incidencias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ejecucion_id INTEGER NOT NULL,
    regla_id INTEGER NOT NULL,
    fila_numero INTEGER,
    registro_referencia TEXT,
    columna TEXT,
    valor_original TEXT,
    valor_sugerido TEXT,
    severidad TEXT NOT NULL CHECK (severidad IN ('INFO','ADVERTENCIA','CRITICA')),
    tratamiento TEXT NOT NULL CHECK (tratamiento IN ('AUTOMATICA_SEGURA','REVISION_HUMANA')),
    confianza REAL CHECK (confianza IS NULL OR (confianza >= 0 AND confianza <= 1)),
    estado TEXT NOT NULL DEFAULT 'ABIERTA' CHECK (estado IN ('ABIERTA','SUGERIDA','APLICADA','RESUELTA','IGNORADA')),
    descripcion TEXT NOT NULL,
    metadata_json TEXT,
    creada_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ejecucion_id) REFERENCES ejecuciones(id) ON DELETE CASCADE,
    FOREIGN KEY (regla_id) REFERENCES reglas(id)
);

CREATE TABLE IF NOT EXISTS revisiones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    incidencia_id INTEGER NOT NULL,
    decision TEXT NOT NULL CHECK (decision IN ('CORREGIR','ACEPTAR_SUGERENCIA','IGNORAR','ELIMINAR_REGISTRO','POSPONER')),
    valor_final TEXT,
    comentario TEXT,
    revisado_por TEXT DEFAULT 'usuario_local',
    revisada_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (incidencia_id) REFERENCES incidencias(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS cambios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ejecucion_id INTEGER NOT NULL,
    incidencia_id INTEGER,
    fila_numero INTEGER,
    registro_referencia TEXT,
    columna TEXT NOT NULL,
    regla_codigo TEXT NOT NULL,
    origen TEXT NOT NULL CHECK (origen IN ('AUTOMATICO','REVISION_HUMANA','CONFIGURACION')),
    valor_anterior TEXT,
    valor_nuevo TEXT,
    aplicado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ejecucion_id) REFERENCES ejecuciones(id) ON DELETE CASCADE,
    FOREIGN KEY (incidencia_id) REFERENCES incidencias(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS perfiles_mapeo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL UNIQUE,
    descripcion TEXT,
    firma_columnas TEXT,
    configuracion_json TEXT NOT NULL,
    activa INTEGER NOT NULL DEFAULT 1 CHECK (activa IN (0,1)),
    creada_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    actualizada_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS mapeos_ejecucion (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ejecucion_id INTEGER NOT NULL,
    perfil_id INTEGER,
    columna_origen TEXT NOT NULL,
    campo_canonico TEXT NOT NULL,
    metodo TEXT NOT NULL CHECK (metodo IN ('AUTOMATICO','MANUAL','PERFIL_GUARDADO')),
    confianza REAL CHECK (confianza IS NULL OR (confianza >= 0 AND confianza <= 1)),
    confirmado INTEGER NOT NULL DEFAULT 0 CHECK (confirmado IN (0,1)),
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ejecucion_id) REFERENCES ejecuciones(id) ON DELETE CASCADE,
    FOREIGN KEY (perfil_id) REFERENCES perfiles_mapeo(id) ON DELETE SET NULL,
    UNIQUE(ejecucion_id, columna_origen)
);

CREATE TABLE IF NOT EXISTS exportaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ejecucion_id INTEGER NOT NULL,
    formato TEXT NOT NULL CHECK (formato IN ('CSV','XLSX','JSON')),
    nombre_archivo TEXT NOT NULL,
    ruta_archivo TEXT NOT NULL,
    hash_sha256 TEXT,
    tamano_bytes INTEGER CHECK (tamano_bytes IS NULL OR tamano_bytes >= 0),
    creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (ejecucion_id) REFERENCES ejecuciones(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_columnas_dataset_dataset ON columnas_dataset(dataset_id);
CREATE INDEX IF NOT EXISTS idx_ejecuciones_dataset ON ejecuciones(dataset_id);
CREATE INDEX IF NOT EXISTS idx_incidencias_ejecucion ON incidencias(ejecucion_id);
CREATE INDEX IF NOT EXISTS idx_incidencias_regla ON incidencias(regla_id);
CREATE INDEX IF NOT EXISTS idx_incidencias_estado ON incidencias(estado);
CREATE INDEX IF NOT EXISTS idx_revisiones_incidencia ON revisiones(incidencia_id);
CREATE INDEX IF NOT EXISTS idx_cambios_ejecucion ON cambios(ejecucion_id);
CREATE INDEX IF NOT EXISTS idx_mapeos_ejecucion ON mapeos_ejecucion(ejecucion_id);
CREATE INDEX IF NOT EXISTS idx_exportaciones_ejecucion ON exportaciones(ejecucion_id);

INSERT OR IGNORE INTO reglas
(codigo,nombre,categoria,descripcion,alcance,severidad_predeterminada,tratamiento_predeterminado,configuracion_json)
VALUES
('R001','Espacios externos','TEXTO','Detecta espacios al inicio o al final de valores de texto.','CELDA','INFO','AUTOMATICA_SEGURA','{"accion":"strip"}'),
('R002','Espacios multiples','TEXTO','Detecta secuencias de dos o mas espacios dentro de un texto.','CELDA','INFO','AUTOMATICA_SEGURA','{"accion":"colapsar_espacios"}'),
('R003','Capitalizacion inconsistente','TEXTO','Detecta variaciones de mayusculas y minusculas en campos configurados como nombres o categorias.','CELDA','INFO','AUTOMATICA_SEGURA','{"accion":"normalizacion_contextual"}'),
('R004','Columna completamente vacia','ESTRUCTURA','Detecta columnas sin ningun valor util.','COLUMNA','ADVERTENCIA','REVISION_HUMANA','{"accion_sugerida":"eliminar_columna"}'),
('R005','Valor faltante','COMPLETITUD','Detecta valores nulos o vacios en campos evaluados.','CELDA','ADVERTENCIA','REVISION_HUMANA','{"permitir_columnas_opcionales":true}'),
('R006','Fila exactamente duplicada','DUPLICADOS','Detecta registros completamente identicos.','FILA','ADVERTENCIA','REVISION_HUMANA','{"accion_sugerida":"eliminar_copia","automatico_configurable":true}'),
('R007','Identificador repetido con datos diferentes','DUPLICADOS','Detecta la misma clave de negocio asociada a informacion no identica.','FILA','CRITICA','REVISION_HUMANA','{"requiere_campo_clave":true}'),
('R008','Correo invalido','FORMATO','Detecta valores que no cumplen una estructura basica de correo electronico.','CELDA','ADVERTENCIA','REVISION_HUMANA','{"validacion":"email_basico"}'),
('R009','Telefono invalido','FORMATO','Detecta telefonos demasiado cortos, texto no telefonico o formatos incompatibles con la configuracion.','CELDA','ADVERTENCIA','REVISION_HUMANA','{"longitud_minima":7}'),
('R010','Fecha invalida','FECHA','Detecta fechas imposibles o que no pueden interpretarse de forma confiable.','CELDA','CRITICA','REVISION_HUMANA','{"permitir_iso":true,"permitir_dmy":true}'),
('R011','Formato de fecha inconsistente','FECHA','Detecta formatos de fecha distintos dentro de una misma columna cuando son interpretables.','COLUMNA','INFO','AUTOMATICA_SEGURA','{"formato_salida":"YYYY-MM-DD"}'),
('R012','Numero almacenado como texto','TIPO','Detecta numeros representados como texto que pueden convertirse sin ambiguedad.','CELDA','INFO','AUTOMATICA_SEGURA','{"conversion":"numerica_segura"}'),
('R013','Valor negativo no permitido','RANGO','Detecta valores negativos en columnas configuradas como no negativas.','CELDA','CRITICA','REVISION_HUMANA','{"minimo":0}'),
('R014','Valor fuera de rango','RANGO','Detecta valores que exceden limites configurados para una columna.','CELDA','ADVERTENCIA','REVISION_HUMANA','{"limites_por_columna":true}'),
('R015','Categoria posiblemente equivalente','CATEGORIA','Detecta variantes de una misma categoria y propone una normalizacion cuando existe suficiente evidencia.','CELDA','INFO','REVISION_HUMANA','{"usar_similitud":true,"umbral_sugerencia":0.85}'),
('R016','Tipo de dato inconsistente','TIPO','Detecta valores incompatibles con el tipo dominante o esperado de la columna.','CELDA','ADVERTENCIA','REVISION_HUMANA','{"comparar_tipo_dominante":true}');
