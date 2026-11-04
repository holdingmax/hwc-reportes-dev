-- PROPUESTA (no aplicada): paso 5 -- el "mes anterior" de JetSmart se lee de la
-- base en vez de pedirle a Anita que suba otra vez los archivos del mes pasado.
--
-- Estado: borrador para revision. Vive en sql/propuestas/ a proposito:
-- scripts/migrate.py aplica TODO lo que haya en sql/migrations/, asi que no se
-- mueve alla hasta que este aprobado. Solo base de Testing.
--
-- Aditiva: no toca ninguna tabla existente (cargas_archivo, liquidaciones,
-- movimientos_awb, ajustes_manuales, tarifas) ni los flujos de Avianca, Gol y
-- LATAM.
--
-- Que guarda: para cada mes liquidado (o cargado como arranque), TODAS las
-- guias de las dos fuentes -- las guias del sistema ("Guias Aerea") y la
-- rendicion de Ariel --, no solo las que se liquidaron. Las reglas del cruce
-- (JetSmart: "declarada dos veces", "pendiente del mes anterior", etc.) se
-- apoyan en saber que guias estaban en cada archivo del mes anterior,
-- incluidas las pendientes de declarar y las excluidas, que la liquidacion
-- guardada hoy (liquidaciones.detalle_totales) no conserva.

-- ---------------------------------------------------------------------------
-- Un renglon por cada "version" de un mes. Cada vez que se procesa de nuevo el
-- mismo mes (se rehace la liquidacion) se agrega una version nueva; la
-- vigente es la de mayor numero (ver la vista de abajo). Nunca se pisa nada:
-- queda el historial de versiones.
-- ---------------------------------------------------------------------------
CREATE TABLE jetsmart_periodos (
    id              BIGSERIAL PRIMARY KEY,
    -- Primer dia del mes liquidado (mes de vuelo = mes de la rendicion de
    -- Ariel). Mismo criterio que liquidaciones.periodo.
    periodo         DATE NOT NULL CHECK (EXTRACT(DAY FROM periodo) = 1),
    version         INTEGER NOT NULL CHECK (version >= 1),
    -- 'liquidacion': guardado al procesar el mes en la app.
    -- 'carga_inicial': mes que se liquido FUERA del sistema (ej. septiembre
    --   2026, antes de este cambio): se sube una sola vez para que el mes
    --   siguiente lo encuentre. No tiene liquidacion asociada.
    origen_dato     TEXT NOT NULL CHECK (origen_dato IN ('liquidacion', 'carga_inicial')),
    -- Liquidacion generada para esta version (NULL en la carga inicial). Si
    -- alguien borra esa liquidacion, las guias se conservan: son la memoria
    -- que usa el mes siguiente.
    liquidacion_id  BIGINT REFERENCES liquidaciones(id) ON DELETE SET NULL,
    -- Archivos de origen, para poder rastrear de donde salio cada mes:
    -- [{"nombre": "...", "hash": "<sha256>", "tipo": "guias_sistema|ariel"}]
    archivos        JSONB NOT NULL DEFAULT '[]'::jsonb,
    -- Filas repetidas identicas que se descartaron de la rendicion de Ariel (aviso).
    ariel_repetidas_identicas INTEGER NOT NULL DEFAULT 0,
    creado_en       TIMESTAMPTZ NOT NULL DEFAULT now(),
    creado_por      TEXT,
    UNIQUE (periodo, version)
);
CREATE INDEX idx_jetsmart_periodos_liquidacion ON jetsmart_periodos(liquidacion_id);

-- ---------------------------------------------------------------------------
-- Una fila por guia y por fuente. Una guia que esta en las dos fuentes
-- (el caso normal) tiene dos filas: una 'sistema' y una 'ariel'.
-- ---------------------------------------------------------------------------
CREATE TABLE jetsmart_guias_periodo (
    id            BIGSERIAL PRIMARY KEY,
    periodo_id    BIGINT NOT NULL REFERENCES jetsmart_periodos(id) ON DELETE CASCADE,
    fuente        TEXT NOT NULL CHECK (fuente IN ('sistema', 'ariel')),
    nro_guia      BIGINT NOT NULL,
    -- 'sistema': fecha de creacion. 'ariel': "Fecha de guia".
    fecha         DATE,
    origen        TEXT,                -- estacion de origen (AEP, COR, ...)
    kg            NUMERIC(14,3),
    -- $ Prioridad (sistema) / Ingreso total flete (Ariel). Sin redondear:
    -- el cruce compara sin tolerancia, ni siquiera centavos.
    ingreso       NUMERIC(16,4),
    estado        TEXT,                -- solo 'sistema' (Entregada, En vuelo, ...)
    -- Que paso con la guia en el cruce de ese mes (igual en sus dos filas):
    --   incluida  = entro a la liquidacion
    --   pendiente = esta en el sistema pero Ariel todavia no la declaro
    --               (se espera el mes que viene)
    --   excluida  = quedo afuera: declarada dos veces, sin confirmar, numero
    --               repetido en Ariel con datos distintos, o ya declarada el
    --               mes anterior
    --   (Regla que NO se puede expresar como CHECK, porque el origen vive en
    --   jetsmart_periodos: las versiones de origen 'liquidacion' SIEMPRE traen
    --   resultado en cada guia. Lo valida src/db.py antes de escribir y lo fija
    --   tests/test_jetsmart_mes_anterior.py.)
    --   NULL      = carga inicial: ese mes se liquido fuera del sistema y aca no
    --               se cruzo, asi que no hay resultado que guardar (no se inventa).
    --               Para el mes siguiente solo importa QUE guias estaban en cada fuente.
    resultado     TEXT CHECK (resultado IN ('incluida', 'pendiente', 'excluida')),
    -- Motivo exacto, con el mismo texto que muestra el cruce (ej. "match",
    -- "Pendiente de declarar", "Declarada también en el Ariel del mes anterior...").
    motivo        TEXT,
    -- Una guia aparece una sola vez por fuente y version. Si Ariel trae el mismo
    -- numero con datos distintos (ej. 97479 de agosto 2026) se guarda una fila
    -- (la primera) con resultado 'excluida'; para el mes siguiente solo
    -- importa que el numero estaba.
    UNIQUE (periodo_id, fuente, nro_guia)
);
CREATE INDEX idx_jetsmart_guias_nro ON jetsmart_guias_periodo(nro_guia);
CREATE INDEX idx_jetsmart_guias_periodo_fuente ON jetsmart_guias_periodo(periodo_id, fuente);

-- ---------------------------------------------------------------------------
-- Version vigente de cada mes: la de mayor numero. Es lo que lee la app para
-- armar el "mes anterior".
-- ---------------------------------------------------------------------------
CREATE VIEW jetsmart_periodo_vigente AS
SELECT DISTINCT ON (periodo) *
FROM jetsmart_periodos
ORDER BY periodo, version DESC;

-- ---------------------------------------------------------------------------
-- Rollback (NO forma parte de la migracion; solo referencia para revisarla):
--   DROP VIEW  IF EXISTS jetsmart_periodo_vigente;
--   DROP TABLE IF EXISTS jetsmart_guias_periodo;
--   DROP TABLE IF EXISTS jetsmart_periodos;
--   (y, si scripts/migrate.py ya la registro:
--    DELETE FROM schema_migrations WHERE filename = '002_jetsmart_guias_periodo.sql';)
-- ---------------------------------------------------------------------------
