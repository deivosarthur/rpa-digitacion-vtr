SELECT *
  FROM [Ctrl_Op_Logistica_N].[dbo].[proceso_digitacion_resumen]


ALTER TABLE proceso_digitacion_resumen
ADD id INT IDENTITY(1,1);

ALTER TABLE proceso_digitacion_resumen
ADD CONSTRAINT PK_proceso_digitacion PRIMARY KEY (id);

ALTER TABLE proceso_digitacion_resumen
ADD id INT IDENTITY(1,1);

ALTER TABLE proceso_digitacion_resumen
ADD bot VARCHAR(50) NULL;

ALTER TABLE proceso_digitacion_resumen
ADD estado_bot VARCHAR(50) NULL;

ALTER TABLE proceso_digitacion_resumen
ADD fecha_proceso DATETIME NULL;
