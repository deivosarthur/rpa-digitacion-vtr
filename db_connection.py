import pyodbc



def get_db_connection():
    conn = pyodbc.connect(
        'DRIVER={SQL Server};'
        'SERVER=45.239.111.133;'
        'DATABASE=Ctrl_Op_Logistica_N;'
        'UID=AdolfoP;'
        'PWD=p7@q$rA!sB^tC'
    )
    return conn

def obtener_ordenes():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    query="""
    SELECT id, OT, Tecnico
    FROM proceso_digitacion_resumen
    WHERE Clasificacion_proceso = 'Al BOT'
      AND Estado_digitacion = 'Pendiente'
    """
      
    cursor.execute(query)
    rows = cursor.fetchall()
    
    conn.close()
    return rows 

def obtener_materiales_por_ot(ot):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    query="""
    SELECT Codigo_a_rebajar, Tipo_de_Actividad, Tecnico, Declarado
    FROM proceso_digitacion_detalle_materiales
    WHERE OT = ?
    """
      
    cursor.execute(query, ot)
    rows = cursor.fetchall()
    
    conn.close()
    return rows

def actualizar_estado(ot, estado):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    query="""
    UPDATE proceso_digitacion_resumen
    SET Estado_digitacion = 'Procesado'
    WHERE OT = ?
    """
    
    cursor.execute(query, ot)
    conn.commit()
    
    conn.close()