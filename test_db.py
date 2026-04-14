from db_connection import obtener_ordenes, obtener_materiales_por_ot


ordenes = obtener_ordenes()


for orden in ordenes:
    print(f"OT: {orden.OT} - Tecnico: {orden.Tecnico}")
    
    materiales = obtener_materiales_por_ot(orden.OT)
    
    for material in materiales:
        print(f"  Codigo a rebajar: {material.Codigo_a_rebajar}, Tipo de Actividad: {material.Tipo_de_Actividad}, Tecnico: {material.Tecnico}")