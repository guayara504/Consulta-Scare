import mysql.connector as sql



class Database_1:
    

    def __init__(self):
        self.conexion = sql.connect(
                    host = "sistemajudicial.com", 
                    user = "diego", 
                    passwd = "S1st3m4w3b",
                    database = "redjudicial")
    
    def consulta(self):
        consulta =   """SELECT z04_estado.z01_radicacion_juzgado AS Juzgado_Estado,z04_estado.z01_radicacion_z01_radicacion AS Radicacion_Estado,z04_estado.demandante AS Demandante_Estado,
		                z01_radicacion_has_z02_abogado.demandante AS Demandante_Cliente,z04_estado.demandado AS Demandado_Estado,z01_radicacion_has_z02_abogado.demandado AS Demandado_Cliente,
		                z04_estado.ciudad, z04_estado.fecha_notificacion,z04_estado.clase_proceso AS Actuacion_Estado,
                        z01_radicacion_has_z02_abogado.z02_abogado_idcedula_z02 AS Cedula_Cliente,
	                	CONCAT(z02_abogado.nombre, ' ', z02_abogado.apellido) Nombre_Cliente
                        FROM z04_estado
		                JOIN z01_radicacion_has_z02_abogado ON z01_radicacion_has_z02_abogado.z01_radicacion_juzgado=z04_estado.z01_radicacion_juzgado
                        AND z01_radicacion_has_z02_abogado.z01_radicacion_z01_radicacion=z04_estado.z01_radicacion_z01_radicacion
                        AND z01_radicacion_has_z02_abogado.ciudad=z04_estado.ciudad
                        JOIN z02_abogado ON z01_radicacion_has_z02_abogado.z02_abogado_idcedula_z02=z02_abogado.idcedula_z02
                        WHERE z02_abogado.privilegio=1 and z04_estado.fecha_notificacion = CURDATE() AND z02_abogado.idcedula_z02 like "860020082%" """

        try:
            print("-"*15)
            print("Consultando")
            cursor = self.conexion.cursor()
            cursor.execute(consulta)
            consultasBD  = [item for item in cursor.fetchall()]
            print("Finalizó la consulta")
            print("-"*15)
        except:
            print("\nNo Hubo conexion con la base de datos...")
        
        return consultasBD
