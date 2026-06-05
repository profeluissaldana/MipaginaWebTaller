from app import app, db, Turno, Alumno, Asistencia, Clase, Auditoria
from datetime import datetime, date

def inicializar_base_de_datos():
    print("⏳ Iniciando la creación y carga de la base de datos...")
    
    with app.app_context():
    db.drop_all()  # Descomentado temporalmente para reiniciar la base de datos con la nueva estructura
    db.create_all()

    turnos_taller = [
        # MAÑANA (5°A)
        Turno(nombre_grupo="5°A - Mañana - Grupo A (Lun/Mie)", turno_horario="Mañana", dias_cursada="Lunes y Miércoles"), # Generará el ID: 1
        Turno(nombre_grupo="5°A - Mañana - Grupo B (Mar/Vie)", turno_horario="Mañana", dias_cursada="Martes y Viernes"), # Generará el ID: 2

        # TARDE (5°C)
        Turno(nombre_grupo="5°C - Tarde - Grupo Martes", turno_horario="Tarde", dias_cursada="Martes"),               # Generará el ID: 3
        Turno(nombre_grupo="5°C - Tarde - Grupo Miércoles", turno_horario="Tarde", dias_cursada="Miércoles"),         # Generará el ID: 4
        Turno(nombre_grupo="5°C - Tarde - Grupo Viernes", turno_horario="Tarde", dias_cursada="Viernes"),             # Generará el ID: 5
        Turno(nombre_grupo="5°C - Tarde - Rotativo Jueves", turno_horario="Tarde", dias_cursada="Jueves Rotativo")     # Generará el ID: 6
    ]
    db.session.add_all(turnos_taller)
    db.session.commit()

        # =========================================================================
        # 1. CARGA DE TURNOS (Configuración de tus horarios escolares)
        # =========================================================================
        turnos_taller = [
            Turno(nombre_grupo="Mañana - Grupo A (Lun/Mie)", turno_horario="Mañana", dias_cursada="Lunes y Miércoles"),
            Turno(nombre_grupo="Mañana - Grupo B (Mar/Vie)", turno_horario="Mañana", dias_cursada="Martes y Viernes"),
            Turno(nombre_grupo="Tarde - Grupo Fijo (Mar/Mie/Vie)", turno_horario="Tarde", dias_cursada="Martes, Miércoles y Viernes"),
            Turno(nombre_grupo="Tarde - Rotativo Jueves (Grupo 1)", turno_horario="Tarde", dias_cursada="Jueves Rotativo")
        ]
        db.session.add_all(turnos_taller)
        db.session.commit() # Confirmamos para obtener los IDs generados
        print("✅ Turnos y esquemas de rotación configurados.")

        # =========================================================================
        # 2. CARGA DE ALUMNOS DE PRUEBA
        # =========================================================================
        alumnos_prueba = [
            Alumno(
                usuario="luis", contrasena="1234", email="luis@tecnica.edu.ar",
                dni="45000111", apellido="Saldaña", nombre="Profesor Luis",
                observaciones_generales="Usuario de prueba del administrador.", turno_id=1
            ),
            Alumno(
                usuario="jgomez", contrasena="alumno5to", email="juan.gomez@tecnica.edu.ar",
                dni="48123456", apellido="Gómez", nombre="Juan",
                observaciones_generales="Presenta buen desempeño en las prácticas de Flask.", turno_id=1
            ),
            Alumno(
                usuario="mrodriguez", contrasena="clave789", email="maria.r@tecnica.edu.ar",
                dni="49765432", apellido="Rodríguez", nombre="María",
                observaciones_generales="Pertenece al grupo rotativo de la tarde.", turno_id=4
            )
        ]
        db.session.add_all(alumnos_prueba)
        db.session.commit()
        print("✅ Alumnos de prueba cargados (Usuarios: 'luis', 'jgomez', 'mrodriguez').")

        # =========================================================================
        # 3. CARGA DE HISTORIAL DE ASISTENCIAS (Simulación del ciclo lectivo)
        # Cargamos registros para Juan Gómez (ID: 2) para ver la analítica funcionar
        # =========================================================================
        print("⏳ Simulando historial de asistencia trimestral...")
        historial_juan = [
            # --- MARZO (1° Trimestre) ---
            Asistencia(fecha=date(2026, 3, 9), trimestre=1, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 3, 11), trimestre=1, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 3, 16), trimestre=1, estado='T', alumno_id=2, comentario_diario="Llegó 15 min tarde por problemas de transporte (Suma 0.5 falta)"),
            Asistencia(fecha=date(2026, 3, 18), trimestre=1, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 3, 23), trimestre=1, estado='A', alumno_id=2), # Ausente (Suma 1 falta)
            
            # --- ABRIL (1° Trimestre) ---
            Asistencia(fecha=date(2026, 4, 6), trimestre=1, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 4, 8), trimestre=1, estado='J', alumno_id=2, comentario_diario="Justificado: Presentó certificado médico por gripe (Suma 1 falta)"),
            Asistencia(fecha=date(2026, 4, 13), trimestre=1, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 4, 15), trimestre=1, estado='T', alumno_id=2, comentario_diario="Demorado en preceptoría (Suma 0.5 falta)"),
            
            # --- MAYO (1° Trimestre) ---
            Asistencia(fecha=date(2026, 5, 4), trimestre=1, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 5, 6), trimestre=1, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 5, 11), trimestre=1, estado='P', alumno_id=2),
            
            # --- JUNIO (Inicio del 2° Trimestre) ---
            Asistencia(fecha=date(2026, 6, 1), trimestre=2, estado='P', alumno_id=2),
            Asistencia(fecha=date(2026, 6, 3), trimestre=2, estado='P', alumno_id=2)
        ]
        
        db.session.add_all(historial_juan)
        
        # =========================================================================
        # 4. CARGA DEL LIBRO DE TEMAS (Seguimiento de Clases)
        # =========================================================================
        clases_dadas = [
            Clase(fecha=date(2026, 3, 9), tema_dado="Introducción a Flask", detalle_clase="Instalación del entorno virtual y arquitectura cliente-servidor.", turno_id=1),
            Clase(fecha=date(2026, 3, 16), tema_dado="Uso de plantillas Jinja2", detalle_clase="Explicación de bloques y herencia con base.html.", turno_id=1),
            Clase(fecha=date(2026, 4, 8), tema_dado="Introducción a Base de Datos", detalle_clase="Modelos relacionales y ORM con SQLAlchemy.", turno_id=1)
        ]
        db.session.add_all(clases_dadas)
        
        db.session.commit()
        print("✅ Historial de asistencias y libro de temas cargados correctamente.")
        print("\n🚀 ¡Todo listo! Ya pueden iniciar el servidor 'python app.py' y probar el sistema.")

if __name__ == '__main__':
    inicializar_base_de_datos()