# BASE DE DATOS

[sistema_eventos](https://docs.google.com/spreadsheets/d/1H_3ldHJ3O8O7TC-CNwcMnjDnhpuYnkRfDQh6lzrys3U/edit?usp=sharing)

```mermaid
erDiagram
    USUARIOS ||--o{ EVENTOS : "gestiona (EmailAsignado)"
    USUARIOS ||--o{ NOTIFICACIONES : "recibe (EmailDestino)"
    USUARIOS ||--o{ ASISTENCIA : "registra (Registrado_Por)"
    USUARIOS ||--o{ AUDITORIA : "genera (Usuario)"

    EVENTOS ||--o{ PARTICIPANTES : "agrupa (Grupo)"
    EVENTOS ||--o{ ACTIVIDADES : "contiene"
    EVENTOS ||--o{ ASISTENCIA : "registra"
    EVENTOS ||--o{ NOTIFICACIONES : "referencia (ID_Referencia)"
    
    PARTICIPANTES ||--o{ CALIFICACIONES : "obtiene"
    PARTICIPANTES ||--o{ ASISTENCIA : "atiende"
    
    ACTIVIDADES ||--o{ CALIFICACIONES : "evalúa"
    
    USUARIOS {
        string ID PK
        string Email UK
        string Password
        string Nombre
        string Rol
        string Token_Recuperacion
        datetime Expiracion_Token
    }

    EVENTOS {
        string ID_Evento PK
        string Nombre_Evento
        datetime inicio
        string Tipo
        string Email_Asignado FK "Relaciona con Usuarios.Email"
        int Cupo
        datetime fin
    }
PARTICIPANTES {
        string ID_Participante PK
        string Nombre
        string Correo
        string Grupo FK "Relacionado con ID_Evento"
        date Diploma_Enviado
    }

ACTIVIDADES {
        string ID_Actividad PK
        string ID_Evento FK
        string Titulo
        string Descripcion
        int Ponderacion
    }

    CALIFICACIONES {
        string ID_Nota PK
        string ID_Actividad FK
        string ID_Participante FK
        float Puntaje
        datetime Fecha_Registro
    }

    ASISTENCIA {
        string ID_Registro PK
        string ID_Evento FK
        string ID_Participante FK
        datetime Fecha_Hora
        string Estado
        string Registrado_Por FK
    }

    NOTIFICACIONES {
        string ID PK
        datetime Fecha
        string EmailDestino FK "Relaciona con Usuarios.Email"
        string Mensaje
        string Tipo
        string Estado
        string ID_Referencia FK "Relaciona con Eventos.ID_Evento"
    }

    AUDITORIA {
        datetime Fecha_Hora PK
        string Usuario FK "Relaciona con Usuarios.Email"
        string Accion
        string Modulo
        string Detalles
    }
```
## Resumen de las Entidades:
- Usuarios: Contiene la información de acceso y perfiles (Admin, Root, Docente).
- Eventos: Registra las actividades programadas, su tipo y el docente responsable asignado mediante su correo.
- Notificaciones: Almacena los mensajes enviados a los usuarios sobre eventos específicos (ID_Referencia).
- Auditoría: Registra el historial de acciones realizadas por los usuarios en el sistema (Login, Crear, Editar, etc.).
- Participantes: Gestiona la inscripción de las personas en los diferentes eventos.
- Actividades: Define tareas, exámenes o trabajos específicos que componen un evento.
- Asistencia: Registra la presencia de los participantes en las sesiones programadas.
- Calificaciones: Gestiona el rendimiento académico de los participantes en las actividades asignadas.
