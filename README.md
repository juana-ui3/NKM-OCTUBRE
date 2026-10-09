# Control de Asistencia (Django)

![Python](https://img.shields.io/badge/python-3.13-blue.svg)
![Django](https://img.shields.io/badge/django-5.x-green.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

Aplicación web robusta para la gestión y registro de asistencia de empleados. Diseñada para ofrecer un control preciso mediante validación por huella digital (simulada vía FingerprintJS) y generación de reportes detallados.

## Características Principales

### Gestión de Asistencia
- *Registro de Eventos*: Entrada, Salida, Inicio/Fin de Almuerzo, Comisión y Otros.
- *Validación por Dispositivo*: Integración con FingerprintJS para asegurar que un dispositivo no registre asistencia para múltiples empleados el mismo día.
- *Reglas de Negocio*: Validación de eventos únicos por día y coherencia temporal.

### Reportes y Exportación
- *Excel Detallado*: Exportación completa de todos los registros de asistencia.
- *Resumen Diario*: Cálculo automático de horas trabajadas, tiempos de almuerzo y comisiones.

### Acceso y Seguridad
- *Panel Administrativo*: Acceso restringido para descarga de reportes (usuarios `is_staff`).
- *Códigos QR*: Generación de QR para acceso rápido a los formularios de registro.
- *APIs*: Endpoints para integración con dispositivos y aplicaciones externas.

## Tecnologías

- *Backend*: Django 5.x
- *Base de Datos*: PostgreSQL / Supabase (obligatorio, sin SQLite)
- *Estáticos*: WhiteNoise
- *Servidor*: Gunicorn
- *Utilidades*: OpenPyXL (Excel), qrcode, python-dotenv

---

## Configuración del Entorno

Dado que este proyecto maneja información sensible, **en desarrollo se usa `.env`; en Railway configura las variables del servicio** en la raíz del proyecto.

Crea un archivo llamado `.env` y agrega las siguientes variables:

```env
# --- Configuración General ---
# Clave secreta de Django (Generar una nueva para producción)
SECRET_KEY=tu_clave_secreta_super_segura_aqui

# Modo Debug (True para desarrollo, False para producción)
DEBUG=True

# --- Base de Datos ---
# Obligatoria: copiar la URL del Session pooler del proyecto Supabase.
# Sustituir los marcadores; codificar la contrasena en formato URL.
DATABASE_URL=postgresql://postgres.PROJECT_REF:PASSWORD_URL_ENCODED@POOLER_HOST:5432/postgres

# Opcional: cargar la lista incluida de 27 empleados y 8 tipos al arrancar.
# Activar solo tras confirmar que la base destino y la lista son correctas.
CARGAR_DATOS_INICIALES=False

# --- Zona Horaria ---
TIME_ZONE=America/Lima
```

> [!IMPORTANT]
> Nunca subas tu archivo `.env` al repositorio. Asegúrate de que esté en `.gitignore`.

---

## Documentación de API

El sistema expone endpoints para trabajar con la identificación por dispositivo (fingerprint):

### Identificar por Fingerprint
Busca un empleado asociado a un dispositivo específico.

- *URL*: `/api/identificar-fingerprint/`
- *Método*: `POST`
- *Body*:

```json
{
  "fingerprint": "hash_del_dispositivo"
}
```

### Vincular Fingerprint
Asocia un dispositivo a un empleado (registro inicial o reasignación).

- *URL*: `/api/vincular-fingerprint/`
- *Método*: `POST`
- *Body*:

```json
{
  "empleado_id": 1,
  "fingerprint": "hash_del_dispositivo"
}
```

### Desvincular Fingerprint
Elimina la asociación de un dispositivo para permitir volver a seleccionar empleado.

- *URL*: `/api/desvincular-fingerprint/`
- *Método*: `POST`
- *Body*:

```json
{
  "fingerprint": "hash_del_dispositivo"
}
```

---

## Instalación y Ejecución Local

1. **Clonar el repositorio**

   ```bash
   git clone https://github.com/thaliat3/asistencia-nakama-service.git
   cd asistencia-nakama-service
   ```

2. **Crear entorno virtual**

   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # Linux/Mac
   source .venv/bin/activate
   ```

3. **Instalar dependencias**

   ```bash
   pip install -r requirements.txt
   ```

4. **Configurar entorno**
   - Crea el archivo `.env` como se indicó en la sección de configuración.

5. **Migraciones y Superusuario**

   ```bash
   python manage.py migrate
   python manage.py createsuperuser
   ```

6. **Cargar datos de prueba**

   ```bash
   python cargar_empleados.py
   ```

7. **Ejecutar servidor**

   ```bash
   python manage.py runserver
   ```

Visita `http://127.0.0.1:8000/` para ver la aplicación.

---

## Despliegue (Railway/Render)

El proyecto incluye `Procfile` y configuración para despliegue en la nube.

1. Configura las variables de entorno en el panel de tu proveedor (`DATABASE_URL`, `SECRET_KEY`, etc.).
2. El comando de inicio automático es:

   ```bash
   python manage.py migrate && gunicorn control_asistencia.wsgi
   ```

3. Los archivos estáticos son servidos automáticamente por WhiteNoise.

---

## Contribuir

1. Haz un Fork del proyecto.
2. Crea una rama para tu funcionalidad:

   ```bash
   git checkout -b feature/nueva-funcionalidad
   ```

3. Haz Commit de tus cambios:

   ```bash
   git commit -m "Agrega nueva funcionalidad"
   ```

4. Haz Push a la rama:

   ```bash
   git push origin feature/nueva-funcionalidad
   ```

5. Abre un Pull Request.


## Conexion y despliegue en Railway

SQLite fue eliminado de la configuracion. No se crea ninguna base local.
`DATABASE_URL` es obligatoria; `DB_LIVE` y `DB_*` no se utilizan.

1. Subir esta version al repositorio que despliega Railway.
2. Definir `DATABASE_URL` con la URL del Session pooler de Supabase (puerto 5432),
   `SECRET_KEY`, `DEBUG=False` y `TIME_ZONE=America/Lima` en el servicio correcto.
3. Desplegar los cambios. `railway.json` y `Procfile` usan `bash start.sh`.
4. Revisar los logs: primero se comprueba PostgreSQL; luego se aplican migraciones,
   se recopilan estaticos y se muestra el numero de empleados, sin revelar claves.
5. Si hay cero empleados, verificar primero el proyecto Supabase de destino.
   Si corresponde cargar la lista incluida, establecer `CARGAR_DATOS_INICIALES=True`
   y redesplegar. Despues volver a `False`. No importa datos de una base anterior.
   El comando usa get_or_create; no borra empleados existentes.

Diagnostico manual de solo lectura: `python manage.py verificar_bd`.
La carga manual es `python manage.py cargar_empleados` (comando Django, no el script
independiente del mismo nombre). No ejecutar cargas contra una base equivocada.
Una contrasena con `@` necesita `%40` dentro de la URL, codificado una sola vez.
Nunca publicar `DATABASE_URL` ni `SECRET_KEY` en capturas o repositorios.
