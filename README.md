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
- *Base de Datos*: PostgreSQL (Producción) / SQLite (Desarrollo)
- *Estáticos*: WhiteNoise
- *Servidor*: Gunicorn
- *Utilidades*: OpenPyXL (Excel), qrcode, python-dotenv

---

## Configuración del Entorno

Dado que este proyecto maneja información sensible, **es necesario crear un archivo `.env`** en la raíz del proyecto.

Crea un archivo llamado `.env` y agrega las siguientes variables:

```env
# --- Configuración General ---
# Clave secreta de Django (Generar una nueva para producción)
SECRET_KEY=tu_clave_secreta_super_segura_aqui

# Modo Debug (True para desarrollo, False para producción)
DEBUG=True

# --- Base de Datos ---
# DB_LIVE controla si se usa PostgreSQL (True) o SQLite (False)
DB_LIVE=False

# URL de conexión a base de datos (Prioridad sobre variables individuales)
# Ejemplo: postgresql://usuario:password@localhost:5432/mi_db
DATABASE_URL=

# Variables individuales (Usadas si DB_LIVE=True y no hay DATABASE_URL)
DB_NAME=nombre_db
DB_USER=usuario_db
DB_PASSWORD=password_db
DB_HOST=localhost
DB_PORT=5432

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

1. Configura las variables de entorno en el panel de tu proveedor (`DATABASE_URL`, `SECRET_KEY`, `DB_LIVE=True`, etc.).
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
