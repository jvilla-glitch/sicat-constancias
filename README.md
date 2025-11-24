# 🏥 Sistema de Constancias Sanitarias - COEPRISS Sinaloa

Sistema web para la gestión de constancias sanitarias de la Comisión Estatal para la Protección Contra Riesgos Sanitarios de Sinaloa.

## 📋 Características

- ✅ **Gestión de Constancias Manuales**: Crear, editar, ver y eliminar constancias individuales
- ✅ **Importación Masiva**: Cargar múltiples constancias desde archivos Excel/CSV
- ✅ **Generación de Documentos**: Exportar constancias a formato Word (.docx)
- ✅ **Dashboard**: Visualización de estadísticas y métricas del sistema
- ✅ **Sistema de Autenticación**: Login seguro con usuarios predefinidos
- ✅ **Interfaz Responsive**: Compatible con dispositivos móviles y de escritorio

## 🛠️ Tecnologías

- **Backend**: Flask 3.0
- **Base de Datos**: SQLite (desarrollo) / PostgreSQL (producción)
- **ORM**: SQLAlchemy
- **Autenticación**: Flask-Login
- **Frontend**: Bootstrap 5, HTML5, CSS3, JavaScript
- **Documentos**: python-docx
- **Procesamiento de Datos**: pandas, openpyxl

## 📦 Instalación Local

### Requisitos
- Python 3.11+
- pip

### Pasos

1. **Clonar el repositorio**
```bash
git clone https://github.com/TU-USUARIO/sistema-constancias-coepriss.git
cd sistema-constancias-coepriss
```

2. **Crear entorno virtual**
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

3. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

4. **Ejecutar la aplicación**
```bash
python app.py
```

5. **Abrir en navegador**
```
http://127.0.0.1:5000
```

## 👥 Usuarios Predefinidos

El sistema incluye usuarios predeterminados:

- ALEJANDRA SEPULVEDA
- JOSUE VILLA
- BRENDA GONZALEZ
- ERIK CHAVEZ
- MARLEN ROCHA
- LILIA PEÑA

**Contraseña por defecto:** `cis2025`

## 📄 Estructura del Proyecto

```
sistema-constancias-coepriss/
├── app.py                    # Aplicación principal Flask
├── models.py                 # Modelos de base de datos
├── requirements.txt          # Dependencias Python
├── Procfile                  # Configuración para Render
├── templates/                # Plantillas HTML
│   ├── dashboard.html
│   ├── login.html
│   ├── plantilla_constancia.html
│   ├── manual/
│   │   ├── index.html
│   │   └── formulario.html
│   └── masivo/
│       └── index.html
└── static/                   # Archivos estáticos
    └── css/
        └── styles.css
```

## 🚀 Despliegue en Render

1. Crear cuenta en [Render.com](https://render.com)
2. Conectar repositorio de GitHub
3. Crear nuevo Web Service
4. Configurar variables de entorno:
   - `SECRET_KEY`: Clave secreta para Flask
5. Agregar PostgreSQL Database
6. Deploy automático

## 📊 Formato de Archivo Excel para Importación

El archivo debe contener las siguientes columnas:

| Columna | Descripción |
|---------|-------------|
| folio | Número de folio único |
| fecha | Fecha de emisión |
| vigencia | Fecha de vigencia |
| expediente | Número de expediente |
| nombre_propietario | Nombre del propietario |
| domicilio_propietario | Dirección del propietario |
| giro | Tipo de establecimiento |
| denominado | Nombre del establecimiento |
| ubicado | Dirección del establecimiento |
| entre_calles | Calles entre las que se ubica |
| colonia | Colonia |
| ciudad | Ciudad |
| codigo_postal | Código postal |
| nombre_comisionado | Nombre del comisionado que firma |

## 🔒 Seguridad

- Autenticación de usuarios requerida para todas las funcionalidades
- Contraseñas hasheadas con werkzeug.security
- Variables de entorno para información sensible
- Validación de folios únicos

## 📝 API Endpoints

### Constancias
- `GET /api/constancias` - Listar todas las constancias
- `GET /api/constancias/<id>` - Obtener constancia específica
- `POST /api/constancias` - Crear nueva constancia
- `PUT /api/constancias/<id>` - Actualizar constancia
- `DELETE /api/constancias/<id>` - Eliminar constancia

### Documentos
- `GET /constancia/<id>/pdf` - Generar documento Word individual
- `POST /generar-masivo` - Generar documento Word con múltiples constancias
- `POST /masivo/generar-pdfs` - Generar documentos masivos (alias)

### Importación
- `POST /masivo/importar` - Importar constancias desde Excel/CSV

### Estadísticas
- `GET /api/estadisticas` - Obtener estadísticas del dashboard

## 🤝 Contribuir

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto fue desarrollado para la Comisión Estatal para la Protección Contra Riesgos Sanitarios de Sinaloa (COEPRISS).

## 📧 Contacto

COEPRISS Sinaloa - Sistema de Gestión de Constancias Sanitarias

---

⭐ Si este proyecto te fue útil, ¡dale una estrella!
