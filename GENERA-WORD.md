# 📄 SISTEMA ACTUALIZADO - GENERA DOCUMENTOS WORD

## ✨ ¡NUEVO! Ahora Genera Archivos .docx (Word)

He modificado el sistema para que genere **documentos Word** (.docx) en lugar de PDFs.

### 🎯 Ventajas de Usar Word

✅ **Editable** - Puedes modificar el documento después
✅ **Sin programas externos** - No necesitas wkhtmltopdf ni nada
✅ **Compatible** - Se abre en Word, LibreOffice, Google Docs
✅ **Formato profesional** - Times New Roman, márgenes correctos
✅ **Más fácil** - Instalación super simple

---

## ⚡ INSTALACIÓN (2 minutos)

### Paso 1: Instalar dependencias

```powershell
pip install Flask==3.0.0 Flask-SQLAlchemy==3.1.1 pandas openpyxl==3.1.2 python-docx==1.1.0
```

### Paso 2: Ejecutar

```powershell
python app.py
```

### Paso 3: Abrir navegador

```
http://localhost:5000
```

¡Listo! 🎊

---

## 📋 ¿Qué cambió?

### ANTES (Versión PDF):
- ❌ Generaba PDFs (no editables)
- ❌ Necesitaba wkhtmltopdf o WeasyPrint
- ❌ Más complicado de instalar

### AHORA (Versión Word):
- ✅ Genera archivos .docx (editables)
- ✅ Solo necesita python-docx (simple)
- ✅ Instalación super fácil
- ✅ Puedes editar los documentos después

---

## 🎨 Formato del Documento Word

El sistema genera documentos con:

- **Fuente**: Times New Roman 12pt
- **Márgenes**: 2.5cm arriba, 2cm abajo, 2cm laterales
- **Formato**: Exactamente igual a la plantilla original
- **Encabezado**: Alineado a la derecha con folio y fecha
- **Cuerpo**: Justificado
- **Firma**: Centrada al final

---

## 💾 Cómo Funciona

### Módulo Manual:
1. Creas o editas una constancia
2. Haces clic en 📄
3. Se descarga un archivo `.docx`
4. Lo abres en Word y listo

### Módulo Masivo:
1. Importas Excel con múltiples constancias
2. Seleccionas las que quieres
3. Clic en "Generar Word Seleccionados"
4. Se descarga un ZIP con todos los `.docx`

---

## 🔧 Archivos Modificados

1. **requirements.txt**
   - weasyprint → **python-docx==1.1.0**

2. **app.py**
   - Agregada función `crear_documento_word()`
   - Actualizado para generar .docx
   - Mantiene toda la funcionalidad

3. **templates/**
   - Botones ahora dicen "Descargar Word"
   - Mensajes actualizados

4. **instalar.ps1**
   - Instala python-docx en lugar de weasyprint

---

## 📊 Comparación PDF vs Word

| Característica | PDF | Word (Actual) |
|----------------|-----|---------------|
| **Editable** | No ❌ | Sí ✅ |
| **Instalación** | Compleja | Simple ✅ |
| **Compatibilidad** | Universal ✅ | Universal ✅ |
| **Formato** | Fijo ✅ | Fijo ✅ |
| **Tamaño archivo** | Pequeño ✅ | Pequeño ✅ |
| **Se puede modificar después** | No ❌ | Sí ✅ |
| **Dependencias externas** | Sí (wkhtmltopdf) ❌ | No ✅ |

---

## ✅ Ventajas de Word sobre PDF

1. **Editable**: Puedes corregir errores sin regenerar
2. **Firmas digitales**: Más fácil agregar firmas después
3. **Personalización**: Cada documento se puede ajustar
4. **Compatible**: Word, Google Docs, LibreOffice lo abren
5. **Instalación**: Solo `pip install python-docx`

---

## 🎯 Casos de Uso

### Si necesitas editar después:
✅ **Word es mejor** - Puedes modificar el documento

### Si necesitas documento final inmutable:
⚠️ Genera Word, luego guarda como PDF desde Word

### Si necesitas ambos:
1. Descarga el .docx del sistema
2. Ábrelo en Word
3. Guarda como PDF
4. ¡Tienes ambos formatos!

---

## 🔄 ¿Y si prefiero PDF?

Si necesitas PDFs en lugar de Word:

**Opción 1: Convertir en Word**
1. Abre el .docx
2. Archivo → Guardar como → PDF

**Opción 2: Usar versión anterior**
Puedo darte la versión con WeasyPrint si lo prefieres

---

## 📝 Ejemplo de Uso

```powershell
# Instalar
pip install Flask==3.0.0 Flask-SQLAlchemy==3.1.1 pandas openpyxl==3.1.2 python-docx==1.1.0

# Ejecutar
python app.py

# Abrir navegador → localhost:5000
# Dashboard → Módulo Manual → Nueva Constancia
# Llenar formulario → Guardar
# Click en 📄 → Descarga constancia_001.docx
# Abrir en Word → ¡Listo!
```

---

## 🐛 Solución de Problemas

### Error: "No module named 'docx'"

```powershell
pip install python-docx==1.1.0
```

### Documento se descarga pero no abre

El archivo está bien, solo necesitas Word instalado:
- **Windows**: Microsoft Word
- **Mac**: Microsoft Word o Pages
- **Linux**: LibreOffice Writer
- **Cualquiera**: Google Docs (subir el archivo)

### Formato se ve diferente

El documento está diseñado para Times New Roman.
Si no la tienes, instala la fuente o se usará una similar.

---

## 💡 Tips

1. **Editar plantilla**: Modifica la función `crear_documento_word()` en app.py
2. **Cambiar fuente**: Busca `run.font.size = Pt(12)` y modifica
3. **Ajustar márgenes**: Busca `Inches(1)` y cambia valores
4. **Agregar logo**: Usa `doc.add_picture('logo.png')`

---

## 🎉 Resumen

**Instalación:**
```powershell
pip install Flask==3.0.0 Flask-SQLAlchemy==3.1.1 pandas openpyxl==3.1.2 python-docx==1.1.0
```

**Ejecutar:**
```powershell
python app.py
```

**Usar:**
1. Crear constancia
2. Click en 📄
3. Descargar .docx
4. Abrir en Word
5. ¡Listo!

---

**Formato de salida**: `.docx` (Microsoft Word)
**Editable**: ✅ Sí
**Programas externos necesarios**: ❌ Ninguno
**Tiempo de instalación**: ~2 minutos
**Dificultad**: ⭐ Muy fácil

¡Disfruta tu sistema con documentos Word editables! 📄✨
