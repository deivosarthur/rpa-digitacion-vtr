# 🤖 RPA Digitación VTR

Automatización de procesos de digitación de órdenes de trabajo (OT) utilizando Python, Selenium y SQL Server.

## 🚀 Funcionalidades

* Automatización de login en sistema EPS VTR
* Procesamiento masivo de órdenes de trabajo
* Integración con base de datos SQL Server
* Control de estados mediante flag (PENDIENTE, OK, ERROR, etc.)
* Ejecución en paralelo mediante múltiples bots (monitores)
* Registro de logs de ejecución

## 🧠 Arquitectura

SQL Server → Bot (Python + Selenium) → Sistema VTR → Actualización de estado en SQL

## 📊 Base de Datos

Tabla principal:

* proceso_digitacion_resumen

Campos relevantes:

* OT
* Estado_digitacion (usado como flag)
* id (Primary Key)

## ⚙️ Requisitos

* Python 3.x
* Google Chrome
* ChromeDriver compatible
* SQL Server
* Librerías:

  * selenium
  * pyodbc
  * openpyxl

Instalar dependencias:

```bash
pip install -r requirements.txt
```

## ▶️ Ejecución

Ejecutar bot individual:

```bash
python digitacion1.py
```

Ejecutar interfaz:

```bash
python Home2.py
```

## ⚠️ Seguridad

El archivo `keys.txt` contiene credenciales sensibles y NO debe subirse al repositorio.

## 📌 Notas

* Proyecto en evolución hacia arquitectura full SQL (sin uso de Excel)
* Preparado para escalabilidad y multi-bot
