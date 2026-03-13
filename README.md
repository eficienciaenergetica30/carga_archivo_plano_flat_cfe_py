# Proyecto Flask

Aplicación web desarrollada con **Python** y **Flask**.

---

# 📋 Requisitos Previos

Antes de ejecutar el proyecto asegúrate de tener instalado:

* **Python 3.8 o superior**
* **pip** (gestor de paquetes de Python)
* **Git** (opcional, para clonar el repositorio)

---

# 🚀 Instalación y Configuración

## 1. Clonar el repositorio (si aplica)

```bash
git clone <url-del-repositorio>
cd nombre-del-proyecto
```

---

## 2. Crear y activar el entorno virtual

### Crear el entorno virtual

```bash
python -m venv venv
```

### Activar el entorno virtual

**Windows**

```bash
venv\Scripts\activate
```

**Linux / Mac**

```bash
source venv/bin/activate
```

---

## 3. Instalar dependencias

Con el entorno virtual activado:

```bash
pip install -r requirements.txt
```

Si instalas nuevas dependencias:

```bash
pip install <nombre-del-paquete>
pip freeze > requirements.txt
```

---

# ▶️ Ejecutar la aplicación

Con el entorno virtual activado:

```bash
python run.py
```

---

# 🧰 Comandos útiles

### Desactivar el entorno virtual

```bash
deactivate
```

### Actualizar requirements.txt

```bash
pip freeze > requirements.txt
```

### Instalar dependencias del proyecto

```bash
pip install -r requirements.txt
```

### Ver paquetes instalados

```bash
pip list
```
