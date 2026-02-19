# Horas Extraordinarias - Registro en Parte Mensual

Aplicación de escritorio en Python para registrar horas extraordinarias en el parte mensual oficial de Excel, preservando **absolutamente todo** el formato del archivo original.

## Características

- Preserva imágenes incrustadas, celdas combinadas, anchos de columna, alturas de fila, filas ocultas, estilos, bordes y fuentes
- Solo modifica los **valores** en las celdas indicadas
- Interfaz gráfica sencilla con Tkinter
- Guarda la ruta del archivo para no pedirla cada vez
- Soporte para fila con fórmula `=DAY(Bx)` en columna A

## Requisitos

- Python 3.8 o superior
- openpyxl >= 3.1.0

## Instalación

```bash
# Clonar el repositorio
git clone https://github.com/pikinaita/horas-extras-excel.git
cd horas-extras-excel

# Instalar dependencias
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

### Primera ejecución

La primera vez pedirá seleccionar el archivo Excel. La ruta se guarda en `config.json` para futuras ejecuciones.

### Formulario

1. **Fecha**: Selecciona el día y mes de las horas extra
2. **Turno**: Mañana (columna C) / Tarde (columna D) / Noche (columna E)
3. **A quién cubre**: Sustitución Operador / Sustitución Jefe de Turno

Pulsa **"Registrar y guardar"** para guardar en el archivo original, o **"Registrar en copia..."** para guardar una copia.

## Estructura del proyecto

```
horas-extras-excel/
├── main.py           # Aplicación principal
├── requirements.txt  # Dependencias
├── .gitignore        # Excluye config.json y archivos xlsx
└── README.md         # Este archivo
```

## Notas

- El archivo `config.json` (generado automáticamente) y los archivos `.xlsx` están excluidos por `.gitignore`
- Versión: 2.1.0
