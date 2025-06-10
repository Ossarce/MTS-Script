# MTS Script

Este script analiza archivos Excel con registros de detenciones de línea de producción. Agrupa y resume las horas mensuales por tipo de detención: producción, fallas, mantenciones, micro paradas, entre otros.

## Instalación

1.- **Clona este repositorio** o copia el script en un directorio:
```bash
git clone https://github.com/Ossarce/MTS-Script.git
cd MTS-Script
```

2.- Crea y activa un entorno virtual (opcional pero recomendado):
```python
python3 -m venv venv
source venv/bin/activate
```

3.- Instala las dependencias necesarias:
```python
pip install pandas openpyxl
```

## Uso

1.- Asegúrate de que los archivos Excel que deseas analizar estén en el mismo directorio donde se encuentra el script ```app.py```

2.- Luego, simplemente ejecuta el script desde la terminal 
``` python 
python app.py
```

## License

[MIT](https://choosealicense.com/licenses/mit/)