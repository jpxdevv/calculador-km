# Calculador de Kilometraje y Rendimiento

Programa de escritorio (Windows y Mac) que toma el Excel de cargas de diésel y calcula automáticamente:

- **MES** de la fecha
- **KM INICIAL**: km de la última carga anterior del mismo tracto con folio de liquidación > 0
- **KM FINAL**: km de la carga
- **KM RECORRIDOS**: final − inicial
- **RENDIMIENTO** (km/L), **REND. META** (de la hoja catálogo; 2.5 por defecto), **LTS META** y **FALTANTE** en litros y en pesos

Las cargas del mismo tracto en la misma fecha se juntan en una sola línea: se suman los litros y los tickets van en una sola celda.

## Descargar

Ve a **[Releases](../../releases/latest)** y baja el `.zip` de tu sistema:

| Sistema | Archivo | Cómo abrir |
|---|---|---|
| Windows | `CalculadorKM-Windows.zip` | Descomprime y abre `CalculadorKM.exe`. Si Windows avisa "protegió su PC", da clic en **Más información → Ejecutar de todas formas**. |
| Mac | `CalculadorKM-Mac.zip` | Descomprime, mueve `CalculadorKM.app` a Aplicaciones. La primera vez: **clic derecho → Abrir → Abrir**. |

## Uso

1. Abre el programa y da clic en **Seleccionar Excel…**
2. Espera a que procese y revisa el resumen en pantalla.
3. Guarda el resultado. El Excel generado trae las hojas:
   - **Resumen por folio**: km, litros, rendimiento, meta y faltante por folio de liquidación
   - **Detalle por día**: una línea por tracto y fecha
   - **Revisar**: líneas con 0 km o km menor al anterior
   - **Sin km (remolque)**: cargas sin odómetro, si las hay

### Formato del Excel de entrada

- La **primera hoja** es la base de datos, con las columnas `TICKET`, `FECHA`, `TRACTO`, `LTS`, `FOLIO LIQ.` y `ODOMETRO` (o `kms`). Opcionales: `TOTAL`, `OPERADOR`, `NOM OPERADOR`, etc.
- Una hoja con `TRACTO` y `REND META SB` se usa como catálogo de metas.

## Para desarrolladores

```bash
python -m venv venv
./venv/bin/pip install -r requirements.txt
./venv/bin/python app.py                    # interfaz
./venv/bin/python procesador.py archivo.xlsx   # desde terminal
```

Para publicar una versión nueva: `git tag v1.0.1 && git push --tags`. GitHub Actions construye las apps de Windows y Mac y las sube a Releases.
