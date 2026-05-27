# Riego por Cuartel Siracusa

Herramienta de auditoria de riego agricola. Distribuye los m3 regados por sector entre los cuarteles que lo componen, usando una relacion muchos-a-muchos (N:M) basada en hectareas proporcionales.

## Uso

```bash
python auditar_riegos.py 1    # Equipo 1
python auditar_riegos.py 2    # Equipo 2
python auditar_riegos.py 4    # Equipo 4
```

O doble clic en `Auditar.bat`.

## Requisitos

- Python 3.9+
- openpyxl
- pandas (opcional, para analisis avanzado)

```bash
pip install openpyxl pandas
```

## Archivos de entrada

1. `Cuartel x Sector.xlsx` — Tabla de relacion cuartel-sector con porcentajes
2. `Historial/` — Carpeta con planillas mensuales de riego exportadas del software

## Salida

`Auditoria_Riego_EquipoX.xlsx` con 5 hojas:
- **Sectores** — Datos de cada sector
- **Cuartel_x_Sector** — Relacion N:M con % y hectareas
- **Historial_Riegos** — Todos los riegos cargados
- **Resumen_Cuartel_x_Mes** — m3 por cuartel por mes
- **Resumen_Detallado** — m3 + m3/ha por cuartel por mes (columnas de eficiencia en negrita)
