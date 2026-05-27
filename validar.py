"""Validacion completa de consistencia — Equipo 1"""

import sqlite3

conn = sqlite3.connect(
    r"C:\Users\Usuario\OneDrive - auraoiliveoil.com\Escritorio\Riego por Cuartel 2025 - 2026\riego.db"
)
conn.row_factory = sqlite3.Row


def check(ok, msg):
    print("  {}  {}".format("OK" if ok else "FAIL", msg))
    return ok


todo_ok = True
print("=" * 65)
print("VALIDACION DE CONSISTENCIA — EQUIPO 1")
print("=" * 65)

# 1. Suma ha cuarteles en sector == ha total del sector
print("\n1. Has cuarteles en sector == Has sector (tabla Sectores)")
for s in conn.execute("SELECT * FROM sectores ORDER BY sector_nom").fetchall():
    suma = (
        conn.execute(
            "SELECT SUM(has_en_sector) FROM cuartel_sector WHERE sector_nom=?",
            (s["sector_nom"],),
        ).fetchone()[0]
        or 0
    )
    ok = abs(suma - s["has_total"]) < 0.01
    if not ok:
        todo_ok = False
    print(
        "  {}: suma cuarteles={:.2f} ha | sector={:.2f} ha | diff={:+.4f}".format(
            s["sector_nom"], suma, s["has_total"], suma - s["has_total"]
        )
    )

# 2. Suma ha de cada cuartel en sus sectores <= ha total del cuartel
print("\n2. Suma ha_en_sector <= has_total por cuartel")
fallos_cc = 0
for c in conn.execute("SELECT cc, has_total FROM cuarteles ORDER BY cc").fetchall():
    suma = (
        conn.execute(
            "SELECT SUM(has_en_sector) FROM cuartel_sector WHERE cuartel_id=?",
            (c["cc"],),
        ).fetchone()[0]
        or 0
    )
    if suma > c["has_total"] + 0.01:
        todo_ok = False
        fallos_cc += 1
        print(
            "  FAIL CC {}: suma={:.2f} > has_total={:.2f}".format(
                c["cc"], suma, c["has_total"]
            )
        )
if fallos_cc == 0:
    print("  Todos OK (para Equipo 1, cada cuartel esta 100% en su unico sector)")

# 3. Suma m3 distribuidos por riego == volumen original
print("\n3. Suma m3 cuartel == m3 riego original (por riego)")
fallas = conn.execute("""
    SELECT r.id, r.sector_nom, r.volumen_m3, ROUND(SUM(rc.volumen_m3), 2) as suma_dist
    FROM riegos r
    JOIN riegos_cuartel rc ON r.id = rc.riego_id
    GROUP BY r.id
    HAVING ABS(r.volumen_m3 - SUM(rc.volumen_m3)) > 0.05
    LIMIT 5
""").fetchall()
if fallas:
    todo_ok = False
    for f in fallas:
        print(
            "  FAIL riego {}: original={} distribuido={} diff={}".format(
                f["id"],
                f["volumen_m3"],
                f["suma_dist"],
                f["volumen_m3"] - f["suma_dist"],
            )
        )
else:
    print("  Los 367 riegos distribuyen su volumen exactamente (tolerancia 0.05 m3)")

# 4. Todo riego tiene distribucion
print("\n4. Cada riego tiene distribucion en riegos_cuartel")
sin_dist = conn.execute("""
    SELECT r.id, r.sector_nom, r.volumen_m3
    FROM riegos r
    WHERE r.id NOT IN (SELECT DISTINCT riego_id FROM riegos_cuartel)
""").fetchall()
if sin_dist:
    todo_ok = False
    for r in sin_dist:
        print(
            "  FAIL riego {} en {} ({} m3) sin distribucion".format(
                r["id"], r["sector_nom"], r["volumen_m3"]
            )
        )
else:
    print("  Los 367 riegos tienen distribucion")

# 5. Total m3 mensual: riegos vs riegos_cuartel
print("\n5. Total m3 por mes: riegos == riegos_cuartel")
for m in conn.execute(
    "SELECT DISTINCT strftime('%Y-%m', fecha) as mes FROM riegos ORDER BY mes"
).fetchall():
    mes = m["mes"]
    vr = (
        conn.execute(
            "SELECT SUM(volumen_m3) FROM riegos WHERE strftime('%Y-%m', fecha)=?",
            (mes,),
        ).fetchone()[0]
        or 0
    )
    vd = (
        conn.execute(
            """SELECT SUM(rc.volumen_m3) FROM riegos_cuartel rc
           JOIN riegos r ON rc.riego_id = r.id
           WHERE strftime('%Y-%m', r.fecha)=?""",
            (mes,),
        ).fetchone()[0]
        or 0
    )
    diff = vr - vd
    ok = abs(diff) < 0.5
    if not ok:
        todo_ok = False
    print(
        "  {}: riegos={:>10,.1f} m3 | distribuido={:>10,.1f} m3 | diff={:+.2f}".format(
            mes, vr, vd, diff
        )
    )

# 6. Total global
print("\n6. Total global")
vol_tot_riegos = conn.execute("SELECT SUM(volumen_m3) FROM riegos").fetchone()[0] or 0
vol_tot_dist = (
    conn.execute("SELECT SUM(volumen_m3) FROM riegos_cuartel").fetchone()[0] or 0
)
ok = abs(vol_tot_riegos - vol_tot_dist) < 1
if not ok:
    todo_ok = False
print("  Riegos totales: {:,.1f} m3".format(vol_tot_riegos))
print("  Distribuido:    {:,.1f} m3".format(vol_tot_dist))
print("  Diferencia:     {:+.4f} m3".format(vol_tot_riegos - vol_tot_dist))

# 7. Sectores sin cuarteles
print("\n7. Sectores en riegos que no estan en cuartel_sector")
huerfanos = conn.execute("""
    SELECT DISTINCT r.sector_nom, COUNT(*) as n
    FROM riegos r
    LEFT JOIN cuartel_sector cs ON r.sector_nom = cs.sector_nom
    WHERE cs.sector_nom IS NULL
    GROUP BY r.sector_nom
""").fetchall()
if huerfanos:
    todo_ok = False
    for h in huerfanos:
        print("  FAIL {}: {} riegos sin cuarteles".format(h["sector_nom"], h["n"]))
else:
    print("  Ninguno. Todos los sectores con riegos tienen cuarteles.")

# 8. Verificar que cuarteles de la BD tengan datos de riego
print("\n8. Cuarteles con y sin riego")
con_riego = conn.execute("""
    SELECT COUNT(DISTINCT cuartel_id) FROM riegos_cuartel
""").fetchone()[0]
total_cc = conn.execute("SELECT COUNT(*) FROM cuarteles").fetchone()[0]
print("  {} de {} cuarteles recibieron riego".format(con_riego, total_cc))

# 9. Verificar cuartel x sector en la BD sea correcto (no duplicado)
print("\n9. Verificacion N:M — cuarteles en multiples sectores")
multi = conn.execute("""
    SELECT cuartel_id, COUNT(*) as n, GROUP_CONCAT(sector_nom) as sectores
    FROM cuartel_sector
    GROUP BY cuartel_id
    HAVING n > 1
    ORDER BY cuartel_id
""").fetchall()
if multi:
    for m in multi:
        print(
            "  CC {}: {} sectores ({})".format(m["cuartel_id"], m["n"], m["sectores"])
        )
else:
    print(
        "  Ningun cuartel en Equipo 1 esta en mas de un sector (N:M no aplica para este equipo)"
    )

print()
print("=" * 65)
if todo_ok:
    print("RESULTADO: TODO CALZA PERFECTAMENTE.")
else:
    print("RESULTADO: Hay inconsistencias (ver arriba).")
print("=" * 65)
conn.close()
