"""
Dashboard TGM (Transvisa E.I.R.L.) - Baja eficacia del cumplimiento de mantenimiento preventivo
Versión 4: identidad corporativa azul marino + rojo, tipografía negra (modo claro y oscuro) y Resumen Ejecutivo
con los 6 indicadores clave: brecha técnica, impacto económico, inversión total, ahorro anual, VAN y TIR.
Todos los indicadores se calculan con los datos y las fórmulas del Excel del proyecto (T3_Equipo_06_3.xlsx).
Ejecutar con:  streamlit run dashboard_tgm_v4.py
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="TGM | Mantenimiento preventivo", page_icon="🚛", layout="wide")

# ======================= DATOS (del Excel del equipo) =======================
# Quincenas 2025: [periodo, OT programadas, OT ejecutadas, costo tracto S/, costo tolva/cisterna S/]
# Fuente: hoja "Análisis Cuantitativo" (filas 216-241) = hoja "Ficha Técnica" (Tabla 9). Precisión original del Excel.
SEMANAL = [
    ['S01-02', 15, 10, 14747.8, 19894.03],
    ['S03-04', 8, 5, 3347.84, 8468.14],
    ['S04-05', 11, 6, 10635.51, 4471.74],
    ['S06-07', 11, 7, 18621.93, 1407.28],
    ['S08-09', 12, 8, 5261.22, 18017.83],
    ['S10-11', 14, 9, 11229.77, 29989.02],
    ['S12-13', 14, 10, 13467.644, 17211.7],
    ['S14-15', 8, 8, 18270.0061, 3658.39],
    ['S16-17', 2, 2, 5066.83, 192.66],
    ['S18-19', 5, 4, 3133.43, 5416.55],
    ['S20-21', 7, 4, 8958.43, 2817.64],
    ['S22-23', 7, 4, 695.32, 134.08],
    ['S24-25', 4, 1, 1888.82, 560.2],
    ['S26-27', 18, 12, 12858.83, 7151.18],
    ['S28-29', 17, 12, 14418.105, 9635.9],
    ['S30-31', 15, 11, 7220.69, 8395.15],
    ['S32-33', 10, 8, 4846.09, 56401.09],
    ['S34-35', 7, 7, 4219.55, 2183.42],
    ['S36-37', 5, 3, 18895.25, 6572.13],
    ['S38-39', 14, 10, 21726.3466, 4019.43],
    ['S40-41', 11, 11, 19618.58, 10931.71],
    ['S42-43', 7, 4, 20759.2, 1042.59],
    ['S44-45', 10, 8, 5648.05, 7963.19],
    ['S46-47', 10, 8, 11969.7752, 2604.6966],
    ['S48-49', 9, 9, 15490.2966, 5380.8826],
    ['S50-51', 11, 9, 0, 0]
]
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
ABREV = [m[:3].upper() for m in MESES]
# AS-IS mensual 2025. Fuente: hoja "Ficha Técnica" (Tablas 10-14, filas 53-64); los % están multiplicados por 100.
ASIS_MENSUAL = {
    'capacitacion': [66.666667, 58.333333, 75.0, 50.0, 41.666667, 66.666667, 75.0, 58.333333, 58.333333, 50.0, 75.0, 66.666667],
    'tiempo_muerto': [11.666667, 15, 10.833333, 3.333333, 5.833333, 7, 2, 8.333333, 3.666667, 8.666667, 2.083333, 1.333333],
    'disponibilidad': [88.888889, 83.72093, 92.307692, 78.26087, 86.956522, 82.379863, 78.431373, 74.226804, 89.440994, 86.330935, 85.106383, 93.264249],
    'reprocesos': [22.222222, 33.333333, 10.0, 14.285714, 12.5, 22.222222, 28.571429, 40.0, 10.0, 25.0, 11.111111, 9.090909],
    'setup': [520, 580, 470, 180, 310, 390, 100, 450, 200, 470, 110, 80]
}
# Asistentes AS-IS por mes (Ficha Técnica, col. P) y total de colaboradores (col. O)
CAP_ASIST_ASIS = [8, 7, 9, 6, 5, 8, 9, 7, 7, 6, 9, 8]
TOTAL_COLAB = 12

META_OT = 90.0                 # Valor de referencia VR (Planteamiento!D13)
FACTURACION_ANUAL = 2_585_000  # Ingresos anuales de TGM (Planteamiento!D21)

# ---------- Paleta corporativa: azul marino + rojo ----------
NAVY, NAVY2, AZUL, AZUL_CLARO = "#0B1F4B", "#173A7A", "#2457A6", "#6B93D6"
ROJO, ROJO_OSC = "#C8102E", "#8E0C22"
GRIS = "#5B6475"
NEGRO = "#000000"

MOTIVOS = [
    dict(id="C1", icono="🎓", color=NAVY, nombre="Capacitación", col="capacitacion",
         ind="% de asistencia a capacitaciones", unidad="%", as_is=61.8, to_be=91.8, estandar=90.0, mejor="mayor",
         metodo="ADDIE", metodo_det="ADDIE (Análisis, Diseño, Desarrollo, Implementación, Evaluación) · Luo et al. (2024)",
         formula="C(asist) = (N° de colaboradores que asistieron / Total de colaboradores del área) × 100",
         causa="Ausencia de un programa de capacitación continua en mantenimiento preventivo y autónomo."),
    dict(id="C2", icono="⏱️", color=AZUL, nombre="Tiempos muertos", col="tiempo_muerto",
         ind="Tiempo muerto mensual", unidad="horas", as_is=6.65, to_be=3.18, estandar=3.42, mejor="menor",
         metodo="ISOS", metodo_det="Intelligent Scheduling Optimization System (programación inteligente de recursos)",
         formula="Horas muertas por mes = Σ horas sin actividad productiva durante el servicio",
         causa="Ausencia de un procedimiento para asegurar oportunamente los recursos requeridos para el mantenimiento."),
    dict(id="C3", icono="🛠️", color=AZUL_CLARO, nombre="Disponibilidad", col="disponibilidad",
         ind="Disponibilidad de equipos", unidad="%", as_is=84.9, to_be=95.3, estandar=95.0, mejor="mayor",
         metodo="TPM", metodo_det="TPM (Mantenimiento Productivo Total)",
         formula="Disponibilidad = MTBF / (MTBF + MTTR) × 100",
         causa="Averías recurrentes por mantenimiento correctivo no programado que reducen las ventanas de tiempo para el preventivo."),
    dict(id="C4", icono="🔁", color=ROJO, nombre="Reprocesos", col="reprocesos",
         ind="% de equipos con segundo mantenimiento", unidad="%", as_is=19.9, to_be=3.4, estandar=5.0, mejor="menor",
         metodo="DMAIC", metodo_det="DMAIC (Definir, Medir, Analizar, Mejorar, Controlar) · Six Sigma",
         formula="Reprocesos = (N° equipos con segundo mantenimiento / Total de equipos atendidos) × 100",
         causa="Ausencia de un protocolo de verificación de conformidad al finalizar el mantenimiento preventivo."),
    dict(id="C5", icono="⚙️", color=ROJO_OSC, nombre="Setup", col="setup",
         ind="Tiempo de setup mensual", unidad="minutos", as_is=321.7, to_be=166.1, estandar=140.4, mejor="menor",
         metodo="SMED", metodo_det="SMED (Cambio rápido de herramientas / preparación)",
         formula="Setup mensual = Σ minutos de preparación previa al servicio de mantenimiento",
         causa="Ausencia de un procedimiento estandarizado para la preparación previa del servicio de mantenimiento."),
]

# ======================= MODELO FINANCIERO (réplica de la hoja "Flujo Caja") =======================
# Inversión inicial: [concepto, propuesta, cantidad, costo unitario S/, clase de activo]  -> Flujo Caja!B57:G86
INVERSION = [
    ['Asesoría externa en diseño instruccional (2 semanas)', 'C1 · ADDIE', 1, 2400, 'INT-20'],
    ['Grabación con celular y edición de micro-videos de 4 minutos', 'C1 · ADDIE', 12, 120, 'INT-20'],
    ['Proyector y ecran para el aula del taller', 'C1 · ADDIE', 1, 2200, 'AF-10'],
    ['Componentes de descarte reacondicionados para las prácticas', 'C1 · ADDIE', 1, 1500, 'AF-10'],
    ['Material impreso, checklists, certificados y afiches', 'C1 · ADDIE', 1, 900, 'INT-20'],
    ['Configuración de reglas de programación en el ERP (horas del proveedor)', 'C2 · ISOS', 90, 95, 'INT-20'],
    ['Tablero semáforo "OT lista" en Power BI', 'C2 · ISOS', 1, 2200, 'INT-20'],
    ['Tablets de uso rudo para técnicos', 'C2 · ISOS', 3, 1500, 'AF-25'],
    ['Capacitación al programador y almacenero', 'C2 · ISOS', 1, 800, 'INT-20'],
    ['Restauración del deterioro de equipos clase A (repuestos y MO)', 'C3 · TPM', 4, 1800, 'INT-20'],
    ['Servicio externo de termografía y vibraciones – línea base', 'C3 · TPM', 4, 450, 'INT-20'],
    ['Análisis de aceite en laboratorio – línea base', 'C3 · TPM', 8, 180, 'INT-20'],
    ['Instrumentos básicos de inspección (pirómetro, estetoscopio, medidores)', 'C3 · TPM', 1, 1400, 'AF-10'],
    ['Capacitación en el pilar de Mantenimiento Planificado', 'C3 · TPM', 1, 1800, 'INT-20'],
    ['Hoja de vida de equipos en el ERP', 'C3 · TPM', 1, 1200, 'INT-20'],
    ['Tableros de control visual y tarjetas TPM', 'C3 · TPM', 1, 600, 'INT-20'],
    ['Stock inicial de repuestos críticos de equipos clase A', 'C3 · TPM', 1, 4000, 'KT'],
    ['Torquímetros calibrados', 'C4 · DMAIC', 2, 650, 'AF-10'],
    ['Calibración externa anual', 'C4 · DMAIC', 1, 800, 'INT-20'],
    ['Tablets para el checklist digital', 'C4 · DMAIC', 2, 1200, 'AF-25'],
    ['Configuración del checklist y bloqueo en el ERP', 'C4 · DMAIC', 1, 1500, 'INT-20'],
    ['Equipo para prueba de fugas (manómetro y detector UV)', 'C4 · DMAIC', 1, 1100, 'AF-10'],
    ['Capacitación en DMAIC y en el estándar de conformidad', 'C4 · DMAIC', 1, 2000, 'INT-20'],
    ['Impresión de instructivos y fotos patrón', 'C4 · DMAIC', 1, 500, 'INT-20'],
    ['Carro con tablero de sombras', 'C5 · SMED', 1, 1500, 'AF-10'],
    ['Organizadores para kits de preparación', 'C5 · SMED', 1, 400, 'AF-10'],
    ['Gata hidráulica de accionamiento rápido', 'C5 · SMED', 1, 1800, 'AF-10'],
    ['Kits de bloqueo por bahía', 'C5 · SMED', 2, 175, 'AF-10'],
    ['Capacitación en SMED y 5S', 'C5 · SMED', 1, 800, 'INT-20'],
    ['Señalización de bahías', 'C5 · SMED', 1, 400, 'INT-20']
]
# Costos anuales de sostenimiento: [concepto, propuesta, cantidad/año, costo unitario S/, se paga desde el año] -> Flujo Caja!B255:G263
SOSTENIMIENTO = [
    ['Bonos anuales por asistencia ≥ 90 % y post-test aprobado', 'C1 · ADDIE', 12, 100, 1],
    ['Horas de capacitación dentro de la jornada', 'C1 · ADDIE', 144, 8.33, 1],
    ['Horas del jefe de mantenimiento como instructor interno', 'C1 · ADDIE', 24, 12.5, 1],
    ['Soporte mensual del ISOS en el ERP', 'C2 · ISOS', 12, 250, 1],
    ['Licencias Power BI Pro (2 usuarios × 12 meses)', 'C2 · ISOS', 24, 52, 1],
    ['Servicio externo de termografía y vibraciones (semestral)', 'C3 · TPM', 8, 350, 1],
    ['Análisis de aceite periódico de equipos clase A', 'C3 · TPM', 16, 180, 1],
    ['Calibración externa de torquímetros', 'C4 · DMAIC', 1, 800, 2],
    ['Reposición de consumibles de kits SMED y señalización', 'C5 · SMED', 1, 400, 1]
]
PROG_MES = [34, 23, 28, 10, 12, 29, 32, 17, 19, 18, 20, 20]   # OT preventivas programadas 2025 (Base datos - Demanda!F122:F133)
TM_H = ASIS_MENSUAL["tiempo_muerto"]                           # horas muertas por mes (Ficha Técnica col. Z)
EQ2 = [2, 2, 1, 1, 1, 2, 2, 2, 1, 2, 1, 1]                     # equipos con 2.º mantenimiento por mes (Ficha Técnica col. AT)
SETUP_H = [round(v / 60, 6) for v in ASIS_MENSUAL["setup"]]    # horas de setup por mes (Ficha Técnica col. BE = minutos / 60)
FIN = dict(ir=0.295, horizonte=10, crec_ing=0.02, crec_cost=0.03, adopcion=0.85, imprevistos=0.05,
           hh_tec=8.33, hh_sup=12.5, tecnicos=2, supervisores=1, margen=0.20, costo_fallas=99000.0,
           prestamo=0.40, tea=0.165, plazo=5,
           rf=0.0484, beta_u=0.85, prima=0.0423, riesgo_pais=0.0108, prima_tam=0.03,
           dep_af10=0.10, dep_af25=0.25, dep_int=0.20,
           tm=(79.8, 38.2), fallas=(78, 29), eq2=(18, 3.4), setup=(64.3, 33.2))   # (AS-IS, TO-BE) anuales: Flujo Caja!C128:D128, C148:D148, C169:D169, C189:D189

# Valores que calcula el propio Excel (referencia; la tabla de validación ya no se muestra en el dashboard)
REF_EXCEL = dict(brecha=0.17307961646196968, impacto=507515.9427, pct_ingresos=0.19633111903288203, inv_total=61719.0,
                 ahorro=102790.86906590508, van=320909.10829136206, tir=1.1306078965338897, cppc=0.15123231,
                 cok=0.14205385, bc=6.199518921099857)


def _irr(fl, lo=-0.99, hi=100.0):
    """TIR por bisección (equivale a =IRR de Excel)."""
    f = lambda r: sum(c / (1 + r) ** t for t, c in enumerate(fl))
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(300):
        mid = (lo + hi) / 2
        if f(lo) * f(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def _meses_txt(anios):
    m = round(anios * 12)
    a, r = divmod(m, 12)
    p = []
    if a:
        p.append(f"{a} año" + ("" if a == 1 else "s"))
    if r:
        p.append(f"{r} mes" + ("" if r == 1 else "es"))
    return " y ".join(p) if p else "0 meses"


def modelo_financiero(vip, vr, costo_mp_total, mp_ejecutados, P=FIN):
    """Reproduce paso a paso la hoja 'Flujo Caja' del Excel: inversión, ahorro, flujo de caja, CPPC, VAN y TIR.
    vip = cumplimiento OT AS-IS (fracción) · vr = valor de referencia (fracción) · el costo del MP y los MP ejecutados salen de las quincenas."""
    r = {}
    # 1) Inversión total = Σ(cantidad × costo unitario) + imprevistos
    sub = sum(c * u for _, _, c, u, _ in INVERSION)
    r["inv_sub"], r["imprev"] = sub, sub * P["imprevistos"]
    r["inv_total"] = inv = sub * (1 + P["imprevistos"])
    # 2) Ahorro anual = Σ (costo AS-IS − costo TO-BE) de las 5 causas
    costo_ot = costo_mp_total / mp_ejecutados
    hh_cuad = P["tecnicos"] * P["hh_tec"]
    hh_eq = hh_cuad + P["supervisores"] * P["hh_sup"]
    asis = {"C1": 0.0, "C2": sum(TM_H) * hh_cuad, "C3": P["costo_fallas"], "C4": sum(EQ2) * costo_ot, "C5": sum(SETUP_H) * hh_eq}
    ratio = {"C1": 1.0, "C2": P["tm"][1] / P["tm"][0], "C3": P["fallas"][1] / P["fallas"][0],
             "C4": P["eq2"][1] / P["eq2"][0], "C5": P["setup"][1] / P["setup"][0]}
    tobe = {k: asis[k] * ratio[k] for k in asis}
    r["causas"] = {k: dict(asis=asis[k], tobe=tobe[k], ahorro=asis[k] - tobe[k]) for k in asis}
    r["costo_asis"], r["costo_tobe"] = sum(asis.values()), sum(tobe.values())
    r["ahorro"] = ahorro = r["costo_asis"] - r["costo_tobe"]
    # 3) OT adicionales al pasar de VIP a VR
    precio = costo_ot * (1 + P["margen"])
    ot_add = sum(PROG_MES) * (vr - vip)
    ing_ot, cost_ot = ot_add * precio, ot_add * costo_ot
    r["ing_ot"], r["cost_ot"], r["margen_ot"] = ing_ot, cost_ot, ing_ot - cost_ot
    # 4) Financiamiento (método francés) y costo de capital (CAPM -> CPPC)
    deuda, equity = inv * P["prestamo"], inv * (1 - P["prestamo"])
    cuota = deuda * P["tea"] / (1 - (1 + P["tea"]) ** -P["plazo"])
    saldo, interes, princ = deuda, [], []
    for t in range(1, P["horizonte"] + 1):
        if t <= P["plazo"]:
            i = saldo * P["tea"]
            a = cuota - i
            saldo -= a
        else:
            i = a = 0.0
        interes.append(i)
        princ.append(a)
    beta_l = P["beta_u"] * (1 + (1 - P["ir"]) * deuda / equity)
    cok = P["rf"] + beta_l * P["prima"] + P["riesgo_pais"] + P["prima_tam"]
    cppc = P["prestamo"] * P["tea"] + (1 - P["prestamo"]) * cok
    r.update(cok=cok, cppc=cppc, beta_l=beta_l, deuda=deuda, equity=equity, cuota=cuota)
    # 5) Depreciación, amortización, reinversión y valor residual
    af10 = sum(c * u for _, _, c, u, k in INVERSION if k == "AF-10")
    af25 = sum(c * u for _, _, c, u, k in INVERSION if k == "AF-25")
    intg = sum(c * u for _, _, c, u, k in INVERSION if k == "INT-20") + r["imprev"]
    kt = sum(c * u for _, _, c, u, k in INVERSION if k == "KT")
    H = P["horizonte"]
    vida25 = round(1 / P["dep_af25"])
    dep, acum10, acumi = [], 0.0, 0.0
    for t in range(1, H + 1):
        d10 = max(0.0, min(af10 * P["dep_af10"], af10 - acum10))
        acum10 += d10
        d25 = af25 * P["dep_af25"]
        di = max(0.0, min(intg * P["dep_int"], intg - acumi))
        acumi += di
        dep.append(d10 + d25 + di)
    reinv = [af25 if (t % vida25 == 0 and t < H) else 0.0 for t in range(1, H + 1)]
    residual = (af10 - acum10) + af25 * (1 - P["dep_af25"] * (H % vida25)) + kt
    # 6) Flujo de caja anual (años 1..10)
    sost_t = lambda t: sum(c * u for _, _, c, u, desde in SOSTENIMIENTO if desde <= t)
    flujo, filas = [-inv], []
    for t in range(1, H + 1):
        adop = P["adopcion"] if t == 1 else 1.0
        ing = (ing_ot + ahorro) * (1 + P["crec_ing"]) ** (t - 1) * adop
        gas = cost_ot * (1 + P["crec_cost"]) ** (t - 1) * adop + sost_t(t) * (1 + P["crec_cost"]) ** (t - 1)
        ub = ing - gas - dep[t - 1] - interes[t - 1]
        imp = ub * P["ir"]
        fc = ub + dep[t - 1] - princ[t - 1] - imp - reinv[t - 1] + (residual if t == H else 0.0)
        flujo.append(fc)
        filas.append(dict(anio=t, ingresos=ing, gastos=gas, dep=dep[t - 1], interes=interes[t - 1], utilidad=ub, impuesto=imp, flujo=fc))
    r["flujo"], r["filas"] = flujo, filas
    # 7) VAN, TIR, B/C y recuperación de la inversión
    vp = [flujo[t] / (1 + cppc) ** t for t in range(H + 1)]
    r["vp"], r["van"], r["tir"] = vp, sum(vp), _irr(flujo)
    r["bc"] = (r["van"] + inv) / inv
    ac, acd, rec_s, rec_d, s, sd = [], [], None, None, 0.0, 0.0
    for t in range(H + 1):
        s_prev, sd_prev = s, sd
        s += flujo[t]
        sd += vp[t]
        ac.append(s)
        acd.append(sd)
        if t >= 1 and rec_s is None and s_prev < 0 <= s:
            rec_s = (t - 1) + (-s_prev) / flujo[t]
        if t >= 1 and rec_d is None and sd_prev < 0 <= sd:
            rec_d = (t - 1) + (-sd_prev) / vp[t]
    r["acum"], r["acum_desc"], r["rec_simple"], r["rec_desc"] = ac, acd, rec_s, rec_d
    r["rec_simple_txt"] = _meses_txt(rec_s) if rec_s is not None else "No recupera"
    r["rec_desc_txt"] = _meses_txt(rec_d) if rec_d is not None else "No recupera"
    return r


# ---------- DataFrames ----------
semanal = pd.DataFrame(SEMANAL, columns=["semana", "programados", "ejecutados", "costo_tracto", "costo_tolva"])
semanal["cumplimiento"] = semanal["ejecutados"] / semanal["programados"] * 100
semanal["costo_total"] = semanal["costo_tracto"] + semanal["costo_tolva"]


def serie_tobe(m):
    """To-Be mensual con la regla oficial del Excel (hoja 'BASE DE DATOS - dashboard', nota en U16):
    'menor es mejor' -> TO-BE = AS-IS × (TO-BE anual / promedio AS-IS)
    'mayor es mejor' -> TO-BE = 100 − (100 − AS-IS) × (100 − TO-BE anual) / (100 − promedio AS-IS)
    Así el promedio anual de la serie coincide con el To-Be del Excel."""
    base = ASIS_MENSUAL[m["col"]]
    prom = sum(base) / len(base)
    if m["mejor"] == "mayor":
        f = (100 - m["to_be"]) / (100 - prom)
        return [round(100 - (100 - v) * f, 4) for v in base]
    f = m["to_be"] / prom
    return [round(v * f, 4) for v in base]


mensual = pd.DataFrame({"mes": MESES, "abrev": ABREV})
for m in MOTIVOS:
    mensual[m["col"] + "_asis"] = ASIS_MENSUAL[m["col"]]
    mensual[m["col"] + "_tobe"] = serie_tobe(m)
CAP_ASIST_TOBE = [round(TOTAL_COLAB * v / 100, 1) for v in mensual["capacitacion_tobe"]]

resumen = pd.DataFrame(MOTIVOS)

# ---------- Indicadores clave del Resumen Ejecutivo (todos calculados con los datos del proyecto) ----------
VIP = semanal["cumplimiento"].mean() / 100                  # cumplimiento OT As-Is = promedio de las 26 quincenas (Análisis Cuantitativo!F242)
VR = META_OT / 100
BRECHA_PP = (VR - VIP) * 100                                # Brecha técnica BT = VR − VIP (Planteamiento!E8), en puntos porcentuales
IMPACTO = semanal["costo_total"].sum()                      # Impacto económico = Σ costo tracto + tolva/cisterna (Análisis Cuantitativo!I242)
PCT_ING = IMPACTO / FACTURACION_ANUAL                       # Impacto / ingresos anuales (Planteamiento!D19)
FINRES = modelo_financiero(VIP, VR, IMPACTO, int(semanal["ejecutados"].sum()))


def mejora(as_is, to_be, mejor):
    return (to_be - as_is) / as_is * 100 if mejor == "mayor" else (as_is - to_be) / as_is * 100


def cumple(valor, est, mejor):
    return valor >= est if mejor == "mayor" else valor <= est


def soles(x, dec=0):
    return f"S/ {x:,.{dec}f}"


# ======================= ESTILOS =======================
CSS = """
<style>
/* ---- Base: siempre superficies claras y texto negro (modo claro Y modo oscuro) ---- */
:root, html, body, .stApp { color-scheme: light !important; }
html { font-size: 16.5px !important; }
.stApp { background: linear-gradient(135deg,#EAF0FA 0%,#FFFFFF 52%,#FCEFF1 100%) !important; }
[data-testid="stHeader"], header[data-testid="stHeader"] { background: transparent !important; }
.block-container { padding-top: 2.4rem; max-width: 1500px; }
.stApp, .stApp * { color: __NEGRO__ !important; }
.stApp [data-testid="stMarkdownContainer"], .stApp [data-testid="stMarkdownContainer"] *, .stApp label, .stApp label *,
.stApp [data-testid="stCaptionContainer"] * { -webkit-text-fill-color: __NEGRO__ !important; opacity: 1 !important; }
[data-testid="stToolbar"] svg, [data-testid="stHeader"] svg,
[data-testid="stExpandSidebarButton"] svg, [data-testid="stSidebarCollapsedControl"] svg { fill: __NEGRO__ !important; color: __NEGRO__ !important; }
.stApp [data-testid="stMarkdownContainer"] p, .stApp [data-testid="stMarkdownContainer"] li { font-size: 1.04rem; }
.stApp [data-testid="stCaptionContainer"] * { font-size: .92rem !important; }

/* ---- Panel lateral: barra de navegación azul marino y botones rojos ---- */
section[data-testid="stSidebar"] { border-right: 5px solid __ROJO__; }
section[data-testid="stSidebar"] > div { background: linear-gradient(180deg,#0B1F4B 0%,#0E2A66 100%) !important; }
.stApp section[data-testid="stSidebar"] *, .stApp section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] *,
.stApp section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * { color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; opacity: 1 !important; }
.stApp section[data-testid="stSidebar"] svg { fill: #FFFFFF !important; color: #FFFFFF !important; }
.stApp section[data-testid="stSidebar"] .stButton > button {
    width: 100%; justify-content: flex-start; text-align: left; border-radius: 10px; border: 1.5px solid rgba(255,255,255,.55);
    background: __ROJO__ !important; padding: .6rem .85rem; transition: all .2s; box-shadow: 0 2px 5px rgba(0,0,0,.30); }
.stApp section[data-testid="stSidebar"] .stButton > button p { font-size: 1rem !important; font-weight: 700 !important; text-align: left; }
.stApp section[data-testid="stSidebar"] .stButton > button:hover { background: #E0243F !important; border-color: #FFFFFF; transform: translateX(4px); }
.stApp section[data-testid="stSidebar"] .stButton > button[kind="primary"],
.stApp section[data-testid="stSidebar"] .stButton > button[data-testid="stBaseButton-primary"] {
    background: __ROJOOSC__ !important; border: 1.5px solid #FFFFFF; border-left: 8px solid #FFFFFF; }
.empresa { text-align: center; padding: 6px 0 2px 0; }
.empresa h2 { margin: 0; font-size: 1.55rem; font-weight: 900; letter-spacing: 1px; padding: 0; }
.empresa .sub { font-size: .95rem; font-weight: 700; }
.servicio { background: #FFFFFF !important; border-left: 6px solid __ROJO__; border-radius: 10px; padding: 8px 12px; margin: 7px 0;
            font-size: 1rem; font-weight: 700; box-shadow: 0 1px 3px rgba(11,31,75,.12); }
.stApp .seccion { display: inline-block; font-size: .9rem; font-weight: 900; letter-spacing: 1.6px; margin: 16px 0 8px 0; padding-bottom: 3px; border-bottom: 3px solid __ROJO__; }

.stApp section[data-testid="stSidebar"] div.servicio, .stApp section[data-testid="stSidebar"] div.servicio * { color: __NEGRO__ !important; -webkit-text-fill-color: __NEGRO__ !important; }

/* ---- Encabezados ---- */
.stApp .banner { background: linear-gradient(100deg,#DCE6F7 0%,#FFFFFF 78%); border: 1px solid #CFD8E8; border-left: 12px solid __NAVY__;
          border-bottom: 5px solid __ROJO__; border-radius: 16px; padding: 20px 28px; box-shadow: 0 6px 18px rgba(11,31,75,.10); margin-bottom: 16px; }
.stApp .banner h1 { margin: 0; padding: 0; font-size: 2.1rem; font-weight: 900; line-height: 1.2; }
.stApp .banner p { margin: 6px 0 0 0; font-size: 1.08rem; }
.chip { display: inline-block; padding: 6px 14px; border-radius: 999px; background: #FFFFFF; font-size: .95rem; font-weight: 700;
        margin: 10px 10px 0 0; border: 1.5px solid __NAVY2__; }
.sec-titulo { font-size: 1.5rem; font-weight: 900; margin: 26px 0 12px 0; padding-left: 14px; border-left: 7px solid __ROJO__; line-height: 1.25; }
.sec-titulo.azul { border-left-color: __NAVY__; }
.sec-sub { font-size: .98rem; margin: -6px 0 12px 21px; }

/* ---- Tarjetas ---- */
.kpi { background: #FFFFFF; border: 1px solid #D5DCE8; border-top: 7px solid __NAVY__; border-radius: 14px; padding: 13px 17px 15px;
       box-shadow: 0 5px 14px rgba(11,31,75,.10); min-height: 135px; }
.kpi .t { font-size: .86rem; font-weight: 900; letter-spacing: .5px; text-transform: uppercase; }
.kpi .s { font-size: .95rem; }
.kpi .v { font-size: 2.3rem; font-weight: 900; line-height: 1.15; margin: 5px 0; }
.bloque { background: #FFFFFF; border: 1px solid #D5DCE8; border-top: 9px solid __NAVY__; border-radius: 16px; padding: 15px 20px 18px;
          box-shadow: 0 6px 18px rgba(11,31,75,.12); min-height: 262px; margin-bottom: 4px; }
.b-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; margin-bottom: 4px; min-height: 2.7rem; }
.b-tit { font-size: 1.02rem; font-weight: 900; letter-spacing: .7px; }
.b-tag { font-size: .72rem; font-weight: 800; letter-spacing: 1px; border: 2px solid __NAVY__; border-radius: 999px; padding: 2px 10px; background: #FFFFFF; white-space: nowrap; }
.b-val { font-size: 2.75rem; font-weight: 900; line-height: 1.1; margin: 6px 0 4px 0; }
.b-form { font-size: .95rem; font-weight: 700; border-top: 1px dashed #AEBBD6; padding-top: 7px; margin-top: 6px; }
.b-det { font-size: .95rem; margin: 5px 0 11px 0; }
.b-estado { display: inline-block; font-size: .95rem; font-weight: 700; padding: 6px 12px; border-radius: 10px; border: 1.5px solid __NAVY2__; background: #E1EAF8; }
.b-estado.rojo { background: #FBE3E7; border-color: __ROJO__; }
.estado-chip { display: inline-block; font-weight: 800; padding: 3px 10px; border-radius: 8px; border: 1.5px solid __NAVY2__; background: #E1EAF8; }
.estado-chip.rojo { background: #FBE3E7; border-color: __ROJO__; }

/* ---- Tablas (HTML propio: siempre negro sobre claro) ---- */
.tabla-wrap { overflow: auto; border: 1px solid #CFD8E8; border-radius: 12px; background: #FFFFFF; box-shadow: 0 3px 10px rgba(11,31,75,.08); margin-bottom: 10px; }
.stApp table.tabla { width: 100%; border-collapse: collapse; font-size: 1rem; margin: 0; }
.stApp table.tabla th { position: sticky; top: 0; background: #DCE6F7 !important; border: none; border-bottom: 3px solid __ROJO__; text-align: left;
                        padding: .55rem .8rem; font-weight: 800; white-space: nowrap; }
.stApp table.tabla td { padding: .5rem .8rem; border: none; border-bottom: 1px solid #E3E8F1; background: #FFFFFF !important; }
.stApp table.tabla tr:nth-child(even) td { background: #F4F7FC !important; }
.stApp table.tabla td.num, .stApp table.tabla th.num { text-align: right; white-space: nowrap; }
.stApp table.tabla tr.total td { background: #E1EAF8 !important; font-weight: 800; border-top: 2px solid __NAVY__; }

/* ---- Controles y contenedores ---- */
[data-testid="stPlotlyChart"] { background: #FFFFFF; border: 1px solid #E0E6F0; border-radius: 14px; box-shadow: 0 4px 12px rgba(11,31,75,.08); padding: 6px 8px; }
[data-testid="stExpander"] { background: #FFFFFF !important; border: 1.5px solid #CFD8E8 !important; border-radius: 12px !important; }
[data-testid="stExpander"] details, [data-testid="stExpander"] summary { background: #FFFFFF !important; }
[data-testid="stExpander"] summary p { font-size: 1.05rem; font-weight: 800; }
[data-testid="stBaseButton-pills"], button[kind="pills"] { background: #FFFFFF !important; border: 1.5px solid __ROJO__ !important; }
[data-testid="stBaseButton-pillsActive"], button[kind="pillsActive"] { background: __ROJO__ !important; border: 2px solid __ROJOOSC__ !important; }
.stApp button[data-testid="stBaseButton-pillsActive"] *, .stApp button[kind="pillsActive"] * { color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; }
[data-testid^="stBaseButton-pills"] p { font-size: 1rem !important; font-weight: 700 !important; }
[data-testid="stRadio"] label p { font-size: 1rem !important; font-weight: 700 !important; }
div[data-baseweb="select"] > div { background: #FFFFFF !important; }
</style>
"""
for _k, _v in dict(NEGRO=NEGRO, NAVY=NAVY, NAVY2=NAVY2, ROJO=ROJO, ROJOOSC=ROJO_OSC).items():
    CSS = CSS.replace(f"__{_k}__", _v)
st.markdown(CSS, unsafe_allow_html=True)

FUENTE = dict(family="Arial, Helvetica, sans-serif", size=13, color=NEGRO)


def estilo(fig, alto=None):
    """Estilo corporativo común: fondo blanco, texto negro y legible en modo claro y oscuro (se usa con theme=None)."""
    fig.update_layout(template="plotly_white", font=FUENTE, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                      title_font=dict(size=17, color=NEGRO),
                      legend=dict(font=dict(size=13, color=NEGRO), bgcolor="rgba(255,255,255,0)"),
                      hoverlabel=dict(bgcolor="#FFFFFF", font=dict(color=NEGRO, size=13), bordercolor=NAVY))
    eje = dict(tickfont=dict(size=12, color=NEGRO), title_font=dict(size=13, color=NEGRO),
               gridcolor="#E3E8F1", linecolor="#9AA5B8", zerolinecolor="#9AA5B8")
    fig.update_xaxes(**eje)
    fig.update_yaxes(**eje)
    if alto:
        fig.update_layout(height=alto)
    return fig


def grafico(fig):
    st.plotly_chart(fig, theme=None)


def tabla(df, alto=None, num=()):
    """Tabla HTML (reemplaza st.dataframe, que dibuja con los colores del tema y pierde legibilidad en modo oscuro)."""
    cab = "".join(f'<th class="{"num" if c in num else ""}">{c}</th>' for c in df.columns)
    filas = ""
    for row in df.itertuples(index=False):
        filas += "<tr>" + "".join(f'<td class="{"num" if c in num else ""}">{v}</td>' for c, v in zip(df.columns, row)) + "</tr>"
    est = f' style="max-height:{alto}px;"' if alto else ""
    st.markdown(f'<div class="tabla-wrap"{est}><table class="tabla"><thead><tr>{cab}</tr></thead><tbody>{filas}</tbody></table></div>',
                unsafe_allow_html=True)


def kpi(titulo, valor, sub, color):
    st.markdown(f'<div class="kpi" style="border-top-color:{color};"><div class="t">{titulo}</div>'
                f'<div class="v">{valor}</div><div class="s">{sub}</div></div>', unsafe_allow_html=True)


def bloque(icono, titulo, valor, formula, detalle, estado, color, tag, tono="azul"):
    """Bloque independiente de un indicador clave del Resumen Ejecutivo."""
    st.markdown(f'<div class="bloque" style="border-top-color:{color};">'
                f'<div class="b-head"><span class="b-tit">{icono} {titulo}</span><span class="b-tag" style="border-color:{color};">{tag}</span></div>'
                f'<div class="b-val">{valor}</div><div class="b-form">{formula}</div><div class="b-det">{detalle}</div>'
                f'<div class="b-estado {"rojo" if tono == "rojo" else ""}">{estado}</div></div>', unsafe_allow_html=True)


def banner(titulo, sub, color, color2, chips=()):
    ch = "".join(f'<span class="chip">{c}</span>' for c in chips)
    st.markdown(f'<div class="banner" style="border-left-color:{color};border-bottom-color:{color2};">'
                f'<h1>{titulo}</h1><p>{sub}</p>{ch}</div>', unsafe_allow_html=True)


def titulo_seccion(texto, color="rojo", sub=None):
    st.markdown(f'<div class="sec-titulo {"azul" if color == "azul" else ""}">{texto}</div>'
                + (f'<div class="sec-sub">{sub}</div>' if sub else ""), unsafe_allow_html=True)


LOGO_B64 = "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"

# ======================= ESTADO (interconexión) =======================
if "pagina" not in st.session_state:
    st.session_state.pagina = "resumen"
if "meses_sel" not in st.session_state:
    st.session_state.meses_sel = []     # vacío = todos los meses
if "escenario" not in st.session_state:
    st.session_state.escenario = "Ambos"


def ir(p):
    st.session_state.pagina = p


# ======================= PANEL IZQUIERDO =======================
with st.sidebar:
    st.markdown(f"""<div class="empresa"><div style="background:#FFFFFF;border-radius:16px;padding:10px 14px;margin:4px 0 10px 0;box-shadow:0 4px 12px rgba(11,31,75,.25);border-bottom:4px solid {ROJO}"><img src="data:image/png;base64,{LOGO_B64}" style="width:100%;max-width:190px"></div>
        <h2>TRANSVISA E.I.R.L.</h2><div class="sub">TGM · Año 2025</div></div>""",
                unsafe_allow_html=True)
    st.markdown('<div class="seccion">SERVICIOS</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="servicio" style="border-left-color:{NAVY}">🚚 Transporte de carga pesada</div>
    <div class="servicio" style="border-left-color:{AZUL_CLARO}">🛢️ Fabricación de cisternas</div>
    <div class="servicio" style="border-left-color:{ROJO}">🔧 Mantenimiento</div>""", unsafe_allow_html=True)

    st.markdown('<div class="seccion">PANEL</div>', unsafe_allow_html=True)
    st.button("📋  Resumen ejecutivo", on_click=ir, args=("resumen",), key="b_res",
              type="primary" if st.session_state.pagina == "resumen" else "secondary")
    st.button("🚨  Problema general", on_click=ir, args=("general",), key="b_gen",
              type="primary" if st.session_state.pagina == "general" else "secondary")

    st.markdown('<div class="seccion">LAS 5 CAUSAS</div>', unsafe_allow_html=True)
    puntos = {"C1": "🔵", "C2": "🔷", "C3": "🔹", "C4": "🔴", "C5": "🟥"}
    # Solo cambia el orden de visualización del menú; cada causa conserva su id (C1..C5) y todos sus datos.
    ORDEN_MENU = ["C1", "C3", "C5", "C2", "C4"]
    por_id = {x["id"]: x for x in MOTIVOS}
    for n_menu, cid in enumerate(ORDEN_MENU, start=1):
        mm = por_id[cid]
        st.button(f"{puntos[cid]} {mm['icono']}  CAUSA {n_menu} · {mm['nombre'].upper()} ({mm['metodo']})", on_click=ir, args=(cid,),
                  key="b_" + cid, type="primary" if st.session_state.pagina == cid else "secondary")
    st.caption("Fuente: Excel del equipo (hojas Ficha Técnica, Planteamiento, Análisis Cuantitativo y Flujo Caja)")

pagina = st.session_state.pagina

# ======================= PÁGINA: RESUMEN EJECUTIVO =======================
if pagina == "resumen":
    banner("📋 RESUMEN EJECUTIVO", "Baja eficacia del cumplimiento de mantenimiento preventivo · TGM / Transvisa E.I.R.L.",
           NAVY, ROJO, ["🎯 5 causas raíz", "📈 As-Is vs To-Be", "🧰 5 metodologías", "💰 Evaluación a 10 años"])
    cumpl = VIP * 100
    ok = sum(cumple(m["to_be"], m["estandar"], m["mejor"]) for m in MOTIVOS)
    F = FINRES
    mayor_c = max(F["causas"].items(), key=lambda kv: kv[1]["ahorro"])
    nombre_c = {m["id"]: f"{m['id']} {m['metodo']}" for m in MOTIVOS}
    tir_ok = F["tir"] is not None and F["tir"] > F["cppc"]

    titulo_seccion("🔑 Indicadores clave del proyecto", sub="Calculados con los datos reales del Excel del proyecto; sin supuestos adicionales.")
    r1 = st.columns(3)
    with r1[0]:
        bloque("📏", "BRECHA TÉCNICA", f"{BRECHA_PP:.1f} p.p.", "BT = VR − VIP = " + f"{META_OT:.1f}% − {cumpl:.1f}%",
               f"VIP = promedio del % de cumplimiento de OT de las {len(semanal)} quincenas 2025 ({int(semanal['ejecutados'].sum())} de {int(semanal['programados'].sum())} OT ejecutadas)",
               f"Faltan {BRECHA_PP:.1f} puntos porcentuales para la meta", ROJO, "DIAGNÓSTICO", "rojo")
    with r1[1]:
        bloque("💰", "IMPACTO ECONÓMICO", soles(IMPACTO), "Σ costo tracto + costo tolva/cisterna de las quincenas 2025",
               f"Tracto {soles(semanal['costo_tracto'].sum())} · Tolva/cisterna {soles(semanal['costo_tolva'].sum())}",
               f"{PCT_ING * 100:.1f}% de los ingresos anuales ({soles(FACTURACION_ANUAL)})", ROJO, "DIAGNÓSTICO", "rojo")
    with r1[2]:
        bloque("🧾", "INVERSIÓN TOTAL", soles(F["inv_total"]), "Σ (cantidad × costo unitario) + 5% de imprevistos",
               f"{len(INVERSION)} ítems {soles(F['inv_sub'])} + imprevistos {soles(F['imprev'])} · 5 propuestas (ADDIE, ISOS, TPM, DMAIC, SMED)",
               f"Año 0 · préstamo {soles(F['deuda'])} + aporte propio {soles(F['equity'])}", NAVY, "PROPUESTA")
    r2 = st.columns(3)
    with r2[0]:
        bloque("📉", "AHORRO ANUAL", soles(F["ahorro"]), "Σ (costo AS-IS − costo TO-BE) de las 5 causas",
               f"Costo AS-IS {soles(F['costo_asis'])} → TO-BE {soles(F['costo_tobe'])} · año de operación plena",
               f"Mayor aporte: {nombre_c[mayor_c[0]]} ({mayor_c[1]['ahorro'] / F['ahorro'] * 100:.1f}% del ahorro)", NAVY, "PROPUESTA")
    with r2[1]:
        bloque("📈", "VAN", soles(F["van"]), "VAN = −I<sub>0</sub> + Σ FC<sub>t</sub> / (1 + CPPC)<sup>t</sup>",
               f"Descontado al CPPC de {F['cppc'] * 100:.2f}% · horizonte de {len(F['flujo']) - 1} años",
               "✔ VAN &gt; 0: el proyecto genera valor" if F["van"] > 0 else "✖ VAN &lt; 0: el proyecto no es rentable",
               NAVY, "PROPUESTA", "azul" if F["van"] > 0 else "rojo")
    with r2[2]:
        bloque("🔁", "TIR", f"{F['tir'] * 100:.1f}%" if F["tir"] is not None else "N/D", "TIR: tasa que hace VAN = 0 (flujos de los años 0 a 10)",
               f"Tasa de corte CPPC {F['cppc'] * 100:.2f}% · COK {F['cok'] * 100:.2f}%",
               "✔ TIR &gt; CPPC: rinde más de lo que cuesta financiarse" if tir_ok else "✖ TIR &lt; CPPC", NAVY, "PROPUESTA",
               "azul" if tir_ok else "rojo")

    titulo_seccion("🧭 Contexto operativo", "azul")
    c1, c2, c3 = st.columns(3)
    with c1: kpi("✅ Cumplimiento OT (As-Is)", f"{cumpl:.1f}%", f"Meta {META_OT:.0f}% · {cumpl - META_OT:.1f} pp", ROJO)
    with c2: kpi("📝 OT ejecutadas / programadas", f"{int(semanal['ejecutados'].sum())} / {int(semanal['programados'].sum())}", "Total del año 2025", AZUL)
    with c3: kpi("🏆 To-Be cumple estándar", f"{ok} de {len(MOTIVOS)}", "Causas que alcanzan su estándar", NAVY)

    titulo_seccion("🎯 Las 5 causas de un vistazo")
    cols = st.columns(5)
    for col, m in zip(cols, MOTIVOS):
        u = "%" if m["unidad"] == "%" else f" {m['unidad'][:3]}."
        mj = mejora(m["as_is"], m["to_be"], m["mejor"])
        c = cumple(m["to_be"], m["estandar"], m["mejor"])
        with col:
            kpi(f"{m['icono']} {m['id']} · {m['nombre']}", f"{m['to_be']:,.1f}{u}",
                f"As-Is {m['as_is']:,.1f}{u} · ▲ {mj:.0f}% · {'✔ cumple' if c else '✖ no cumple'}", m["color"])

    resumen["mejora_pct"] = resumen.apply(lambda r: mejora(r.as_is, r.to_be, r.mejor), axis=1)
    resumen["etiqueta"] = resumen["id"] + " · " + resumen["nombre"]
    fig = px.bar(resumen, x="etiqueta", y="mejora_pct", color="etiqueta",
                 color_discrete_sequence=[m["color"] for m in MOTIVOS],
                 text=resumen["mejora_pct"].round(1).astype(str) + "%",
                 title="📈 Mejora esperada por causa (To-Be vs As-Is)", labels={"mejora_pct": "% de mejora", "etiqueta": ""})
    fig.update_traces(textposition="outside", textfont=dict(size=13, color=NEGRO), cliponaxis=False)
    fig.update_layout(showlegend=False)
    grafico(estilo(fig))

    tb = resumen[["id", "nombre", "unidad", "as_is", "to_be", "estandar", "metodo"]].copy()
    tb["Cumple"] = [("✔ Sí" if cumple(r.to_be, r.estandar, r.mejor) else "✖ No") for r in resumen.itertuples()]
    for c_ in ("as_is", "to_be", "estandar"):
        tb[c_] = tb[c_].map(lambda v: f"{v:,.1f}")
    tb.columns = ["ID", "Causa", "Unidad", "As-Is", "To-Be", "Estándar", "Metodología", "To-Be cumple estándar"]
    tabla(tb, num=("As-Is", "To-Be", "Estándar"))

# ======================= PÁGINA: PROBLEMA GENERAL =======================
elif pagina == "general":
    banner("🚨 PROBLEMA GENERAL", "Baja eficacia del cumplimiento de mantenimiento preventivo", ROJO, NAVY,
           ["📅 Periodos quincenales 2025", f"🎯 Meta {META_OT:.0f}% de cumplimiento de OT"])
    prog, ejec = semanal["programados"].sum(), semanal["ejecutados"].sum()
    cumpl = semanal["cumplimiento"].mean()
    costo = semanal["costo_total"].sum()
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi("✅ Cumplimiento OT", f"{cumpl:.1f}%", f"{cumpl - META_OT:.1f} pp vs meta {META_OT:.0f}%", ROJO)
    with c2: kpi("📝 OT ejecutadas / programadas", f"{ejec} / {prog}", "Total del año", AZUL)
    with c3: kpi("💰 Impacto económico", f"S/ {costo:,.0f}", "Tracto + tolva/cisterna", ROJO_OSC)
    with c4: kpi("📉 % facturación anual", f"{costo / FACTURACION_ANUAL * 100:.1f}%", "Sobre ingresos anuales", NAVY)

    fig = px.line(semanal, x="semana", y="cumplimiento", markers=True,
                  title="📈 % de cumplimiento de órdenes de trabajo por quincena")
    fig.update_traces(line_color=ROJO, line_width=3, marker=dict(size=9, color=NAVY))
    fig.add_hline(y=META_OT, line_dash="dash", line_color=NAVY, line_width=2, annotation_text=f"Meta {META_OT:.0f}%",
                  annotation_position="top left", annotation_font=dict(color=NEGRO, size=13))
    fig.update_layout(yaxis_title="% cumplimiento", xaxis_title="", yaxis_range=[0, 110])
    grafico(estilo(fig))

    costos = semanal.melt(id_vars="semana", value_vars=["costo_tracto", "costo_tolva"], var_name="tipo", value_name="costo")
    costos["tipo"] = costos["tipo"].map({"costo_tracto": "🚛 Tracto", "costo_tolva": "🛢️ Tolva / cisterna"})
    fig2 = px.bar(costos, x="semana", y="costo", color="tipo", title="💸 Costo por periodo (S/)",
                  color_discrete_map={"🚛 Tracto": NAVY, "🛢️ Tolva / cisterna": ROJO},
                  labels={"costo": "S/", "semana": "", "tipo": ""})
    grafico(estilo(fig2))

    with st.expander("📑 Ver datos del cumplimiento semanal"):
        sv = semanal.copy()
        for c_ in ("costo_tracto", "costo_tolva", "costo_total"):
            sv[c_] = sv[c_].map(lambda v: f"{v:,.2f}")
        sv["cumplimiento"] = sv["cumplimiento"].map(lambda v: f"{v:,.1f}%")
        sv.columns = ["Periodo", "OT programadas", "OT ejecutadas", "Costo tracto (S/)", "Costo tolva/cisterna (S/)", "% cumplimiento", "Costo total (S/)"]
        tabla(sv, alto=420, num=tuple(sv.columns[1:]))

# ======================= PÁGINAS: CAUSAS C1..C5 =======================
else:
    m = next(x for x in MOTIVOS if x["id"] == pagina)
    color = m["color"]
    u = "%" if m["unidad"] == "%" else f" {m['unidad']}"

    banner(f"{m['icono']} CAUSA {ORDEN_MENU.index(m['id']) + 1} · {m['nombre'].upper()}",
           f"Causa raíz: {m['causa']}", color, ROJO if color != ROJO else NAVY,
           [f"📐 Indicador: {m['ind']}", f"🧰 Metodología: {m['metodo']}",
            f"📏 Parámetro (estándar): {m['estandar']:,.1f}{u}"])
    with st.expander("🔎 Fórmula y herramienta / metodología aplicada", expanded=False):
        st.markdown(f"**📐 Indicador:** {m['formula']}")
        st.markdown(f"**🧰 Herramienta / metodología:** {m['metodo_det']}")
        st.markdown(f"**📏 Parámetro (estándar):** {m['estandar']:,.1f}{u} "
                    f"({'mayor es mejor' if m['mejor'] == 'mayor' else 'menor es mejor'})")

    # ---------- Filtros interconectados (botones de mes + As-Is / To-Be) ----------
    f1, f2 = st.columns([4, 1.4])
    with f1:
        st.markdown("**📅 Meses** (sin selección = todo el año)")
        if hasattr(st, "pills"):
            sel = st.pills("Meses", ABREV, selection_mode="multi", key="meses_sel", label_visibility="collapsed")
        else:
            sel = st.multiselect("Meses", ABREV, key="meses_sel", label_visibility="collapsed")
    with f2:
        st.markdown("**🔀 Escenario**")
        st.radio("Escenario", ["Ambos", "As-Is", "To-Be"], key="escenario", horizontal=True,
                 label_visibility="collapsed")
    sel = sel or ABREV
    esc = st.session_state.escenario
    df = mensual[mensual["abrev"].isin(sel)].reset_index(drop=True)
    a_col, t_col = m["col"] + "_asis", m["col"] + "_tobe"

    val_asis, val_tobe = df[a_col].mean(), df[t_col].mean()
    mj = mejora(val_asis, val_tobe, m["mejor"])
    cumple_tobe = cumple(val_tobe, m["estandar"], m["mejor"])
    meses_txt = "Todo el año" if len(sel) == 12 else ", ".join(sel)

    k = st.columns(5)
    with k[0]: kpi("🔴 As-Is · Situación actual", f"{val_asis:,.1f}{u}", f"Promedio · {meses_txt}", ROJO)
    with k[1]: kpi("🔵 To-Be · Simulación", f"{val_tobe:,.1f}{u}", f"Aplicando {m['metodo']}", NAVY)
    with k[2]: kpi("🎯 Estándar / Meta", f"{m['estandar']:,.1f}{u}", "Valor de referencia (parámetro)", GRIS)
    with k[3]: kpi("🚀 Mejora esperada", f"{'▲' if mj >= 0 else '▼'} {abs(mj):.1f}%", "Frente al As-Is", AZUL)
    with k[4]:
        kpi("🏁 Estado del To-Be", "CUMPLE ✔" if cumple_tobe else "NO CUMPLE ✖", "Comparado con el estándar",
            NAVY if cumple_tobe else ROJO)

    g1, g2 = st.columns([1.6, 1])
    # ---------- Evolución mensual ----------
    with g1:
        fig = go.Figure()
        if esc in ("Ambos", "As-Is"):
            fig.add_trace(go.Scatter(x=df["abrev"], y=df[a_col], mode="lines+markers+text", name="As-Is",
                                     text=[f"{v:,.1f}" for v in df[a_col]], textposition="bottom center",
                                     textfont=dict(size=11, color=NEGRO),
                                     line=dict(color=ROJO, width=3), marker=dict(size=10, color=ROJO)))
        if esc in ("Ambos", "To-Be"):
            fig.add_trace(go.Scatter(x=df["abrev"], y=df[t_col], mode="lines+markers+text", name="To-Be",
                                     text=[f"{v:,.1f}" for v in df[t_col]], textposition="top center",
                                     textfont=dict(size=11, color=NEGRO),
                                     line=dict(color=NAVY, width=3), marker=dict(size=10, color=NAVY)))
        fig.add_hline(y=m["estandar"], line_dash="dot", line_color=GRIS, line_width=2,
                      annotation_text=f"Estándar: {m['estandar']:,.1f}{u}", annotation_position="top right",
                      annotation_font=dict(color=NEGRO, size=12))
        fig.update_layout(title=f"📈 Evolución mensual: {esc} · {m['ind']}", yaxis_title=m["unidad"], xaxis_title="",
                          legend=dict(orientation="h", y=1.12))
        grafico(estilo(fig, 400))

    # ---------- Termómetro (gauge) ----------
    with g2:
        val_g = val_asis if esc == "As-Is" else val_tobe
        etiqueta = "As-Is" if esc == "As-Is" else "To-Be"
        max_g = 100 if m["unidad"] == "%" else max(m["as_is"], m["estandar"], max(df[a_col])) * 1.25
        mal, bien = "#F6CBD2", "#CFDDF3"
        if m["mejor"] == "mayor":
            pasos = [dict(range=[0, m["estandar"]], color=mal), dict(range=[m["estandar"], max_g], color=bien)]
        else:
            pasos = [dict(range=[0, m["estandar"]], color=bien), dict(range=[m["estandar"], max_g], color=mal)]
        ok_g = cumple(val_g, m["estandar"], m["mejor"])
        gauge = go.Figure(go.Indicator(
            mode="gauge+number", value=round(val_g, 1),
            number=dict(suffix="%" if m["unidad"] == "%" else "", font=dict(size=45, color=NEGRO)),
            title=dict(text=f"🌡️ Termómetro vs estándar · {etiqueta}", font=dict(size=16, color=NEGRO)),
            gauge=dict(axis=dict(range=[0, max_g], tickfont=dict(size=12, color=NEGRO), tickcolor=NEGRO),
                       bar=dict(color=NAVY if ok_g else ROJO), steps=pasos, bgcolor="#FFFFFF", borderwidth=1, bordercolor="#9AA5B8",
                       threshold=dict(line=dict(color=NEGRO, width=5), thickness=0.85, value=m["estandar"]))))
        gauge.update_layout(margin=dict(t=80, b=10))
        grafico(estilo(gauge, 400))

    h1, h2 = st.columns([1.6, 1])
    # ---------- Gráfico inferior ----------
    with h1:
        fig3 = go.Figure()
        tf = dict(size=12, color=NEGRO)
        if m["col"] == "capacitacion":
            idx = [ABREV.index(a) for a in df["abrev"]]
            if esc in ("Ambos", "As-Is"):
                fig3.add_trace(go.Bar(x=df["abrev"], y=[CAP_ASIST_ASIS[i] for i in idx], name="Asistieron As-Is",
                                      marker_color=ROJO, text=[CAP_ASIST_ASIS[i] for i in idx], textposition="outside", textfont=tf))
            if esc in ("Ambos", "To-Be"):
                fig3.add_trace(go.Bar(x=df["abrev"], y=[CAP_ASIST_TOBE[i] for i in idx], name="Asistieron To-Be",
                                      marker_color=NAVY, text=[CAP_ASIST_TOBE[i] for i in idx], textposition="outside", textfont=tf))
            fig3.add_hline(y=TOTAL_COLAB, line_dash="dot", line_color=GRIS,
                           annotation_text=f"Total colaboradores: {TOTAL_COLAB}", annotation_position="top right",
                           annotation_font=dict(color=NEGRO, size=12))
            fig3.update_layout(title="👥 Colaboradores que asistieron a la capacitación por mes", barmode="group")
        else:
            sig = 1 if m["mejor"] == "mayor" else -1
            if esc in ("Ambos", "As-Is"):
                brecha = [(v - m["estandar"]) * sig for v in df[a_col]]
                fig3.add_trace(go.Bar(x=df["abrev"], y=brecha, name="Brecha As-Is", marker_color=ROJO,
                                      text=[f"{b:+,.1f}" for b in brecha], textposition="outside", textfont=tf))
            if esc in ("Ambos", "To-Be"):
                brecha = [(v - m["estandar"]) * sig for v in df[t_col]]
                fig3.add_trace(go.Bar(x=df["abrev"], y=brecha, name="Brecha To-Be", marker_color=NAVY,
                                      text=[f"{b:+,.1f}" for b in brecha], textposition="outside", textfont=tf))
            fig3.add_hline(y=0, line_color=NEGRO)
            fig3.update_layout(title="📊 Brecha vs estándar por mes (positivo = cumple)", barmode="group")
        fig3.update_traces(cliponaxis=False)
        fig3.update_layout(xaxis_title="", legend=dict(orientation="h", y=1.15))
        grafico(estilo(fig3, 360))

    # ---------- Detalle mensual ----------
    with h2:
        st.markdown(f"**📋 Detalle mensual ({m['unidad']})**")
        det = pd.DataFrame({"Mes": df["abrev"], "As-Is": df[a_col].round(1), "To-Be": df[t_col].round(1),
                            "Estándar": m["estandar"]})
        if esc == "As-Is":
            det = det.drop(columns=["To-Be"])
        elif esc == "To-Be":
            det = det.drop(columns=["As-Is"])
        for c_ in det.columns[1:]:
            det[c_] = det[c_].map(lambda v: f"{v:,.1f}")
        tabla(det, alto=320, num=tuple(det.columns[1:]))
        