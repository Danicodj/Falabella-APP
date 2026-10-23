import os
import glob
import json
import math
import warnings
import pandas as pd

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

print("==================================================")
print("   PROCESANDO DATOS SENSING (SUPERVISIÓN AVANZADA)")
print("==================================================")

def obtener_archivo_mas_reciente(carpeta):
    archivos = glob.glob(os.path.join(carpeta, "*.xlsx")) + glob.glob(os.path.join(carpeta, "*.xls"))
    if not archivos:
        raise FileNotFoundError(f"No se encontró archivo en: '{carpeta}'")
    archivos.sort(key=os.path.getmtime)
    return archivos[-1]

STORE_MAP = {
    'TITAN': 'FALABELLA TITAN',
    'COLINA': 'FALABELLA COLINA',
    'CENTRO MAYOR': 'FALABELLA CENTRO MAYOR',
    'SANTA FE': 'FALABELLA SANTAFE',
    'SANTAFE MEDELLIN': 'FALABELLA SANTAFE MEDELLIN',
    'SAN DIEGO': 'FALABELLA SAN DIEGO',
    'WTC CALI': 'FALABELLA WTC Cali',
    'JARDIN PLAZA': 'FALABELLA JARDIN PLAZA',
    'UNICENTRO': 'FALABELLA UNICENTRO',
    'FONTANAR': 'FALABELLA FONTANAR',
    'ACQUA': 'FALABELLA ACQUA',
    'ARKADIA': 'FALABELLA ARKADIA',
    'EL CASTILLO': 'FALABELLA El Castillo',
    'CARACOLI': 'FALABELLA  CARACOLI',
    'BUENAVISTA': 'FALABELLA Buenavista',
    'LA CAROLA': 'FALABELLA  LA CAROLA',
    'GALERIAS': 'FALABELLA GALERIAS',
    'PARQUE ARBOLEDA': 'FALABELLA ARBOLEDA',
    'DIVER PLAZA': 'FALABELLA DIVER PLAZA',
    'HAYUELOS': 'FALABELLA HAYUELOS',
    'LA FELICIDAD': 'FALABELLA LA FELICIDAD',
    'PLAZA CENTRAL': 'FALABELLA PLAZA',
    'PRIMAVERA': 'FALABELLA PRIMAVER',
    'TIENDA ALEGRA': 'FALABELLA ALLEGRA BARRANQUILLA',
    'TIENDA FALABELLA .COM': 'FALABELLA ONLINE',
    'FALABELLA. COM ALTERNO': 'FALABELLA ONLINE // .COM ALTERNATIVO',
    'SUBA': 'FALABELLA PLAZA IMPERIAL'
}

r_tgt = obtener_archivo_mas_reciente('1_target')
r_nhc = obtener_archivo_mas_reciente('2_nhc')
r_vta = obtener_archivo_mas_reciente('3_ventas')
r_inv = obtener_archivo_mas_reciente('4_inventario')
r_ihs = obtener_archivo_mas_reciente('5_ihs')

df_tgt_tv = pd.read_excel(r_tgt, sheet_name='DISTRIB SEGM TV', skiprows=1)
df_tgt_av = pd.read_excel(r_tgt, sheet_name='DIST POS AV')
df_nhc = pd.read_excel(r_nhc, sheet_name='NHC')
df_dir = pd.read_excel(r_nhc, sheet_name='DIRECTORY ALL')
df_vta = pd.read_excel(r_vta)
df_inv = pd.read_excel(r_inv)
df_ihs_raw = pd.read_excel(r_ihs, sheet_name='IHS by POS', skiprows=4)

# 1. SUPERVISORES EXCLUIDOS DE CONSULTORES
supervisores_dir = df_dir[df_dir['CATEGORIA'].astype(str).str.strip() == 'SUPERVISOR'].copy()
cedulas_supervisores = set(supervisores_dir['CEDULA'].astype(str).str.strip().unique())

# 2. IHS Y ASISTENCIA
df_ihs_samsung = df_ihs_raw[df_ihs_raw['Unnamed: 3'] == 'SAMSUNG'].copy()
df_ihs_samsung = df_ihs_samsung[['Suma de amt COP (CON IVA)', 'Unnamed: 7', 'Unnamed: 8']]
df_ihs_samsung.columns = ['Store_IHS', 'W41', 'W42']

ihs_dict = {}
for _, r in df_ihs_samsung.iterrows():
    tienda = str(r['Store_IHS']).strip()
    w41 = round(float(r['W41']) * 100, 1) if pd.notna(r['W41']) else 0.0
    w42 = round(float(r['W42']) * 100, 1) if pd.notna(r['W42']) else 0.0
    diff = round(w42 - w41, 1)
    ihs_dict[tienda] = {
        "w41": w41, "w42": w42, "diff": diff,
        "tendencia": "CRECIENDO" if diff >= 0 else "DECRECIENDO",
        "mensaje": f"▲ +{diff}% vs W41" if diff >= 0 else f"▼ {diff}% vs W41"
    }

asistencia_dict = {}
for _, row in df_ihs_raw.iloc[1:].iterrows():
    p = str(row.iloc[14]).strip()
    b = str(row.iloc[15]).strip()
    a = str(row.iloc[25]).strip()
    if b == 'SAMSUNG' and p and p != 'nan':
        asistencia_dict[p.upper()] = "Asistida" if a in ['1', '2', '3'] else "No Asistida"

# 3. SEGMENTOS
def clasificar_segmento(row):
    mod = str(row['Modelo']).upper()
    subc = str(row['Subclase']).upper()
    if 'AUDIO' in mod or 'BARRA' in mod or 'HW-' in mod:
        return 'AUDIO'
    if any(x in mod for x in ['UN75', 'UN85', 'UN98', 'QN75', 'QN85', 'QN98', '75"', '85"', '98"']) or '70 A 90' in subc:
        return '75↑'
    if any(x in mod for x in ['OLED', 'S90', 'S85', 'LS03', 'FRAME', 'QN65QN', 'QN55QN']):
        return 'PREMIUM'
    if 'QLED' in mod or 'MINI LED' in mod or 'M75H' in mod:
        return 'QLED / MINILED'
    return 'UHD'

df_vta['Segmento'] = df_vta.apply(clasificar_segmento, axis=1)
df_vta['Store_TGT'] = df_vta['Local'].map(STORE_MAP).fillna(df_vta['Local'])
df_vta['Fecha'] = pd.to_datetime(df_vta['Fecha'])
df_vta['Semana'] = 'W' + df_vta['Fecha'].dt.isocalendar().week.astype(str)

df_inv['Store_TGT'] = df_inv['Local'].map(STORE_MAP).fillna(df_inv['Local'])
inv_detalle = df_inv.groupby(['Store_TGT', 'ID Modelo', 'Producto'])['Stock Disponible'].sum().reset_index()

stock_por_tienda = {}
for tienda, grp in inv_detalle.groupby('Store_TGT'):
    skus_disp = grp[grp['Stock Disponible'] > 0].sort_values(by='Stock Disponible', ascending=False)
    stock_por_tienda[tienda] = {
        "total": int(skus_disp['Stock Disponible'].sum()),
        "skus": skus_disp.to_dict(orient='records')
    }

df_tgt_tv_fal = df_tgt_tv[df_tgt_tv['Account'] == 'FALABELLA'].copy()
df_tgt_av_fal = df_tgt_av[df_tgt_av['Account'] == 'FALABELLA'].copy()

df_nhc_fal = df_nhc[(df_nhc['CADENA'] == 'FALABELLA') & (df_nhc['ESTADO'] == 'ACTIVO')].copy()
df_nhc_fal = df_nhc_fal[~df_nhc_fal['CEDULA'].astype(str).str.strip().isin(cedulas_supervisores)].copy()

# 4. DATOS POR TIENDA
tiendas_falabella = set(df_tgt_tv_fal['Store'].dropna().unique()).union(set(df_tgt_av_fal['Store'].dropna().unique()))
datos_tiendas = {}

for pos in tiendas_falabella:
    ttv = df_tgt_tv_fal[df_tgt_tv_fal['Store'] == pos]
    tav = df_tgt_av_fal[df_tgt_av_fal['Store'] == pos]
    
    tmtd = ttv['TGT SEGM'].sum() + tav['TTL AV'].sum()
    
    v_t = df_vta[df_vta['Store_TGT'] == pos]
    vmtd = v_t['Unidades'].sum()
    vpesos = int(v_t['Venta neta'].sum()) if not v_t.empty else 0
    
    stk = stock_por_tienda.get(pos, {"total": 0, "skus": []})
    ihs_t = ihs_dict.get(pos.upper(), ihs_dict.get(pos, {"w41": 0.0, "w42": 0.0, "diff": 0.0, "mensaje": "Sin datos"}))
    asist = asistencia_dict.get(pos.upper(), "No Asistida")
    
    cumpl_t = round((vmtd / tmtd) * 100, 1) if tmtd > 0 else 0.0
    
    datos_tiendas[pos] = {
        "tienda": str(pos),
        "asistencia": str(asist),
        "target_mtd": int(math.ceil(tmtd)),
        "venta_mtd": int(math.ceil(vmtd)),
        "cumplimiento_mtd": float(cumpl_t),
        "debe_mtd": int(max(0, math.ceil(tmtd) - math.ceil(vmtd))),
        "venta_pesos": int(vpesos),
        "stock": int(stk['total']),
        "ihs_w42": float(ihs_t['w42']),
        "ihs_w41": float(ihs_t['w41']),
        "ihs_mensaje": str(ihs_t['mensaje'])
    }

# 5. CONSULTORES REALES
consultores_salida = []
grupos_cedula = df_nhc_fal.groupby('CEDULA')

for cedula, grupo in grupos_cedula:
    ced_str = str(cedula).strip()
    nombre = str(grupo['NOMBRE'].iloc[0]).strip()
    supervisor = str(grupo['SUPERVISOR'].iloc[0]).strip()
    ciudades = list(grupo['CIUDAD'].unique())
    tiendas_asignadas = list(grupo['NOMBRE POS'].unique())
    
    t_w43, v_w43, vp_w43 = 0.0, 0.0, 0.0
    t_w42, v_w42, vp_w42 = 0.0, 0.0, 0.0
    t_mtd, v_mtd, vp_mtd = 0.0, 0.0, 0.0
    seg_stats = {'75↑': [0, 0], 'PREMIUM': [0, 0], 'QLED / MINILED': [0, 0], 'UHD': [0, 0], 'AUDIO': [0, 0]}
    skus_totales = []
    total_stock_acumulado = 0
    ihs_info = {"w41": 0.0, "w42": 0.0, "diff": 0.0, "tendencia": "ESTABLE", "mensaje": "Sin datos"}
    
    for _, fila in grupo.iterrows():
        pos = fila['NOMBRE POS']
        disp = float(fila['DISPERSION DE CUPOS']) if pd.notna(fila['DISPERSION DE CUPOS']) else 1.0
        
        tgt_pos_tv = df_tgt_tv_fal[df_tgt_tv_fal['Store'] == pos]
        tgt_pos_av = df_tgt_av_fal[df_tgt_av_fal['Store'] == pos]
        
        tw43_t = (tgt_pos_tv['W43'].sum() if 'W43' in tgt_pos_tv.columns else 0) + (tgt_pos_av['W43'].sum() if 'W43' in tgt_pos_av.columns else 0)
        tw42_t = (tgt_pos_tv['W42'].sum() if 'W42' in tgt_pos_tv.columns else 0) + (tgt_pos_av['W42'].sum() if 'W42' in tgt_pos_av.columns else 0)
        tmtd_t = tgt_pos_tv['TGT SEGM'].sum() + tgt_pos_av['TTL AV'].sum()
        
        t_w43 += tw43_t * disp
        t_w42 += tw42_t * disp
        t_mtd += tmtd_t * disp
        
        v_tienda = df_vta[df_vta['Store_TGT'] == pos]
        vw43 = v_tienda[v_tienda['Semana'] == 'W43']
        vw42 = v_tienda[v_tienda['Semana'] == 'W42']
        
        v_w43 += vw43['Unidades'].sum() * disp
        vp_w43 += vw43['Venta neta'].sum() * disp
        
        v_w42 += vw42['Unidades'].sum() * disp
        vp_w42 += vw42['Venta neta'].sum() * disp
        
        v_mtd += v_tienda['Unidades'].sum() * disp
        vp_mtd += v_tienda['Venta neta'].sum() * disp
        
        for s in seg_stats.keys():
            v_seg = v_tienda[v_tienda['Segmento'] == s]['Unidades'].sum() * disp
            t_seg = (tgt_pos_tv[tgt_pos_tv['SEGMENTO AGENCIA'] == s]['TGT SEGM'].sum() if s != 'AUDIO' else tgt_pos_av['TTL AV'].sum()) * disp
            seg_stats[s][0] += v_seg
            seg_stats[s][1] += t_seg
            
        if pos in stock_por_tienda:
            total_stock_acumulado += stock_por_tienda[pos]['total']
            skus_totales.extend(stock_por_tienda[pos]['skus'])
            
        if pos in ihs_dict:
            ihs_info = ihs_dict[pos]

    tgt_w43_int = int(math.ceil(t_w43))
    vta_w43_int = int(math.ceil(v_w43))
    debe_w43_int = int(max(0, tgt_w43_int - vta_w43_int))
    cumpl_w43 = round((vta_w43_int / tgt_w43_int) * 100, 1) if tgt_w43_int > 0 else 0.0

    tgt_w42_int = int(math.ceil(t_w42))
    vta_w42_int = int(math.ceil(v_w42))
    debe_w42_int = int(max(0, tgt_w42_int - vta_w42_int))
    cumpl_w42 = round((vta_w42_int / tgt_w42_int) * 100, 1) if tgt_w42_int > 0 else 0.0

    tgt_mtd_int = int(math.ceil(t_mtd))
    vta_mtd_int = int(math.ceil(v_mtd))
    debe_mtd_int = int(max(0, tgt_mtd_int - vta_mtd_int))
    cumpl_mtd = round((vta_mtd_int / tgt_mtd_int) * 100, 1) if tgt_mtd_int > 0 else 0.0

    puntos_venta = int(vta_mtd_int * 10)
    puntos_bono = 500 if cumpl_mtd >= 100 else (200 if cumpl_mtd >= 80 else 50)
    puntos_totales = int(puntos_venta + puntos_bono)

    detalle_segmentos = []
    for seg, (vta, tgt) in seg_stats.items():
        vta_int = int(math.ceil(vta))
        tgt_int = int(math.ceil(tgt))
        debe_int = int(max(0, tgt_int - vta_int))
        c_pct = round((vta_int / tgt_int) * 100, 1) if tgt_int > 0 else 0.0
        texto_debe = "Al día" if debe_int == 0 else f"{debe_int} {'ud' if debe_int == 1 else 'uds'}"
        
        detalle_segmentos.append({
            "segmento": seg, "target": tgt_int, "venta": vta_int,
            "debe_unidades": debe_int, "cumplimiento_pct": c_pct, "estado": texto_debe
        })

    consultores_salida.append({
        "cedula": ced_str,
        "nombre": nombre,
        "tienda": " + ".join(tiendas_asignadas),
        "ciudades": ", ".join(ciudades),
        "supervisor": supervisor,
        "es_compartida": len(tiendas_asignadas) > 1,
        "ihs": ihs_info,
        "puntos": {
            "total": puntos_totales, "por_venta": puntos_venta, "bono": puntos_bono,
            "explicacion": f"{vta_mtd_int} unidades vendidas en el mes (x10 pts = {puntos_venta} pts) + bono de {puntos_bono} pts por cumplimiento MTD ({cumpl_mtd}%)."
        },
        "periodos": {
            "W43": {"label": "Semana en Curso (W43)", "target": tgt_w43_int, "venta_uds": vta_w43_int, "venta_pesos": int(round(vp_w43, 0)), "cumplimiento": cumpl_w43, "semaforo": "VERDE" if cumpl_w43 >= 100 else ("AMARILLO" if cumpl_w43 >= 80 else "ROJO"), "debe": debe_w43_int},
            "W42": {"label": "Semana Anterior (W42)", "target": tgt_w42_int, "venta_uds": vta_w42_int, "venta_pesos": int(round(vp_w42, 0)), "cumplimiento": cumpl_w42, "semaforo": "VERDE" if cumpl_w42 >= 100 else ("AMARILLO" if cumpl_w42 >= 80 else "ROJO"), "debe": debe_w42_int},
            "MTD": {"label": "Mes Acumulado Octubre (MTD)", "target": tgt_mtd_int, "venta_uds": vta_mtd_int, "venta_pesos": int(round(vp_mtd, 0)), "cumplimiento": cumpl_mtd, "semaforo": "VERDE" if cumpl_mtd >= 100 else ("AMARILLO" if cumpl_mtd >= 80 else "ROJO"), "debe": debe_mtd_int}
        },
        "segmentos": detalle_segmentos,
        "total_stock": int(total_stock_acumulado),
        "stock_skus": skus_totales
    })

# 6. SUPERVISORES
supervisores_salida = []

for _, s_row in supervisores_dir.iterrows():
    s_ced = str(s_row['CEDULA']).strip()
    s_nom = str(s_row['NOMBRE']).strip()
    s_ciu = str(s_row['CIUDAD']).strip()
    
    asesores_cargo = [c for c in consultores_salida if s_nom in c['supervisor']]
    if not asesores_cargo:
        continue
    
    tiendas_asignadas_sup = list(set([t for c in asesores_cargo for t in c['tienda'].split(' + ')]))
    
    informe_tiendas_sup = []
    for t_nom in tiendas_asignadas_sup:
        if t_nom in datos_tiendas:
            informe_tiendas_sup.append(datos_tiendas[t_nom])
            
    t_mtd_sup = int(sum([t['target_mtd'] for t in informe_tiendas_sup]))
    v_mtd_sup = int(sum([t['venta_mtd'] for t in informe_tiendas_sup]))
    debe_mtd_sup = int(max(0, t_mtd_sup - v_mtd_sup))
    cumpl_sup = round((v_mtd_sup / t_mtd_sup) * 100, 1) if t_mtd_sup > 0 else 0.0
    
    ihs_valores = [t['ihs_w42'] for t in informe_tiendas_sup if t['ihs_w42'] > 0]
    prom_ihs = round(sum(ihs_valores) / len(ihs_valores), 1) if ihs_valores else 0.0
    
    supervisores_salida.append({
        "cedula": s_ced,
        "nombre": s_nom,
        "ciudad_base": s_ciu,
        "total_asesores": int(len(asesores_cargo)),
        "total_tiendas": int(len(informe_tiendas_sup)),
        "target_mtd": t_mtd_sup,
        "venta_mtd": v_mtd_sup,
        "debe_mtd": debe_mtd_sup,
        "cumplimiento_mtd": float(cumpl_sup),
        "promedio_ihs_w42": float(prom_ihs),
        "semaforo": "VERDE" if cumpl_sup >= 100 else ("AMARILLO" if cumpl_sup >= 80 else "ROJO"),
        "tiendas_detalle": informe_tiendas_sup,
        "asesores": asesores_cargo
    })

data_final = {
    "fecha_corte": pd.Timestamp.now().strftime("%d/%m/%Y %H:%M"),
    "consultores": consultores_salida,
    "supervisores": supervisores_salida
}

# Serializador seguro contra tipos int64/float64 de NumPy/Pandas
def convertir_tipos(obj):
    if hasattr(obj, 'item'):
        return obj.item()
    return str(obj)

with open("data_app.json", "w", encoding="utf-8") as f:
    json.dump(data_final, f, ensure_ascii=False, indent=2, default=convertir_tipos)

print("==================================================")
print(f"-> ¡ÉXITO! 'data_app.json' generado sin errores.")
print(f"-> Consultores reales procesados: {len(consultores_salida)}")
print(f"-> Supervisores configurados: {len(supervisores_salida)}")
print("==================================================")