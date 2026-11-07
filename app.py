import os
import sqlite3
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import json

app = Flask(__name__)
app.secret_key = 'repuestos_cucuta_secret_2026'
DB_PATH = 'repuestos_cucuta.db'
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # CONFIGURACIÓN GENERAL DEL NEGOCIO
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS configuracion (
            clave TEXT PRIMARY KEY,
            valor TEXT
        )
    ''')
    
    config_defaults = {
        'nombre_negocio': 'AUTOREPUESTOS MERP',
        'nit_negocio': '88208896-7',
        'regimen_negocio': 'RÉGIMEN COMÚN',
        'direccion_negocio': 'CALLE 8 #6-27 BARRIO CENTRO VILLA DEL ROSARIO',
        'telefono_negocio': '3213942887',
        'ciudad_negocio': 'VILLA DEL ROSARIO - CÚCUTA, COLOMBIA',
        'encargada_nombre': 'Dayana Rujano',
        'valor_uvt': '52374',
        'presupuesto_fijos': '2500000',
        'presupuesto_variables': '1200000'
    }
    for k, v in config_defaults.items():
        cursor.execute("INSERT OR IGNORE INTO configuracion (clave, valor) VALUES (?, ?)", (k, v))

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS inventario (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_interno TEXT,
            descripcion TEXT,
            categoria TEXT,
            codigo_barras TEXT,
            codigo_opcional TEXT,
            serial TEXT,
            nro_factura_prov TEXT,
            nro_control TEXT,
            costo REAL,
            precio_venta REAL,
            stock REAL,
            proveedor TEXT,
            dias_credito INTEGER,
            garantia TEXT,
            imagen TEXT,
            imagen2 TEXT,
            imagen3 TEXT,
            fecha_registro TEXT
        )
    ''')
    
    for col, definition in [
        ('codigo_opcional', 'TEXT'),
        ('nro_control', 'TEXT'),
        ('serial', 'TEXT'),
        ('nro_factura_prov', 'TEXT'),
        ('stock_en_reclamo', 'REAL DEFAULT 0'),
        ('imagen2', 'TEXT'),
        ('imagen3', 'TEXT')
    ]:
        try:
            cursor.execute(f"ALTER TABLE inventario ADD COLUMN {col} {definition}")
            conn.commit()
        except sqlite3.OperationalError:
            pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cuentas_por_pagar (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nro_factura TEXT UNIQUE,
            nro_control TEXT,
            proveedor TEXT,
            fecha_emision TEXT,
            dias_plazo INTEGER,
            fecha_vencimiento TEXT,
            total_factura REAL,
            monto_abonado REAL,
            monto_restante REAL,
            estado_pago TEXT,
            productos_resumen TEXT,
            aplica_iva TEXT,
            monto_iva REAL,
            monto_retencion REAL,
            tiene_nota_credito TEXT DEFAULT 'No',
            iva_declarado TEXT DEFAULT 'No',
            proveedor_declara_renta TEXT DEFAULT 'Si',
            monto_retefuente REAL DEFAULT 0,
            base_imponible REAL DEFAULT 0
        )
    ''')
    for col, definition in [
        ('iva_declarado', 'TEXT DEFAULT "No"'),
        ('proveedor_declara_renta', 'TEXT DEFAULT "Si"'),
        ('monto_retefuente', 'REAL DEFAULT 0'),
        ('base_imponible', 'REAL DEFAULT 0')
    ]:
        try:
            cursor.execute(f"ALTER TABLE cuentas_por_pagar ADD COLUMN {col} {definition}")
            conn.commit()
        except sqlite3.OperationalError:
            pass
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS retenciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nro_factura TEXT,
            nro_control TEXT,
            proveedor TEXT,
            fecha TEXT,
            base_imponible REAL,
            monto_iva REAL,
            monto_retencion REAL,
            tipo_retencion TEXT DEFAULT 'ReteFuente 2.5%'
        )
    ''')
    try:
        cursor.execute("ALTER TABLE retenciones ADD COLUMN tipo_retencion TEXT DEFAULT 'ReteFuente 2.5%'")
        conn.commit()
    except sqlite3.OperationalError:
        pass
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            documento TEXT UNIQUE,
            nombre TEXT,
            telefono TEXT,
            direccion TEXT,
            compras_count INTEGER DEFAULT 1
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS facturas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_doc TEXT,
            cliente_nombre TEXT,
            cliente_telefono TEXT,
            cliente_direccion TEXT,
            fecha TEXT,
            total REAL,
            tipo_pago TEXT,
            monto_abonado REAL,
            monto_restante REAL,
            dias_plazo INTEGER,
            estado_pago TEXT,
            fecha_vencimiento TEXT,
            tiene_nota_credito TEXT DEFAULT 'No',
            monto_iva REAL DEFAULT 0,
            iva_declarado TEXT DEFAULT 'No',
            es_agente_retencion TEXT DEFAULT 'No',
            monto_retefuente REAL DEFAULT 0,
            base_imponible REAL DEFAULT 0,
            tipo_venta TEXT DEFAULT 'Repuestos'
        )
    ''')
    for col, definition in [
        ('monto_iva', 'REAL DEFAULT 0'),
        ('iva_declarado', 'TEXT DEFAULT "No"'),
        ('es_agente_retencion', 'TEXT DEFAULT "No"'),
        ('monto_retefuente', 'REAL DEFAULT 0'),
        ('base_imponible', 'REAL DEFAULT 0'),
        ('tipo_venta', 'TEXT DEFAULT "Repuestos"')
    ]:
        try:
            cursor.execute(f"ALTER TABLE facturas ADD COLUMN {col} {definition}")
            conn.commit()
        except sqlite3.OperationalError:
            pass
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS factura_detalles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            factura_id INTEGER,
            codigo_interno TEXT,
            descripcion TEXT,
            cantidad REAL,
            precio_unitario REAL,
            subtotal REAL,
            es_servicio INTEGER DEFAULT 0
        )
    ''')
    try:
        cursor.execute("ALTER TABLE factura_detalles ADD COLUMN es_servicio INTEGER DEFAULT 0")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS devoluciones_espera (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo_flujo TEXT,
            referencia_factura TEXT,
            tercero TEXT,
            codigo_interno TEXT,
            descripcion TEXT,
            cantidad REAL,
            opcion_reclamo TEXT,
            monto_afectado REAL,
            motivo TEXT,
            estado TEXT DEFAULT 'En Espera',
            fecha TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE devoluciones_espera ADD COLUMN motivo TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS caja_movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            fecha_dia TEXT,
            tipo TEXT,
            referencia TEXT,
            cliente_o_concepto TEXT,
            metodo_pago TEXT,
            monto REAL,
            monto_usd REAL DEFAULT 0,
            tasa_cambio REAL DEFAULT 0,
            cerrado TEXT DEFAULT 'No',
            cierre_id INTEGER DEFAULT 0,
            origen_ingreso TEXT DEFAULT 'Repuestos',
            tipo_gasto TEXT DEFAULT 'Variable'
        )
    ''')
    for col, definition in [
        ('monto_usd', 'REAL DEFAULT 0'),
        ('tasa_cambio', 'REAL DEFAULT 0'),
        ('origen_ingreso', 'TEXT DEFAULT "Repuestos"'),
        ('tipo_gasto', 'TEXT DEFAULT "Variable"')
    ]:
        try:
            cursor.execute(f"ALTER TABLE caja_movimientos ADD COLUMN {col} {definition}")
            conn.commit()
        except sqlite3.OperationalError:
            pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS caja_cierres (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_cierre TEXT,
            total_efectivo REAL,
            total_nequi REAL,
            total_bancolombia REAL,
            total_daviplata REAL,
            total_transferencia REAL,
            total_zelle REAL,
            total_zelle_usd REAL DEFAULT 0,
            total_dolar_efectivo REAL DEFAULT 0,
            total_dolar_efectivo_usd REAL DEFAULT 0,
            total_egresos REAL,
            gran_total REAL,
            observaciones TEXT,
            total_servicios REAL DEFAULT 0,
            total_repuestos REAL DEFAULT 0
        )
    ''')
    for col, definition in [
        ('total_zelle_usd', 'REAL DEFAULT 0'),
        ('total_dolar_efectivo', 'REAL DEFAULT 0'),
        ('total_dolar_efectivo_usd', 'REAL DEFAULT 0'),
        ('total_servicios', 'REAL DEFAULT 0'),
        ('total_repuestos', 'REAL DEFAULT 0')
    ]:
        try:
            cursor.execute(f"ALTER TABLE caja_cierres ADD COLUMN {col} {definition}")
            conn.commit()
        except sqlite3.OperationalError:
            pass

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tareas_pendientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titulo TEXT NOT NULL,
            descripcion TEXT,
            categoria TEXT,
            fecha_vencimiento TEXT,
            frecuencia TEXT,
            estado TEXT DEFAULT 'Pendiente',
            fecha_creacion TEXT
        )
    ''')
    
    conn.commit()
    conn.close()


init_db()


def get_config():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT clave, valor FROM configuracion")
    rows = cursor.fetchall()
    conn.close()
    return {r[0]: r[1] for r in rows}


def numero_a_letras_cop(numero):
    try:
        n = int(round(numero))
        return f"{n:,.0f} PESOS M/CTE".replace(',', '.')
    except:
        return f"{numero} PESOS M/CTE"


@app.route('/')
def index():
    busqueda = request.args.get('q', '').strip()
    cat_filtro = request.args.get('categoria', '').strip()
    hoy = datetime.now()
    config = get_config()
    
    valor_uvt = float(config.get('valor_uvt', 52374))
    tope_uvt_pesos = valor_uvt * 10.0  # 10 UVT base mínima
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    query = """
        SELECT id, codigo_interno, descripcion, categoria, codigo_barras, 
               codigo_opcional, serial, nro_factura_prov, nro_control, 
               costo, precio_venta, stock, proveedor, dias_credito, 
               garantia, imagen, fecha_registro, COALESCE(stock_en_reclamo, 0),
               COALESCE(imagen2, ''), COALESCE(imagen3, '')
        FROM inventario WHERE 1=1
    """
    params = []
    if busqueda:
        query += " AND (codigo_interno LIKE ? OR descripcion LIKE ? OR codigo_barras LIKE ? OR codigo_opcional LIKE ? OR serial LIKE ?)"
        like_term = f"%{busqueda}%"
        params.extend([like_term, like_term, like_term, like_term, like_term])
    if cat_filtro and cat_filtro != 'Todos':
        query += " AND categoria = ?"
        params.append(cat_filtro)
    query += " ORDER BY id DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    
    inventario = []
    for r in rows:
        stk_total = float(r[11] or 0.0)
        stk_reclamo = float(r[17] or 0.0)
        stk_disponible = max(0.0, stk_total - stk_reclamo)
        img1 = r[15] or ''
        img2 = r[18] or ''
        img3 = r[19] or ''
        fotos_list = [f for f in [img1, img2, img3] if f]

        inventario.append({
            'id': r[0], 'codigo_interno': r[1], 'descripcion': r[2],
            'categoria': r[3], 'codigo_barras': r[4], 'codigo_opcional': r[5] or 'N/A',
            'serial': r[6] or 'N/A', 'nro_factura_prov': r[7] or 'S/N', 'nro_control': r[8] or 'S/N',
            'costo': float(r[9]) if r[9] is not None else 0.0, 
            'precio_venta': float(r[10]) if r[10] is not None else 0.0, 
            'stock': stk_total,
            'stock_total': stk_total,
            'stock_reclamo': stk_reclamo,
            'stock_disponible': stk_disponible,
            'proveedor': r[12] or 'N/I', 'dias_credito': int(r[13]) if r[13] is not None else 0, 
            'garantia': r[14] or 'Sin garantía', 
            'imagen': img1, 'imagen2': img2, 'imagen3': img3,
            'fotos': fotos_list, 'fotos_count': len(fotos_list),
            'fecha_registro': r[16] or datetime.now().strftime('%d/%m/%Y')
        })

    cursor.execute("SELECT SUM(costo * (stock - COALESCE(stock_en_reclamo, 0))) FROM inventario")
    inversion_stock = float(cursor.fetchone()[0] or 0.0)

    cursor.execute("SELECT SUM(total), COUNT(id) FROM facturas")
    res_ventas = cursor.fetchone()
    ventas_totales = float(res_ventas[0] or 0.0)
    total_facturas_count = int(res_ventas[1] or 0)

    cursor.execute("SELECT SUM(monto_restante), COUNT(id) FROM facturas WHERE estado_pago != 'Pagado'")
    res_cobros = cursor.fetchone()
    dinero_por_cobrar = float(res_cobros[0] or 0.0)
    facturas_pendientes_count = int(res_cobros[1] or 0)

    cursor.execute("SELECT SUM(monto_restante), COUNT(id) FROM cuentas_por_pagar WHERE estado_pago != 'Pagado'")
    res_cxp = cursor.fetchone()
    total_por_pagar = float(res_cxp[0] or 0.0)
    cxp_pendientes_count = int(res_cxp[1] or 0)

    # IMPUESTOS: IVA Y RETEFUENTE
    cursor.execute("SELECT SUM(monto_iva) FROM cuentas_por_pagar WHERE aplica_iva = 'Si' AND COALESCE(iva_declarado, 'No') = 'No'")
    iva_compras_descontable = float(cursor.fetchone()[0] or 0.0)

    cursor.execute("SELECT SUM(monto_iva) FROM facturas WHERE COALESCE(iva_declarado, 'No') = 'No'")
    iva_ventas_generado = float(cursor.fetchone()[0] or 0.0)
    iva_dian_balance = iva_ventas_generado - iva_compras_descontable

    cursor.execute("SELECT SUM(monto_retefuente) FROM cuentas_por_pagar WHERE COALESCE(iva_declarado, 'No') = 'No'")
    retefuente_compras_por_pagar = float(cursor.fetchone()[0] or 0.0)

    cursor.execute("SELECT SUM(monto_retefuente) FROM facturas WHERE COALESCE(iva_declarado, 'No') = 'No'")
    retefuente_ventas_a_favor = float(cursor.fetchone()[0] or 0.0)

    cursor.execute("SELECT SUM(monto_retencion) FROM retenciones")
    total_retenciones = float(cursor.fetchone()[0] or 0.0)

    # Ranking Mejores Productos
    cursor.execute('''
        SELECT fd.codigo_interno, fd.descripcion, SUM(fd.cantidad) as total_vendido
        FROM factura_detalles fd
        WHERE COALESCE(fd.es_servicio, 0) = 0
        GROUP BY fd.codigo_interno
        ORDER BY total_vendido DESC
    ''')
    ranking_rows = cursor.fetchall()
    mejores_productos_data = []
    for rk in ranking_rows:
        cod_int = rk[0]
        desc = rk[1]
        tot_vendido = float(rk[2] or 0.0)
        cursor.execute("SELECT codigo_barras, serial FROM inventario WHERE codigo_interno = ? LIMIT 1", (cod_int,))
        inv_info = cursor.fetchone()
        cod_barras = inv_info[0] if inv_info and inv_info[0] else 'N/A'
        serial_prod = inv_info[1] if inv_info and inv_info[1] else 'N/A'
        
        cursor.execute('''
            SELECT f.cliente_doc, f.cliente_nombre, SUM(fd.cantidad) as cant_cliente
            FROM factura_detalles fd
            JOIN facturas f ON fd.factura_id = f.id
            WHERE fd.codigo_interno = ?
            GROUP BY f.cliente_doc, f.cliente_nombre
            ORDER BY cant_cliente DESC
        ''', (cod_int,))
        clientes_prod_rows = cursor.fetchall()
        clientes_por_producto = [{'documento': cp[0], 'nombre': cp[1], 'cantidad': float(cp[2] or 0.0)} for cp in clientes_prod_rows]
        
        mejores_productos_data.append({
            'codigo_interno': cod_int, 'descripcion': desc,
            'codigo_barras': cod_barras, 'serial': serial_prod,
            'total_vendido': tot_vendido, 'clientes_compradores': clientes_por_producto
        })

    cursor.execute('''
        SELECT c.documento, c.nombre, c.telefono, c.direccion, c.compras_count,
               COALESCE(SUM(f.total), 0) as total_gastado
        FROM clientes c
        LEFT JOIN facturas f ON c.documento = f.cliente_doc
        GROUP BY c.documento
        ORDER BY total_gastado DESC
    ''')
    clientes_rows = cursor.fetchall()
    clientes_data = []
    for cr in clientes_rows:
        doc_cli = cr[0]
        cursor.execute("SELECT f.id, f.fecha, fd.descripcion, fd.cantidad, fd.subtotal FROM facturas f JOIN factura_detalles fd ON f.id = fd.factura_id WHERE f.cliente_doc = ?", (doc_cli,))
        compras_detalle = cursor.fetchall()
        detalles_lista = [{'factura_id': cd[0], 'fecha': cd[1], 'repuesto': cd[2], 'cantidad': float(cd[3] or 0), 'subtotal': float(cd[4] or 0)} for cd in compras_detalle]
        clientes_data.append({
            'documento': doc_cli, 'nombre': cr[1], 'telefono': cr[2] or 'N/A',
            'direccion': cr[3] or 'N/A', 'facturas_count': int(cr[4] or 0), 
            'total_gastado': float(cr[5] or 0.0), 'compras_detalle': detalles_lista
        })

    cursor.execute("SELECT id, cliente_doc, cliente_nombre, fecha, total, tipo_pago, monto_abonado, monto_restante, dias_plazo, estado_pago, fecha_vencimiento, COALESCE(tiene_nota_credito, 'No') FROM facturas ORDER BY id DESC")
    facturas_rows = cursor.fetchall()
    facturas_data = []
    for fr in facturas_rows:
        f_id, f_doc, f_nom, f_fecha, f_tot, f_tipo, f_abonado, f_restante, f_dias, f_estado, f_venc, f_nc = fr
        dias_restantes_texto = ""
        if f_estado != 'Pagado' and f_venc:
            try:
                dt_venc = datetime.strptime(f_venc, '%d/%m/%Y')
                delta = (dt_venc - hoy).days
                if delta > 0: dias_restantes_texto = f"Vence en {delta} días"
                elif delta == 0: dias_restantes_texto = "Vence HOY"
                else: dias_restantes_texto = f"Vencido hace {abs(delta)} días"
            except: dias_restantes_texto = "Plazo vigente"
        elif f_estado == 'Pagado':
            dias_restantes_texto = "Pagado completo"
        facturas_data.append({
            'id': f_id, 'doc': f_doc, 'nombre': f_nom, 'fecha': f_fecha, 
            'total': float(f_tot or 0), 'tipo_pago': f_tipo, 
            'abonado': float(f_abonado or 0), 'restante': float(f_restante or 0),
            'estado': f_estado, 'vencimiento_texto': dias_restantes_texto,
            'tiene_nota_credito': f_nc
        })

    cursor.execute("SELECT id, nro_factura, nro_control, proveedor, fecha_emision, dias_plazo, fecha_vencimiento, total_factura, monto_abonado, monto_restante, estado_pago, productos_resumen, aplica_iva, monto_iva, monto_retencion, COALESCE(tiene_nota_credito, 'No') FROM cuentas_por_pagar ORDER BY id DESC")
    cxp_rows = cursor.fetchall()
    cuentas_por_pagar_data = []
    for cx in cxp_rows:
        cx_id, cx_nro, cx_ctrl, cx_prov, cx_fec, cx_dias, cx_venc, cx_tot, cx_abon, cx_rest, cx_est, cx_prod, cx_iva_si, cx_miva, cx_mret, cx_nc = cx
        venc_texto = ""
        if cx_est != 'Pagado' and cx_venc:
            try:
                dt_venc = datetime.strptime(cx_venc, '%d/%m/%Y')
                delta = (dt_venc - hoy).days
                if delta > 0: venc_texto = f"Faltan {delta} días"
                elif delta == 0: venc_texto = "Vence HOY"
                else: venc_texto = f"Vencido hace {abs(delta)} días"
            except: venc_texto = "Vigente"
        else:
            venc_texto = "Pagado completo"
        cuentas_por_pagar_data.append({
            'id': cx_id, 'nro_factura': cx_nro, 'nro_control': cx_ctrl or 'S/N', 'proveedor': cx_prov, 'fecha_emision': cx_fec,
            'dias_plazo': cx_dias, 'vencimiento_texto': venc_texto, 
            'total': float(cx_tot or 0), 'abonado': float(cx_abon or 0), 'restante': float(cx_rest or 0), 
            'estado': cx_est, 'productos': cx_prod or '', 'aplica_iva': cx_iva_si, 
            'monto_iva': float(cx_miva or 0), 'monto_retencion': float(cx_mret or 0), 'tiene_nota_credito': cx_nc
        })

    cursor.execute("SELECT id, nro_factura, nro_control, proveedor, fecha, base_imponible, monto_iva, monto_retencion FROM retenciones ORDER BY id DESC")
    retenciones_rows = cursor.fetchall()
    retenciones_data = []
    for ret in retenciones_rows:
        retenciones_data.append({
            'id': ret[0], 'nro_factura': ret[1], 'nro_control': ret[2], 'proveedor': ret[3],
            'fecha': ret[4], 'base_imponible': float(ret[5] or 0), 'monto_iva': float(ret[6] or 0), 'monto_retencion': float(ret[7] or 0)
        })

    # Pestaña 9: Pedidos
    cursor.execute("SELECT f.fecha, fd.codigo_interno, fd.cantidad FROM factura_detalles fd JOIN facturas f ON fd.factura_id = f.id WHERE COALESCE(fd.es_servicio, 0) = 0")
    ventas_historicas = cursor.fetchall()
    ventas_rotacion = {}
    for f_fecha_raw, cod_prod, cant_raw in ventas_historicas:
        if not cod_prod: continue
        dt_v = None
        for fmt in ('%d/%m/%Y %H:%M', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                dt_v = datetime.strptime(str(f_fecha_raw).strip()[:16], fmt)
                break
            except: continue
        cant_v = float(cant_raw or 0.0)
        if cod_prod not in ventas_rotacion: ventas_rotacion[cod_prod] = {30: 0.0, 60: 0.0, 90: 0.0, 120: 0.0}
        if dt_v:
            diff_dias = (hoy - dt_v).days
            if diff_dias <= 30: ventas_rotacion[cod_prod][30] += cant_v
            if diff_dias <= 60: ventas_rotacion[cod_prod][60] += cant_v
            if diff_dias <= 90: ventas_rotacion[cod_prod][90] += cant_v
            if diff_dias <= 120: ventas_rotacion[cod_prod][120] += cant_v
        else:
            ventas_rotacion[cod_prod][30] += cant_v
            ventas_rotacion[cod_prod][60] += cant_v
            ventas_rotacion[cod_prod][90] += cant_v
            ventas_rotacion[cod_prod][120] += cant_v

    cursor.execute("SELECT codigo_interno, descripcion, SUM(stock), SUM(COALESCE(stock_en_reclamo, 0)), MAX(costo), MAX(proveedor) FROM inventario GROUP BY codigo_interno ORDER BY codigo_interno ASC")
    inv_agrupado = cursor.fetchall()
    analisis_pedidos = []
    for r_cod, r_desc, r_stk, r_rec, r_cost, r_prov in inv_agrupado:
        stk_tot = float(r_stk or 0.0)
        stk_rec = float(r_rec or 0.0)
        stk_disp = max(0.0, stk_tot - stk_rec)
        v_rot = ventas_rotacion.get(r_cod, {30: 0.0, 60: 0.0, 90: 0.0, 120: 0.0})
        v30, v60, v90, v120 = v_rot[30], v_rot[60], v_rot[90], v_rot[120]
        demanda = v30 if v30 > 0 else (v60 / 2.0 if v60 > 0 else (v90 / 3.0 if v90 > 0 else (v120 / 4.0 if v120 > 0 else 0.0)))
        sugerido = int(round(demanda - stk_disp)) if demanda > stk_disp else 0
        analisis_pedidos.append({
            'codigo': r_cod, 'descripcion': r_desc, 'stock_total': stk_tot, 'stock_reclamo': stk_rec,
            'stock_disponible': stk_disp, 'v_30': v30, 'v_60': v60, 'v_90': v90, 'v_120': v120,
            'proveedor': r_prov or 'N/I', 'costo': float(r_cost or 0.0), 'sugerido_pedir': sugerido
        })

    # Pestaña 10: Reclamos
    cursor.execute("SELECT id, tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, opcion_reclamo, monto_afectado, estado, fecha, COALESCE(motivo, '') FROM devoluciones_espera ORDER BY id DESC")
    devoluciones_rows = cursor.fetchall()
    devoluciones_compras_data = []
    devoluciones_ventas_data = []
    for d in devoluciones_rows:
        dev_obj = {
            'id': d[0], 'tipo_flujo': d[1], 'referencia_factura': d[2], 'tercero': d[3],
            'codigo_interno': d[4], 'descripcion': d[5], 'cantidad': float(d[6] or 0),
            'opcion_reclamo': d[7], 'monto_afectado': float(d[8] or 0), 'estado': d[9],
            'fecha': d[10], 'motivo': d[11]
        }
        if d[1] == 'Compra': devoluciones_compras_data.append(dev_obj)
        else: devoluciones_ventas_data.append(dev_obj)

    # Pestaña 11: Arqueo de Caja + Separación de Servicios vs Repuestos + Presupuestos
    cursor.execute("SELECT id, fecha, tipo, referencia, cliente_o_concepto, metodo_pago, monto, COALESCE(monto_usd, 0), COALESCE(tasa_cambio, 0), COALESCE(origen_ingreso, 'Repuestos'), COALESCE(tipo_gasto, 'Variable') FROM caja_movimientos WHERE cerrado = 'No' ORDER BY id DESC")
    movimientos_caja_rows = cursor.fetchall()
    
    caja_efectivo_ing = 0.0
    caja_efectivo_egr = 0.0
    caja_nequi = 0.0
    caja_bancolombia = 0.0
    caja_daviplata = 0.0
    caja_transferencia = 0.0
    caja_zelle_cop = 0.0
    caja_zelle_usd = 0.0
    caja_dolar_efectivo_cop = 0.0
    caja_dolar_efectivo_usd = 0.0
    caja_ingresos_servicios = 0.0
    caja_ingresos_repuestos = 0.0
    gastos_fijos_ejecutados = 0.0
    gastos_variables_ejecutados = 0.0
    
    movimientos_caja_activa = []
    for m in movimientos_caja_rows:
        m_id, m_fec, m_tipo, m_ref, m_con, m_met, m_val, m_usd, m_tasa, m_orig, m_tgasto = m
        val_cop = float(m_val or 0.0)
        val_usd = float(m_usd or 0.0)
        movimientos_caja_activa.append({
            'id': m_id, 'fecha': m_fec, 'tipo': m_tipo, 'referencia': m_ref,
            'concepto': m_con, 'metodo': m_met, 'monto': val_cop, 'monto_usd': val_usd, 'tasa': float(m_tasa or 0)
        })
        
        if m_tipo == 'Egreso Gasto':
            if m_tgasto == 'Fijo': gastos_fijos_ejecutados += val_cop
            else: gastos_variables_ejecutados += val_cop
            
            if m_met == 'Efectivo': caja_efectivo_egr += val_cop
            elif m_met == 'Nequi': caja_nequi -= val_cop
            elif m_met == 'Bancolombia': caja_bancolombia -= val_cop
            elif m_met == 'Daviplata': caja_daviplata -= val_cop
            elif m_met == 'Transferencia': caja_transferencia -= val_cop
            elif m_met == 'Zelle ($)': 
                caja_zelle_cop -= val_cop
                caja_zelle_usd -= val_usd
            elif m_met == 'Efectivo Dólar ($)':
                caja_dolar_efectivo_cop -= val_cop
                caja_dolar_efectivo_usd -= val_usd
        else:
            if m_orig == 'Servicios': caja_ingresos_servicios += val_cop
            else: caja_ingresos_repuestos += val_cop
            
            if m_met == 'Efectivo': caja_efectivo_ing += val_cop
            elif m_met == 'Nequi': caja_nequi += val_cop
            elif m_met == 'Bancolombia': caja_bancolombia += val_cop
            elif m_met == 'Daviplata': caja_daviplata += val_cop
            elif m_met == 'Transferencia': caja_transferencia += val_cop
            elif m_met == 'Zelle ($)':
                caja_zelle_cop += val_cop
                caja_zelle_usd += val_usd
            elif m_met == 'Efectivo Dólar ($)':
                caja_dolar_efectivo_cop += val_cop
                caja_dolar_efectivo_usd += val_usd

    caja_efectivo_neto = max(0.0, caja_efectivo_ing - caja_efectivo_egr)
    caja_gran_total = caja_efectivo_neto + caja_nequi + caja_bancolombia + caja_daviplata + caja_transferencia + caja_zelle_cop + caja_dolar_efectivo_cop
    
    # Presupuesto de gastos
    presup_fijos_total = float(config.get('presupuesto_fijos', 2500000))
    presup_vars_total = float(config.get('presupuesto_variables', 1200000))
    presup_fijos_restante = max(0.0, presup_fijos_total - gastos_fijos_ejecutados)
    presup_vars_restante = max(0.0, presup_vars_total - gastos_variables_ejecutados)

    cursor.execute("SELECT id, fecha_cierre, total_efectivo, gran_total, total_egresos, observaciones, COALESCE(total_zelle_usd, 0), COALESCE(total_dolar_efectivo_usd, 0), COALESCE(total_servicios, 0), COALESCE(total_repuestos, 0) FROM caja_cierres ORDER BY id DESC LIMIT 15")
    cierres_historico_rows = cursor.fetchall()
    cierres_historico = [{
        'id': ch[0], 'fecha': ch[1], 'efectivo': float(ch[2] or 0),
        'gran_total': float(ch[3] or 0), 'egresos': float(ch[4] or 0), 'observaciones': ch[5],
        'zelle_usd': float(ch[6] or 0), 'dolar_efectivo_usd': float(ch[7] or 0),
        'servicios': float(ch[8] or 0), 'repuestos': float(ch[9] or 0)
    } for ch in cierres_historico_rows]

    # Pestaña 12: Tareas y Obligaciones
    cursor.execute("SELECT id, titulo, descripcion, categoria, fecha_vencimiento, frecuencia, estado, fecha_creacion FROM tareas_pendientes ORDER BY estado ASC, id DESC")
    tareas_rows = cursor.fetchall()
    tareas_pendientes_data = []
    tareas_realizadas_data = []
    
    for t in tareas_rows:
        t_id, t_tit, t_desc, t_cat, t_venc, t_frec, t_est, t_crea = t
        estado_alerta = "Programada"
        dias_delta = 999
        if t_venc:
            try:
                if '-' in t_venc: dt_t = datetime.strptime(t_venc, '%Y-%m-%d')
                else: dt_t = datetime.strptime(t_venc, '%d/%m/%Y')
                t_venc_fmt = dt_t.strftime('%d/%m/%Y')
                delta = (dt_t.date() - hoy.date()).days
                dias_delta = delta
                if delta < 0: estado_alerta = f"⚠️ Vencida hace {abs(delta)} días"
                elif delta == 0: estado_alerta = "🔔 Vence HOY"
                else: estado_alerta = f"📅 Faltan {delta} días"
            except:
                t_venc_fmt = t_venc
                estado_alerta = "Fecha registrada"
        else:
            t_venc_fmt = "Sin fecha límite"

        t_obj = {
            'id': t_id, 'titulo': t_tit, 'descripcion': t_desc or '',
            'categoria': t_cat or 'General', 'fecha_venc': t_venc_fmt,
            'frecuencia': t_frec or 'Única puntual', 'estado': t_est,
            'alerta': estado_alerta, 'dias_delta': dias_delta
        }
        if t_est == 'Pendiente': tareas_pendientes_data.append(t_obj)
        else: tareas_realizadas_data.append(t_obj)

    tareas_pendientes_data.sort(key=lambda x: x['dias_delta'])

    conn.close()
    current_date = datetime.now().strftime('%d/%m/%Y')
    
    return render_template('index.html', inventario=inventario, current_date=current_date, 
                           busqueda=busqueda, cat_filtro=cat_filtro, config=config,
                           inversion_stock=inversion_stock, ventas_totales=ventas_totales,
                           total_facturas_count=total_facturas_count, 
                           dinero_por_cobrar=dinero_por_cobrar, facturas_pendientes_count=facturas_pendientes_count,
                           total_por_pagar=total_por_pagar, cxp_pendientes_count=cxp_pendientes_count,
                           total_retenciones=total_retenciones,
                           clientes_data=clientes_data, facturas_data=facturas_data,
                           cuentas_por_pagar_data=cuentas_por_pagar_data, retenciones_data=retenciones_data,
                           mejores_productos_data=mejores_productos_data,
                           analisis_pedidos=analisis_pedidos,
                           devoluciones_compras_data=devoluciones_compras_data,
                           devoluciones_ventas_data=devoluciones_ventas_data,
                           caja_efectivo_neto=caja_efectivo_neto,
                           caja_efectivo_egr=caja_efectivo_egr,
                           caja_nequi=caja_nequi,
                           caja_bancolombia=caja_bancolombia,
                           caja_daviplata=caja_daviplata,
                           caja_transferencia=caja_transferencia,
                           caja_zelle_cop=caja_zelle_cop,
                           caja_zelle_usd=caja_zelle_usd,
                           caja_dolar_efectivo_cop=caja_dolar_efectivo_cop,
                           caja_dolar_efectivo_usd=caja_dolar_efectivo_usd,
                           caja_gran_total=caja_gran_total,
                           caja_ingresos_servicios=caja_ingresos_servicios,
                           caja_ingresos_repuestos=caja_ingresos_repuestos,
                           gastos_fijos_ejecutados=gastos_fijos_ejecutados,
                           gastos_variables_ejecutados=gastos_variables_ejecutados,
                           presup_fijos_restante=presup_fijos_restante,
                           presup_vars_restante=presup_vars_restante,
                           movimientos_caja_activa=movimientos_caja_activa,
                           cierres_historico=cierres_historico,
                           tareas_pendientes_data=tareas_pendientes_data,
                           tareas_realizadas_data=tareas_realizadas_data,
                           iva_compras_descontable=iva_compras_descontable,
                           iva_ventas_generado=iva_ventas_generado,
                           iva_dian_balance=iva_dian_balance,
                           retefuente_compras_por_pagar=retefuente_compras_por_pagar,
                           retefuente_ventas_a_favor=retefuente_ventas_a_favor,
                           tope_uvt_pesos=tope_uvt_pesos)


# ACTUALIZAR VALOR UVT Y PRESUPUESTOS DE GASTOS
@app.route('/actualizar_configuracion_uvt', methods=['POST'])
def actualizar_configuracion_uvt():
    val_uvt = request.form.get('valor_uvt', '52374').strip()
    p_fijos = request.form.get('presupuesto_fijos', '2500000').strip()
    p_vars = request.form.get('presupuesto_variables', '1200000').strip()
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES ('valor_uvt', ?)", (val_uvt,))
    cursor.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES ('presupuesto_fijos', ?)", (p_fijos,))
    cursor.execute("INSERT OR REPLACE INTO configuracion (clave, valor) VALUES ('presupuesto_variables', ?)", (p_vars,))
    conn.commit()
    conn.close()
    flash('¡Parámetros de UVT y presupuestos de gastos actualizados con éxito!', 'success')
    return redirect(url_for('index', tab='pestaña12'))


# REINICIAR / LIQUIDAR PERÍODO DE IVA ANTE LA DIAN
@app.route('/liquidar_periodo_iva')
def liquidar_periodo_iva():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE cuentas_por_pagar SET iva_declarado = 'Si' WHERE COALESCE(iva_declarado, 'No') = 'No'")
    cursor.execute("UPDATE facturas SET iva_declarado = 'Si' WHERE COALESCE(iva_declarado, 'No') = 'No'")
    conn.commit()
    conn.close()
    flash('¡Período de IVA liquidado con éxito! Los contadores de IVA Generado, Descontable y ReteFuente se han reiniciado a $ 0.00 sin borrar tus facturas.', 'success')
    return redirect(url_for('index', tab='pestaña12'))


# RUTAS DE TAREAS Y OBLIGACIONES
@app.route('/registrar_tarea', methods=['POST'])
def registrar_tarea():
    titulo = request.form.get('titulo_tarea', '').strip()
    descripcion = request.form.get('descripcion_tarea', '').strip()
    categoria = request.form.get('categoria_tarea', '🏛️ Impuestos / DIAN').strip()
    fecha_venc = request.form.get('fecha_vencimiento', '').strip()
    frecuencia = request.form.get('frecuencia_tarea', 'Fecha única puntual').strip()
    
    if not titulo:
        flash('El título o concepto de la tarea es obligatorio.', 'error')
        return redirect(url_for('index', tab='pestaña12'))
        
    fecha_crea = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO tareas_pendientes (titulo, descripcion, categoria, fecha_vencimiento, frecuencia, estado, fecha_creacion)
        VALUES (?, ?, ?, ?, ?, 'Pendiente', ?)
    ''', (titulo, descripcion, categoria, fecha_venc, frecuencia, fecha_crea))
    conn.commit()
    conn.close()
    
    flash(f'¡Tarea / Obligación "{titulo}" agregada a tu agenda!', 'success')
    return redirect(url_for('index', tab='pestaña12'))


@app.route('/marcar_tarea_realizada/<int:tarea_id>')
def marcar_tarea_realizada(tarea_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE tareas_pendientes SET estado = 'Realizada' WHERE id = ?", (tarea_id,))
    conn.commit()
    conn.close()
    flash('¡Tarea marcada como Realizada con éxito!', 'success')
    return redirect(url_for('index', tab='pestaña12'))


@app.route('/eliminar_tarea/<int:tarea_id>')
def eliminar_tarea(tarea_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM tareas_pendientes WHERE id = ?", (tarea_id,))
    conn.commit()
    conn.close()
    flash('Tarea eliminada de la agenda.', 'success')
    return redirect(url_for('index', tab='pestaña12'))


# GESTIÓN DE HASTA 3 FOTOS POR REPUESTO
@app.route('/subir_fotos_repuesto/<int:id>', methods=['POST'])
def subir_fotos_repuesto(id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT imagen, imagen2, imagen3 FROM inventario WHERE id = ?", (id,))
    row = cursor.fetchone()
    img1, img2, img3 = (row[0] if row else None), (row[1] if row else None), (row[2] if row else None)

    f1 = request.files.get('foto1')
    f2 = request.files.get('foto2')
    f3 = request.files.get('foto3')
    ts = datetime.now().strftime('%Y%m%d%H%M%S')

    if f1 and f1.filename:
        fn1 = secure_filename(f"rep_{id}_f1_{ts}_{f1.filename}")
        f1.save(os.path.join(app.config['UPLOAD_FOLDER'], fn1))
        img1 = fn1

    if f2 and f2.filename:
        fn2 = secure_filename(f"rep_{id}_f2_{ts}_{f2.filename}")
        f2.save(os.path.join(app.config['UPLOAD_FOLDER'], fn2))
        img2 = fn2

    if f3 and f3.filename:
        fn3 = secure_filename(f"rep_{id}_f3_{ts}_{f3.filename}")
        f3.save(os.path.join(app.config['UPLOAD_FOLDER'], fn3))
        img3 = fn3

    if request.form.get('eliminar_foto1') == '1': img1 = None
    if request.form.get('eliminar_foto2') == '1': img2 = None
    if request.form.get('eliminar_foto3') == '1': img3 = None

    cursor.execute("UPDATE inventario SET imagen = ?, imagen2 = ?, imagen3 = ? WHERE id = ?", (img1, img2, img3, id))
    conn.commit()
    conn.close()
    flash('¡Fotos del repuesto actualizadas con éxito!', 'success')
    return redirect(url_for('index', tab='pestaña2'))


# DESCARGAR COPIA DE SEGURIDAD
@app.route('/descargar_copia_seguridad')
def descargar_copia_seguridad():
    if not os.path.exists(DB_PATH):
        flash('No se encontró el archivo de base de datos para respaldar.', 'error')
        return redirect(url_for('index'))
        
    fecha_hoy = datetime.now().strftime('%Y_%m_%d_%H%M')
    nombre_archivo = f"Backup_AutoRepuestosMERP_{fecha_hoy}.db"
    
    buffer = BytesIO()
    with open(DB_PATH, 'rb') as f:
        buffer.write(f.read())
    buffer.seek(0)
    
    flash('¡Copia de seguridad generada y descargada exitosamente!', 'success')
    return send_file(buffer, as_attachment=True, download_name=nombre_archivo, mimetype='application/x-sqlite3')


# REGISTRAR GASTO DE CAJA CHICA CON CLASIFICACIÓN (FIJO / VARIABLE)
@app.route('/registrar_gasto_caja', methods=['POST'])
def registrar_gasto_caja():
    concepto = request.form.get('concepto_gasto', '').strip()
    metodo = request.form.get('metodo_gasto', 'Efectivo').strip()
    tipo_gasto = request.form.get('tipo_gasto', 'Variable').strip()
    try: monto = float(request.form.get('monto_gasto', 0))
    except: monto = 0.0
    
    if monto <= 0 or not concepto:
        flash('Por favor ingresa un concepto y monto válido de gasto.', 'error')
        return redirect(url_for('index', tab='pestaña11'))
        
    fecha_dt = datetime.now()
    fecha_str = fecha_dt.strftime('%d/%m/%Y %H:%M')
    fecha_dia = fecha_dt.strftime('%d/%m/%Y')
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO caja_movimientos (fecha, fecha_dia, tipo, referencia, cliente_o_concepto, metodo_pago, monto, cerrado, tipo_gasto)
        VALUES (?, ?, 'Egreso Gasto', 'Caja Menor / Gasto', ?, ?, ?, 'No', ?)
    ''', (fecha_str, fecha_dia, concepto, metodo, monto, tipo_gasto))
    conn.commit()
    conn.close()
    
    flash(f'¡Egreso ({tipo_gasto}) de $ {monto:,.2f} COP registrado en caja!', 'success')
    return redirect(url_for('index', tab='pestaña11'))


# FINALIZAR ARQUEO / CIERRE DE CAJA
@app.route('/cerrar_arqueo_caja', methods=['POST'])
def cerrar_arqueo_caja():
    observaciones = request.form.get('observaciones_cierre', 'Cierre diario regular').strip()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT tipo, metodo_pago, monto, COALESCE(monto_usd, 0), COALESCE(origen_ingreso, 'Repuestos') FROM caja_movimientos WHERE cerrado = 'No'")
    movs = cursor.fetchall()
    
    if not movs:
        conn.close()
        flash('No hay movimientos activos para realizar cierre de caja.', 'error')
        return redirect(url_for('index', tab='pestaña11'))
        
    tot_efectivo_ing = 0.0
    tot_egresos = 0.0
    tot_nequi = 0.0
    tot_bancolombia = 0.0
    tot_daviplata = 0.0
    tot_transferencia = 0.0
    tot_zelle_cop = 0.0
    tot_zelle_usd = 0.0
    tot_dolar_efectivo_cop = 0.0
    tot_dolar_efectivo_usd = 0.0
    tot_servicios = 0.0
    tot_repuestos = 0.0
    
    for tipo, met, m, m_usd, m_orig in movs:
        val_cop = float(m or 0.0)
        val_usd = float(m_usd or 0.0)
        if tipo == 'Egreso Gasto':
            tot_egresos += val_cop
            if met == 'Efectivo': tot_efectivo_ing -= val_cop
            elif met == 'Nequi': tot_nequi -= val_cop
            elif met == 'Bancolombia': tot_bancolombia -= val_cop
            elif met == 'Daviplata': tot_daviplata -= val_cop
            elif met == 'Transferencia': tot_transferencia -= val_cop
            elif met == 'Zelle ($)': 
                tot_zelle_cop -= val_cop
                tot_zelle_usd -= val_usd
            elif met == 'Efectivo Dólar ($)':
                tot_dolar_efectivo_cop -= val_cop
                tot_dolar_efectivo_usd -= val_usd
        else:
            if m_orig == 'Servicios': tot_servicios += val_cop
            else: tot_repuestos += val_cop
            
            if met == 'Efectivo': tot_efectivo_ing += val_cop
            elif met == 'Nequi': tot_nequi += val_cop
            elif met == 'Bancolombia': tot_bancolombia += val_cop
            elif met == 'Daviplata': tot_daviplata += val_cop
            elif met == 'Transferencia': tot_transferencia += val_cop
            elif met == 'Zelle ($)':
                tot_zelle_cop += val_cop
                tot_zelle_usd += val_usd
            elif met == 'Efectivo Dólar ($)':
                tot_dolar_efectivo_cop += val_cop
                tot_dolar_efectivo_usd += val_usd

    tot_efectivo_neto = max(0.0, tot_efectivo_ing)
    gran_total = tot_efectivo_neto + tot_nequi + tot_bancolombia + tot_daviplata + tot_transferencia + tot_zelle_cop + tot_dolar_efectivo_cop
    fecha_cierre_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    
    cursor.execute('''
        INSERT INTO caja_cierres (fecha_cierre, total_efectivo, total_nequi, total_bancolombia, total_daviplata, total_transferencia, total_zelle, total_zelle_usd, total_dolar_efectivo, total_dolar_efectivo_usd, total_egresos, gran_total, observaciones, total_servicios, total_repuestos)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (fecha_cierre_str, tot_efectivo_neto, tot_nequi, tot_bancolombia, tot_daviplata, tot_transferencia, tot_zelle_cop, tot_zelle_usd, tot_dolar_efectivo_cop, tot_dolar_efectivo_usd, tot_egresos, gran_total, observaciones, tot_servicios, tot_repuestos))
    
    cierre_id = cursor.lastrowid
    cursor.execute("UPDATE caja_movimientos SET cerrado = 'Si', cierre_id = ? WHERE cerrado = 'No'", (cierre_id,))
    conn.commit()
    conn.close()
    
    flash(f'¡Arqueo #{cierre_id} finalizado con éxito! Caja reiniciada a $ 0.00.', 'success')
    return redirect(url_for('index', cierre_descargar=cierre_id, tab='pestaña11'))


# COMPROBANTE PDF DE CIERRE DE CAJA
@app.route('/descargar_pdf_cierre/<int:cierre_id>')
def descargar_pdf_cierre(cierre_id):
    config = get_config()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, fecha_cierre, total_efectivo, total_nequi, total_bancolombia, total_daviplata, total_transferencia, total_zelle, total_egresos, gran_total, observaciones, COALESCE(total_zelle_usd, 0), COALESCE(total_dolar_efectivo, 0), COALESCE(total_dolar_efectivo_usd, 0), COALESCE(total_servicios, 0), COALESCE(total_repuestos, 0) FROM caja_cierres WHERE id = ?", (cierre_id,))
    cie = cursor.fetchone()
    
    cursor.execute("SELECT fecha, tipo, referencia, cliente_o_concepto, metodo_pago, monto, COALESCE(monto_usd, 0), COALESCE(tasa_cambio, 0) FROM caja_movimientos WHERE cierre_id = ? ORDER BY id ASC", (cierre_id,))
    movs = cursor.fetchall()
    conn.close()
    
    if not cie: return "Cierre no encontrado", 404
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#1E3A8A'), alignment=1)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#555555'), alignment=1)
    
    elements.append(Paragraph(config.get('nombre_negocio', 'AUTOREPUESTOS MERP'), title_style))
    elements.append(Paragraph(f"NIT: {config.get('nit_negocio')} | {config.get('direccion_negocio')} | Tel: {config.get('telefono_negocio')}", sub_style))
    elements.append(Paragraph(f"<b>COMPROBANTE OFICIAL DE CIERRE Y ARQUEO DE CAJA N° {cierre_id}</b>", title_style))
    elements.append(Paragraph(f"Fecha de Cierre: {cie[1]} | Encargada: {config.get('encargada_nombre')} | Notas: {cie[10]}", sub_style))
    elements.append(Spacer(1, 12))
    
    t_resumen_data = [
        ["Concepto / Canal", "Monto Verificado (COP)", "Monto en Divisa ($ USD)"],
        ["💵 Efectivo Gaveta (COP)", f"$ {float(cie[2]):,.2f}", "COP Físico"],
        ["💵 Dólares Físicos Gaveta (USD)", f"$ {float(cie[12]):,.2f}", f"$ {float(cie[13]):,.2f} USD"],
        ["🟣 Nequi", f"$ {float(cie[3]):,.2f}", "-"],
        ["🟡 Bancolombia", f"$ {float(cie[4]):,.2f}", "-"],
        ["🔴 Daviplata", f"$ {float(cie[5]):,.2f}", "-"],
        ["🏛️ Transferencias Generales", f"$ {float(cie[6]):,.2f}", "-"],
        ["💲 Zelle (USD)", f"$ {float(cie[7]):,.2f}", f"$ {float(cie[11]):,.2f} USD"],
        ["🔧 Ingresos por Servicios (Mano de obra)", f"$ {float(cie[14]):,.2f}", "-"],
        ["📦 Ingresos por Repuestos", f"$ {float(cie[15]):,.2f}", "-"],
        ["(-) Salidas / Egresos de Caja Menor", f"- $ {float(cie[8]):,.2f}", "-"],
        ["GRAN TOTAL CONCILIADO", f"$ {float(cie[9]):,.2f}", ""]
    ]
    t_res = Table(t_resumen_data, colWidths=[230, 150, 120])
    t_res.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#DCFCE7')),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_res)
    elements.append(Spacer(1, 15))
    
    elements.append(Paragraph("<b>Detalle de Movimientos del Turno:</b>", ParagraphStyle('Sub2', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#1E3A8A'))))
    elements.append(Spacer(1, 4))
    
    t_movs_data = [["Hora", "Tipo", "Referencia / Cliente", "Canal", "Monto COP", "Monto USD"]]
    for mv in movs:
        usd_txt = f"${float(mv[6]):,.2f} USD" if float(mv[6] or 0) > 0 else "-"
        t_movs_data.append([mv[0].split()[-1] if ' ' in mv[0] else mv[0], mv[1], f"{mv[2]} - {mv[3]}", mv[4], f"$ {float(mv[5]):,.2f}", usd_txt])
        
    t_mov = Table(t_movs_data, colWidths=[55, 80, 185, 75, 55, 50])
    t_mov.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#475569')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('FONTSIZE', (0,0), (-1,-1), 7.5),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_mov)
    elements.append(Spacer(1, 30))
    
    t_firmas = Table([["____________________________", "____________________________"],
                      [f"Firma: {config.get('encargada_nombre')}", "Firma Propietario / Auditor"]], colWidths=[250, 250])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    elements.append(t_firmas)
    
    doc.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Cierre_Caja_{cierre_id}.pdf", mimetype='application/pdf')


# EMISIÓN DE FACTURA (CON OPCIÓN DE RETEFUENTE Y SERVICIOS)
@app.route('/emitir_factura', methods=['POST'])
def emitir_factura():
    config = get_config()
    valor_uvt = float(config.get('valor_uvt', 52374))
    tope_uvt = valor_uvt * 10.0
    
    doc = request.form.get('cliente_doc', '').strip()
    nombre = request.form.get('cliente_nombre', '').strip()
    telefono = request.form.get('cliente_telefono', '').strip()
    direccion = request.form.get('cliente_direccion', '').strip()
    tipo_pago = request.form.get('tipo_pago', 'Contado').strip()
    es_agente = request.form.get('es_agente_retencion', 'No').strip()
    try: dias_plazo = int(request.form.get('dias_plazo', 0))
    except: dias_plazo = 0
    
    try: tasa_dolar = float(request.form.get('tasa_dolar_venta', 1) or 1)
    except: tasa_dolar = 1.0
    if tasa_dolar <= 0: tasa_dolar = 1.0
    
    try: p_efectivo = float(request.form.get('pago_efectivo', 0) or 0)
    except: p_efectivo = 0.0
    try: p_nequi = float(request.form.get('pago_nequi', 0) or 0)
    except: p_nequi = 0.0
    try: p_bancolombia = float(request.form.get('pago_bancolombia', 0) or 0)
    except: p_bancolombia = 0.0
    try: p_daviplata = float(request.form.get('pago_daviplata', 0) or 0)
    except: p_daviplata = 0.0
    try: p_transferencia = float(request.form.get('pago_transferencia', 0) or 0)
    except: p_transferencia = 0.0
    try: p_dolar_efectivo_usd = float(request.form.get('pago_dolar_efectivo_usd', 0) or 0)
    except: p_dolar_efectivo_usd = 0.0
    try: p_zelle_usd = float(request.form.get('pago_zelle_usd', 0) or 0)
    except: p_zelle_usd = 0.0
    
    p_dolar_efectivo_cop = p_dolar_efectivo_usd * tasa_dolar
    p_zelle_cop = p_zelle_usd * tasa_dolar
    total_abonado_canales = p_efectivo + p_nequi + p_bancolombia + p_daviplata + p_transferencia + p_dolar_efectivo_cop + p_zelle_cop
    
    carrito_json = request.form.get('carrito_json', '[]')
    try: carrito = json.loads(carrito_json)
    except: carrito = []
    
    if not carrito:
        flash('El carrito de ventas está vacío.', 'error')
        return redirect(url_for('index', tab='pestaña3'))
        
    total_bruto = sum(float(item['subtotal']) for item in carrito)
    base_imponible = total_bruto / 1.19
    monto_iva_venta = total_bruto - base_imponible
    
    # ReteFuente 2.5% si el cliente es Empresa / Agente de Retención y supera 10 UVT
    monto_retefuente_retenida = (base_imponible * 0.025) if (es_agente == 'Si' and base_imponible >= tope_uvt) else 0.0
    total_a_cobrar = total_bruto - monto_retefuente_retenida
    
    # Detectar si incluye servicios
    tiene_servicios = any(bool(item.get('es_servicio')) for item in carrito)
    tiene_repuestos = any(not bool(item.get('es_servicio')) for item in carrito)
    tipo_venta = 'Mixto' if (tiene_servicios and tiene_repuestos) else ('Servicios' if tiene_servicios else 'Repuestos')
    
    if tipo_pago == 'Contado':
        if total_abonado_canales == 0:
            p_efectivo = total_a_cobrar
            total_abonado_canales = total_a_cobrar
        monto_abonado = total_a_cobrar
        monto_restante = 0.0
        estado_pago = 'Pagado'
        dias_plazo = 0
    elif tipo_pago == 'Abono':
        monto_abonado = total_abonado_canales
        monto_restante = max(0.0, total_a_cobrar - monto_abonado)
        estado_pago = 'Pagado' if monto_restante == 0 else 'Pendiente'
    else:
        monto_abonado = 0.0
        monto_restante = total_a_cobrar
        estado_pago = 'Pendiente'
        p_efectivo = p_nequi = p_bancolombia = p_daviplata = p_transferencia = p_dolar_efectivo_cop = p_zelle_cop = 0.0
        p_dolar_efectivo_usd = p_zelle_usd = 0.0
        
    fecha_actual = datetime.now()
    fecha_str = fecha_actual.strftime('%d/%m/%Y %H:%M')
    fecha_dia = fecha_actual.strftime('%d/%m/%Y')
    fecha_venc = (fecha_actual + timedelta(days=dias_plazo)).strftime('%d/%m/%Y')
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, compras_count FROM clientes WHERE documento = ?", (doc,))
    cli_db = cursor.fetchone()
    if cli_db:
        nuevo_count = int(cli_db[1] or 1) + 1
        cursor.execute("UPDATE clientes SET nombre=?, telefono=?, direccion=?, compras_count=? WHERE documento=?", (nombre, telefono, direccion, nuevo_count, doc))
    else:
        cursor.execute("INSERT INTO clientes (documento, nombre, telefono, direccion, compras_count) VALUES (?, ?, ?, ?, 1)", (doc, nombre, telefono, direccion))
        
    cursor.execute('''
        INSERT INTO facturas (cliente_doc, cliente_nombre, cliente_telefono, cliente_direccion, fecha, total, tipo_pago, monto_abonado, monto_restante, dias_plazo, estado_pago, fecha_vencimiento, monto_iva, iva_declarado, es_agente_retencion, monto_retefuente, base_imponible, tipo_venta) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'No', ?, ?, ?, ?)
    ''', (doc, nombre, telefono, direccion, fecha_str, total_a_cobrar, tipo_pago, monto_abonado, monto_restante, dias_plazo, estado_pago, fecha_venc, monto_iva_venta, es_agente, monto_retefuente_retenida, base_imponible, tipo_venta))
    
    factura_id = cursor.lastrowid
    
    for item in carrito:
        rep_id = item.get('id')
        cantidad = float(item.get('cantidad', 1))
        precio_unitario = float(item.get('precio', 0))
        subtotal = float(item.get('subtotal', 0))
        es_serv = 1 if item.get('es_servicio') else 0
        
        # Descontar de stock solo si NO es servicio
        if not es_serv and rep_id:
            cursor.execute("SELECT codigo_interno, descripcion, stock FROM inventario WHERE id = ?", (rep_id,))
            rep = cursor.fetchone()
            if rep:
                codigo_interno, descripcion, stock_actual = rep[0], rep[1], float(rep[2] or 0.0)
                nuevo_stock = max(0.0, stock_actual - cantidad)
                cursor.execute("UPDATE inventario SET stock = ? WHERE id = ?", (nuevo_stock, rep_id))
            else:
                codigo_interno, descripcion = item.get('codigo', 'SERV'), item.get('descripcion', 'Servicio')
        else:
            codigo_interno = item.get('codigo', 'SERV')
            descripcion = item.get('descripcion', 'Servicio Técnico / Mano de Obra')
            
        cursor.execute('''
            INSERT INTO factura_detalles (factura_id, codigo_interno, descripcion, cantidad, precio_unitario, subtotal, es_servicio)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (factura_id, codigo_interno, descripcion, cantidad, precio_unitario, subtotal, es_serv))
                           
    canales_cop = [
        ('Efectivo', p_efectivo), ('Nequi', p_nequi), ('Bancolombia', p_bancolombia),
        ('Daviplata', p_daviplata), ('Transferencia', p_transferencia)
    ]
    for canal_nombre, canal_monto in canales_cop:
        if canal_monto > 0:
            cursor.execute('''
                INSERT INTO caja_movimientos (fecha, fecha_dia, tipo, referencia, cliente_o_concepto, metodo_pago, monto, monto_usd, tasa_cambio, cerrado, origen_ingreso)
                VALUES (?, ?, 'Ingreso Venta', ?, ?, ?, ?, 0, 0, 'No', ?)
            ''', (fecha_str, fecha_dia, f"Factura #{factura_id}", nombre, canal_nombre, canal_monto, tipo_venta))

    if p_dolar_efectivo_cop > 0:
        cursor.execute('''
            INSERT INTO caja_movimientos (fecha, fecha_dia, tipo, referencia, cliente_o_concepto, metodo_pago, monto, monto_usd, tasa_cambio, cerrado, origen_ingreso)
            VALUES (?, ?, 'Ingreso Venta', ?, ?, 'Efectivo Dólar ($)', ?, ?, ?, 'No', ?)
        ''', (fecha_str, fecha_dia, f"Factura #{factura_id}", nombre, p_dolar_efectivo_cop, p_dolar_efectivo_usd, tasa_dolar, tipo_venta))

    if p_zelle_cop > 0:
        cursor.execute('''
            INSERT INTO caja_movimientos (fecha, fecha_dia, tipo, referencia, cliente_o_concepto, metodo_pago, monto, monto_usd, tasa_cambio, cerrado, origen_ingreso)
            VALUES (?, ?, 'Ingreso Venta', ?, ?, 'Zelle ($)', ?, ?, ?, 'No', ?)
        ''', (fecha_str, fecha_dia, f"Factura #{factura_id}", nombre, p_zelle_cop, p_zelle_usd, tasa_dolar, tipo_venta))

    conn.commit()
    conn.close()
    flash(f'¡Factura #{factura_id} emitida con éxito! Se calculó IVA y ReteFuente.', 'success')
    return redirect(url_for('index', factura_emitida=factura_id, tab='pestaña3'))


# GENERAR FACTURA EN FORMATO FORMAL ELECTRÓNICO (MODELO EXACTO DE LA FOTO)
@app.route('/descargar_pdf_factura/<int:factura_id>')
def descargar_pdf_factura(factura_id):
    config = get_config()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT cliente_doc, cliente_nombre, cliente_telefono, cliente_direccion, fecha, total, tipo_pago, monto_abonado, monto_restante, monto_iva, es_agente_retencion, monto_retefuente, base_imponible, fecha_vencimiento FROM facturas WHERE id = ?", (factura_id,))
    fac = cursor.fetchone()
    cursor.execute("SELECT codigo_interno, descripcion, cantidad, precio_unitario, subtotal, es_servicio FROM factura_detalles WHERE factura_id = ?", (factura_id,))
    detalles = cursor.fetchall()
    
    cursor.execute("SELECT metodo_pago, monto, COALESCE(monto_usd, 0), COALESCE(tasa_cambio, 0) FROM caja_movimientos WHERE referencia = ? AND tipo = 'Ingreso Venta'", (f"Factura #{factura_id}",))
    canales_pagados = cursor.fetchall()
    conn.close()
    
    if not fac: return "Factura no encontrada", 404
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    elements = []
    styles = getSampleStyleSheet()
    
    title_bold = ParagraphStyle('TBold', parent=styles['Normal'], fontSize=11, leading=13, fontName='Helvetica-Bold')
    normal_sm = ParagraphStyle('NSm', parent=styles['Normal'], fontSize=8, leading=10)
    normal_sm_bold = ParagraphStyle('NSmBold', parent=styles['Normal'], fontSize=8, leading=10, fontName='Helvetica-Bold')
    center_sm = ParagraphStyle('CSm', parent=styles['Normal'], fontSize=7.5, leading=9.5, alignment=1)
    
    # 1. ENCABEZADO ESTILO FACTURA ELECTRÓNICA
    header_table_data = [
        [
            Paragraph(f"<b>{config.get('nombre_negocio')}</b><br/>NIT: {config.get('nit_negocio')}<br/>{config.get('regimen_negocio')}<br/>{config.get('direccion_negocio')}<br/>TEL: {config.get('telefono_negocio')}<br/>{config.get('ciudad_negocio')}", normal_sm),
            Paragraph(f"<b>FACTURA ELECTRÓNICA DE VENTA</b><br/><b>N°:</b> CEAT - {factura_id:05d}<br/><b>FECHA:</b> {fac[4]}<br/><b>VENCE:</b> {fac[13] or fac[4]}<br/><b>FORMA PAGO:</b> {fac[6].upper()}", normal_sm_bold)
        ]
    ]
    t_head = Table(header_table_data, colWidths=[330, 230])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(t_head)
    elements.append(Spacer(1, 8))
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1E3A8A'), spaceAfter=8))
    
    # 2. DATOS DEL CLIENTE
    client_data = [
        [
            Paragraph(f"<b>NIT / CC:</b> {fac[0]}<br/><b>Razón Social:</b> {fac[1]}<br/><b>Vendedora:</b> {config.get('encargada_nombre')}", normal_sm),
            Paragraph(f"<b>Teléfono:</b> {fac[2] or 'N/A'}<br/><b>Dirección:</b> {fac[3] or 'Villa del Rosario'}<br/><b>Ciudad:</b> Cúcuta / Villa del Rosario", normal_sm)
        ]
    ]
    t_cli = Table(client_data, colWidths=[330, 230])
    t_cli.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(t_cli)
    elements.append(Spacer(1, 10))
    
    # 3. TABLA DE ÍTEMS Y PRODUCTOS
    table_items = [["ITEM", "Código", "Cant.", "UND", "Descripción", "Vr. Unit (COP)", "Vr. Parcial"]]
    item_num = 1
    for d in detalles:
        und_tipo = "SRV" if d[5] == 1 else "UND"
        table_items.append([
            str(item_num),
            str(d[0]),
            f"{float(d[2]):g}",
            und_tipo,
            str(d[1]),
            f"$ {float(d[3]):,.2f}",
            f"$ {float(d[4]):,.2f}"
        ])
        item_num += 1
        
    t_prod = Table(table_items, colWidths=[35, 75, 45, 35, 190, 85, 95])
    t_prod.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ALIGN', (0,0), (3,-1), 'CENTER'),
        ('ALIGN', (5,0), (-1,-1), 'RIGHT'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_prod)
    elements.append(Spacer(1, 10))
    
    # 4. TOTALES DESGLOSADOS Y MEDIOS DE PAGO
    base_imp = float(fac[12] or (float(fac[5]) / 1.19))
    m_iva = float(fac[9] or 0)
    m_rete = float(fac[11] or 0)
    total_neto = float(fac[5] or 0)
    
    medios_txt = f"<b>Medio de Pago:</b> {fac[6]}<br/>"
    for met, val_met, val_usd, tasa in canales_pagados:
        if float(val_usd or 0) > 0:
            medios_txt += f"• {met}: ${float(val_usd):,.2f} USD (Tasa: ${float(tasa):,.0f} = ${float(val_met):,.0f} COP)<br/>"
        else:
            medios_txt += f"• {met}: ${float(val_met):,.2f} COP<br/>"
            
    totales_table_data = [
        [
            Paragraph(medios_txt, normal_sm),
            Paragraph("<b>Total Base Imponible:</b><br/><b>IVA (19%):</b>" + (f"<br/><b>(-) ReteFuente 2.5%:</b>" if m_rete > 0 else "") + "<br/><b>VALOR NETO FACTURA:</b>", normal_sm),
            Paragraph(f"$ {base_imp:,.2f}<br/>$ {m_iva:,.2f}" + (f"<br/>- $ {m_rete:,.2f}" if m_rete > 0 else "") + f"<br/><b>$ {total_neto:,.2f}</b>", ParagraphStyle('Rgt', parent=normal_sm, alignment=2))
        ]
    ]
    t_tot = Table(totales_table_data, colWidths=[260, 160, 140])
    t_tot.setStyle(TableStyle([
        ('GRID', (1,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (1,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(t_tot)
    elements.append(Spacer(1, 10))
    
    # Monto en letras y resolución DIAN
    elements.append(Paragraph(f"<b>SON:</b> {numero_a_letras_cop(total_neto)}", normal_sm_bold))
    elements.append(Spacer(1, 6))
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.gray, spaceAfter=6))
    elements.append(Paragraph("Resolución de autorización DIAN N° 18764107127427 desde CEAT-50001 hasta 60000 con vigencia de 24 meses.<br/>AutoRepuestos MERP - ¡Gracias por su preferencia!", center_sm))
    
    doc.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Factura_{factura_id}.pdf", mimetype='application/pdf')


# REGISTRAR ABONO
@app.route('/registrar_abono/<int:factura_id>', methods=['POST'])
def registrar_abono(factura_id):
    try: nuevo_abono = float(request.form.get('monto_abono', 0))
    except: nuevo_abono = 0.0
    metodo = request.form.get('metodo_abono', 'Efectivo').strip()
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT total, monto_abonado, monto_restante, cliente_nombre FROM facturas WHERE id = ?", (factura_id,))
    fac = cursor.fetchone()
    if fac:
        total = float(fac[0] or 0.0)
        abonado_actual = float(fac[1] or 0.0)
        cliente_nom = fac[3] or 'Cliente'
        total_abonado_nuevo = abonado_actual + nuevo_abono
        total_restante_nuevo = max(0.0, total - total_abonado_nuevo)
        estado = 'Pagado' if total_restante_nuevo == 0 else 'Pendiente'
        cursor.execute('UPDATE facturas SET monto_abonado = ?, monto_restante = ?, estado_pago = ? WHERE id = ?', 
                       (total_abonado_nuevo, total_restante_nuevo, estado, factura_id))
                       
        fecha_dt = datetime.now()
        cursor.execute('''
            INSERT INTO caja_movimientos (fecha, fecha_dia, tipo, referencia, cliente_o_concepto, metodo_pago, monto, cerrado)
            VALUES (?, ?, 'Ingreso Abono', ?, ?, ?, ?, 'No')
        ''', (fecha_dt.strftime('%d/%m/%Y %H:%M'), fecha_dt.strftime('%d/%m/%Y'), f"Abono Factura #{factura_id}", cliente_nom, metodo, nuevo_abono))
        
        conn.commit()
        flash(f'¡Abono de $ {nuevo_abono:,.2f} registrado por {metodo} e ingresado al Arqueo de Caja!', 'success')
    conn.close()
    return redirect(url_for('index', tab='pestaña4'))


@app.route('/marcar_pagado_completo/<int:factura_id>')
def marcar_pagado_completo(factura_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT total, monto_restante, cliente_nombre FROM facturas WHERE id = ?", (factura_id,))
    fac = cursor.fetchone()
    if fac:
        total = float(fac[0] or 0.0)
        restante = float(fac[1] or 0.0)
        cliente_nom = fac[2] or 'Cliente'
        cursor.execute("UPDATE facturas SET monto_abonado = ?, monto_restante = 0.0, estado_pago = 'Pagado' WHERE id = ?", (total, factura_id))
        
        if restante > 0:
            fecha_dt = datetime.now()
            cursor.execute('''
                INSERT INTO caja_movimientos (fecha, fecha_dia, tipo, referencia, cliente_o_concepto, metodo_pago, monto, cerrado)
                VALUES (?, ?, 'Ingreso Abono', ?, ?, 'Efectivo', ?, 'No')
            ''', (fecha_dt.strftime('%d/%m/%Y %H:%M'), fecha_dt.strftime('%d/%m/%Y'), f"Pago Total Factura #{factura_id}", cliente_nom, restante))
            
        conn.commit()
        flash('¡Factura marcada como Pagada Completa e ingresada a caja!', 'success')
    conn.close()
    return redirect(url_for('index', tab='pestaña4'))


@app.route('/liquidar_retenciones')
def liquidar_retenciones():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM retenciones")
    conn.commit()
    conn.close()
    flash('¡Retenciones canceladas ante la DIAN y acumulado reiniciado a $ 0.00!', 'success')
    return redirect(url_for('index', tab='pestaña6'))


@app.route('/descargar_corte_gerencial_excel')
def descargar_corte_gerencial_excel():
    config = get_config()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT SUM(total), COUNT(id) FROM facturas")
    r_v = cursor.fetchone()
    tot_ventas = float(r_v[0] or 0.0)
    cnt_facturas = int(r_v[1] or 0)
    
    cursor.execute("SELECT SUM(monto_restante) FROM facturas WHERE estado_pago != 'Pagado'")
    tot_cxc = float(cursor.fetchone()[0] or 0.0)
    
    cursor.execute("SELECT SUM(monto_restante) FROM cuentas_por_pagar WHERE estado_pago != 'Pagado'")
    tot_cxp = float(cursor.fetchone()[0] or 0.0)
    
    cursor.execute("SELECT SUM(monto_retencion) FROM retenciones")
    tot_ret = float(cursor.fetchone()[0] or 0.0)
    
    cursor.execute('''
        SELECT SUM(fd.cantidad * COALESCE(
            (SELECT MAX(inv.costo) FROM inventario inv WHERE inv.codigo_interno = fd.codigo_interno), 0
        ))
        FROM factura_detalles fd
        WHERE COALESCE(fd.es_servicio, 0) = 0
    ''')
    costo_ventas = float(cursor.fetchone()[0] or 0.0)
    utilidad_bruta = max(0.0, tot_ventas - costo_ventas)
    margen_bruto_pct = (utilidad_bruta / tot_ventas * 100) if tot_ventas > 0 else 0.0
    utilidad_antes_impuestos = max(0.0, utilidad_bruta - tot_ret)
    
    cursor.execute("SELECT SUM(costo * MAX(0.0, stock - COALESCE(stock_en_reclamo, 0))), SUM(costo * COALESCE(stock_en_reclamo, 0)) FROM inventario")
    r_inv = cursor.fetchone()
    inv_disponible_valor = float(r_inv[0] or 0.0)
    inv_reclamo_valor = float(r_inv[1] or 0.0)
    
    cursor.execute("SELECT id, cliente_nombre, cliente_doc, cliente_telefono, fecha, total, monto_abonado, monto_restante, dias_plazo, estado_pago, fecha_vencimiento FROM facturas WHERE estado_pago != 'Pagado' ORDER BY id DESC")
    cxc_detalles = cursor.fetchall()
    
    cursor.execute("SELECT nro_factura, nro_control, proveedor, fecha_emision, total_factura, monto_abonado, monto_restante, dias_plazo, estado_pago, fecha_vencimiento FROM cuentas_por_pagar WHERE estado_pago != 'Pagado' ORDER BY id DESC")
    cxp_detalles = cursor.fetchall()
    
    cursor.execute("SELECT codigo_interno, descripcion, categoria, costo, precio_venta, stock, COALESCE(stock_en_reclamo, 0) FROM inventario ORDER BY id DESC")
    inv_detalles = cursor.fetchall()
    
    cursor.execute("SELECT tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, monto_afectado, opcion_reclamo, estado, fecha FROM devoluciones_espera ORDER BY id DESC")
    reclamos_detalles = cursor.fetchall()
    
    conn.close()
    
    wb = openpyxl.Workbook()
    blue_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    gold_header = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
    green_header = PatternFill(start_color="047857", end_color="047857", fill_type="solid")
    font_bold_white = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_title = Font(name="Calibri", size=15, bold=True, color="1E3A8A")
    font_bold = Font(name="Calibri", size=11, bold=True)
    border_thin = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
                         top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))
                         
    ws1 = wb.active
    ws1.title = "Balance Ejecutivo"
    ws1.views.sheetView[0].showGridLines = True
    
    ws1['A1'] = f"{config.get('nombre_negocio')} - CORTE GERENCIAL Y BALANCE DE ESTADO"
    ws1['A1'].font = font_title
    ws1['A2'] = f"Fecha de Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | Responsable: {config.get('encargada_nombre')}"
    ws1['A2'].font = Font(name="Calibri", size=10, italic=True, color="475569")
    
    ws1['A4'] = "INDICADOR FINANCIERO"
    ws1['B4'] = "VALOR (COP)"
    ws1['A4'].fill = blue_header; ws1['A4'].font = font_bold_white
    ws1['B4'].fill = blue_header; ws1['B4'].font = font_bold_white
    ws1['B4'].alignment = Alignment(horizontal="right")
    
    kpis = [
        ("Total Ventas Facturadas", tot_ventas),
        ("Costo de Ventas (COGS)", costo_ventas),
        ("Utilidad Bruta en Ventas", utilidad_bruta),
        ("Margen Bruto de Utilidad (%)", f"{margen_bruto_pct:.2f}%"),
        ("Retenciones por Declarar", tot_ret),
        ("Utilidad Antes de Impuestos", utilidad_antes_impuestos),
        ("Cuentas por Cobrar (CXC)", tot_cxc),
        ("Cuentas por Pagar (CXP)", tot_cxp),
        ("Inventario Disponible Valorizado", inv_disponible_valor),
        ("Mercancía en Reclamo", inv_reclamo_valor),
        ("Facturas Totales Emitidas", cnt_facturas)
    ]
    
    row_idx = 5
    for kpi, val in kpis:
        ws1.cell(row=row_idx, column=1, value=kpi).border = border_thin
        c2 = ws1.cell(row=row_idx, column=2, value=val)
        c2.border = border_thin
        c2.alignment = Alignment(horizontal="right")
        if isinstance(val, (int, float)): c2.number_format = '$ #,##0.00'
        if "Utilidad" in kpi or "Ventas" in kpi:
            ws1.cell(row=row_idx, column=1).font = font_bold
            c2.font = font_bold
        row_idx += 1
        
    ws1['D4'] = "Rentabilidad"; ws1['E4'] = "Monto (COP)"
    ws1['D4'].fill = green_header; ws1['D4'].font = font_bold_white
    ws1['E4'].fill = green_header; ws1['E4'].font = font_bold_white
    ws1.append(["", "", "", "Costo de Ventas", costo_ventas])
    ws1.append(["", "", "", "Utilidad Bruta", utilidad_bruta])
    for r in range(5, 7):
        ws1.cell(row=r, column=4).border = border_thin
        c = ws1.cell(row=r, column=5); c.border = border_thin; c.number_format = '$ #,##0.00'
        
    pie = PieChart()
    labels = Reference(ws1, min_col=4, min_row=5, max_row=6)
    data = Reference(ws1, min_col=5, min_row=4, max_row=6)
    pie.add_data(data, titles_from_data=True); pie.set_categories(labels)
    pie.title = "Composición Ventas (Costo vs Utilidad)"
    pie.width = 14; pie.height = 7.5
    ws1.add_chart(pie, "D8")
    
    ws1.column_dimensions['A'].width = 46; ws1.column_dimensions['B'].width = 24
    ws1.column_dimensions['D'].width = 28; ws1.column_dimensions['E'].width = 20

    ws2 = wb.create_sheet(title="Cuentas por Cobrar (CXC)")
    ws2.views.sheetView[0].showGridLines = True
    ws2.append(["FACTURA", "CLIENTE", "DOCUMENTO", "TELÉFONO", "FECHA", "TOTAL", "ABONADO", "SALDO RESTANTE", "PLAZO", "ESTADO", "VENCIMIENTO"])
    for cell in ws2[1]: cell.fill = blue_header; cell.font = font_bold_white; cell.alignment = Alignment(horizontal="center")
    for cxc in cxc_detalles: ws2.append([f"#{cxc[0]}", cxc[1], cxc[2], cxc[3] or 'N/A', cxc[4], float(cxc[5] or 0), float(cxc[6] or 0), float(cxc[7] or 0), f"{cxc[8]} días", cxc[9], cxc[10] or 'Vigente'])
    
    ws3 = wb.create_sheet(title="Cuentas por Pagar (CXP)")
    ws3.views.sheetView[0].showGridLines = True
    ws3.append(["FACTURA", "CONTROL", "PROVEEDOR", "EMISIÓN", "TOTAL CON IVA", "ABONADO", "DEUDA RESTANTE", "PLAZO", "ESTADO", "VENCIMIENTO"])
    for cell in ws3[1]: cell.fill = gold_header; cell.font = font_bold_white; cell.alignment = Alignment(horizontal="center")
    for cxp in cxp_detalles: ws3.append([cxp[0], cxp[1], cxp[2], cxp[3], float(cxp[4] or 0), float(cxp[5] or 0), float(cxp[6] or 0), f"{cxp[7]} días", cxp[8], cxp[9] or 'Vigente'])

    ws4 = wb.create_sheet(title="Inventario Disponible")
    ws4.views.sheetView[0].showGridLines = True
    ws4.append(["CÓDIGO", "DESCRIPCIÓN", "CATEGORÍA", "COSTO UNIT.", "PRECIO VENTA", "STOCK TOTAL", "EN RECLAMO", "STOCK DISP.", "VALOR TOTAL"])
    for cell in ws4[1]: cell.fill = green_header; cell.font = font_bold_white; cell.alignment = Alignment(horizontal="center")
    for inv in inv_detalles:
        stk_t, stk_r = float(inv[5] or 0), float(inv[6] or 0)
        stk_d = max(0.0, stk_t - stk_r)
        cost_u = float(inv[3] or 0)
        ws4.append([inv[0], inv[1], inv[2], cost_u, float(inv[4] or 0), stk_t, stk_r, stk_d, stk_d * cost_u])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Balance_AutoRepuestosMERP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# BÚSQUEDA DETALLE COMPLETO DE FACTURA PROVEEDOR PARA RECLAMOS
@app.route('/detalle_factura_proveedor/<path:nro>')
def detalle_factura_proveedor(nro):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT nro_factura, nro_control, proveedor, fecha_emision, total_factura, monto_restante, estado_pago, aplica_iva, monto_iva, monto_retefuente FROM cuentas_por_pagar WHERE nro_factura = ?", (nro,))
    cxp = cursor.fetchone()
    cursor.execute("SELECT id, codigo_interno, descripcion, costo, precio_venta, stock, COALESCE(stock_en_reclamo, 0), garantia FROM inventario WHERE nro_factura_prov = ?", (nro,))
    items = cursor.fetchall()
    conn.close()
    
    if not cxp: return jsonify({'encontrado': False})
    
    items_list = [{
        'id': it[0], 'codigo': it[1], 'descripcion': it[2],
        'costo': float(it[3] or 0), 'precio': float(it[4] or 0),
        'stock': float(it[5] or 0), 'en_reclamo': float(it[6] or 0),
        'garantia': it[7]
    } for it in items]
    
    return jsonify({
        'encontrado': True,
        'nro_factura': cxp[0], 'nro_control': cxp[1], 'proveedor': cxp[2],
        'fecha': cxp[3], 'total': float(cxp[4] or 0), 'restante': float(cxp[5] or 0),
        'estado': cxp[6], 'aplica_iva': cxp[7], 'monto_iva': float(cxp[8] or 0),
        'retefuente': float(cxp[9] or 0),
        'items': items_list
    })


@app.route('/buscar_factura_proveedor_reclamo')
def buscar_factura_proveedor_reclamo():
    nro = request.args.get('nro', '').strip()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT nro_factura, nro_control, proveedor, total_factura, monto_restante, estado_pago FROM cuentas_por_pagar WHERE nro_factura = ? LIMIT 1", (nro,))
    cxp = cursor.fetchone()
    cursor.execute("SELECT id, codigo_interno, descripcion, costo, stock, COALESCE(stock_en_reclamo, 0) FROM inventario WHERE nro_factura_prov = ?", (nro,))
    items_rows = cursor.fetchall()
    conn.close()
    if not cxp and not items_rows: return jsonify({'encontrado': False})
    items = [{'id': it[0], 'codigo_interno': it[1], 'descripcion': it[2], 'costo': float(it[3] or 0), 'stock': float(it[4] or 0), 'en_reclamo': float(it[5] or 0)} for it in items_rows]
    return jsonify({
        'encontrado': True, 'nro_factura': cxp[0] if cxp else nro, 'nro_control': cxp[1] if cxp else 'S/N',
        'proveedor': cxp[2] if cxp else (items[0]['descripcion'] if items else 'Proveedor'),
        'total': float(cxp[3] or 0) if cxp else 0.0, 'restante': float(cxp[4] or 0) if cxp else 0.0,
        'estado_pago': cxp[5] if cxp else 'Desconocido', 'items': items
    })


@app.route('/buscar_factura_venta_reclamo')
def buscar_factura_venta_reclamo():
    nro_raw = request.args.get('nro', '').strip().replace('#', '')
    try: fac_id = int(nro_raw)
    except: return jsonify({'encontrado': False})
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, cliente_nombre, cliente_doc, total, monto_abonado, monto_restante, estado_pago FROM facturas WHERE id = ?", (fac_id,))
    fac = cursor.fetchone()
    cursor.execute("SELECT id, codigo_interno, descripcion, cantidad, precio_unitario, subtotal FROM factura_detalles WHERE factura_id = ?", (fac_id,))
    detalles = cursor.fetchall()
    conn.close()
    if not fac: return jsonify({'encontrado': False})
    items = [{'id': dt[0], 'codigo_interno': dt[1], 'descripcion': dt[2], 'cantidad': float(dt[3] or 0), 'precio': float(dt[4] or 0), 'subtotal': float(dt[5] or 0)} for dt in detalles]
    return jsonify({
        'encontrado': True, 'factura_id': fac[0], 'cliente_nombre': fac[1], 'cliente_doc': fac[2],
        'total': float(fac[3] or 0), 'abonado': float(fac[4] or 0), 'restante': float(fac[5] or 0),
        'estado_pago': fac[6], 'items': items
    })


@app.route('/registrar_devolucion_compra', methods=['POST'])
def registrar_devolucion_compra():
    nro_factura = request.form.get('nro_factura', '').strip()
    proveedor = request.form.get('proveedor', '').strip()
    codigo_interno = request.form.get('codigo_interno', '').strip()
    descripcion = request.form.get('descripcion', '').strip()
    cantidad = float(request.form.get('cantidad', 0) or 0)
    costo_unit = float(request.form.get('costo_unitario', 0) or 0)
    opcion = request.form.get('opcion_reclamo', 'Cambio').strip()
    motivo = request.form.get('motivo', 'Mal estado / Defecto').strip()
    
    if cantidad <= 0 or not codigo_interno:
        flash('Datos de reclamo no válidos.', 'error')
        return redirect(url_for('index', tab='pestaña10'))
        
    monto_afectado = cantidad * costo_unit
    fecha_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    if opcion == 'Cambio':
        cursor.execute("UPDATE inventario SET stock_en_reclamo = COALESCE(stock_en_reclamo, 0) + ? WHERE codigo_interno = ? AND nro_factura_prov = ?", 
                       (cantidad, codigo_interno, nro_factura))
        cursor.execute('''
            INSERT INTO devoluciones_espera (tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, opcion_reclamo, monto_afectado, motivo, estado, fecha)
            VALUES ('Compra', ?, ?, ?, ?, ?, 'Cambio de Producto', ?, ?, 'En Espera de Cambio', ?)
        ''', (nro_factura, proveedor, codigo_interno, descripcion, cantidad, monto_afectado, motivo, fecha_str))
        flash(f'¡Reclamo por cambio registrado! {cantidad} un. aisladas a reclamo (no aparecerán para la venta).', 'success')
        
    elif opcion == 'Reintegro':
        cursor.execute("UPDATE inventario SET stock = MAX(0.0, stock - ?) WHERE codigo_interno = ? AND nro_factura_prov = ?", 
                       (cantidad, codigo_interno, nro_factura))
        cursor.execute('''
            INSERT INTO devoluciones_espera (tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, opcion_reclamo, monto_afectado, motivo, estado, fecha)
            VALUES ('Compra', ?, ?, ?, ?, ?, 'Reintegro de Dinero', ?, ?, 'Pendiente Reintegro', ?)
        ''', (nro_factura, proveedor, codigo_interno, descripcion, cantidad, monto_afectado, motivo, fecha_str))
        flash(f'¡Reclamo registrado! El proveedor tiene pendiente reintegrar $ {monto_afectado:,.2f} COP.', 'success')
        
    elif opcion == 'Descuento':
        cursor.execute("SELECT monto_restante FROM cuentas_por_pagar WHERE nro_factura = ?", (nro_factura,))
        cxp = cursor.fetchone()
        if cxp:
            deuda_act = float(cxp[0] or 0)
            nueva_deuda = max(0.0, deuda_act - monto_afectado)
            nuevo_est = 'Pagado' if nueva_deuda == 0 else 'Pendiente'
            cursor.execute("UPDATE cuentas_por_pagar SET monto_restante = ?, estado_pago = ? WHERE nro_factura = ?", (nueva_deuda, nuevo_est, nro_factura))
            
        cursor.execute("UPDATE inventario SET stock = MAX(0.0, stock - ?) WHERE codigo_interno = ? AND nro_factura_prov = ?", 
                       (cantidad, codigo_interno, nro_factura))
        cursor.execute('''
            INSERT INTO devoluciones_espera (tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, opcion_reclamo, monto_afectado, motivo, estado, fecha)
            VALUES ('Compra', ?, ?, ?, ?, ?, 'Descontado de Factura', ?, ?, 'Descontado de CxP', ?)
        ''', (nro_factura, proveedor, codigo_interno, descripcion, cantidad, monto_afectado, motivo, fecha_str))
        flash(f'¡Se descontaron $ {monto_afectado:,.2f} COP de la Cuenta por Pagar de la Factura #{nro_factura}!', 'success')

    conn.commit()
    conn.close()
    return redirect(url_for('index', tab='pestaña10'))


@app.route('/registrar_devolucion_venta', methods=['POST'])
def registrar_devolucion_venta():
    factura_id = request.form.get('factura_id', '').strip()
    cliente_nombre = request.form.get('cliente_nombre', '').strip()
    codigo_interno = request.form.get('codigo_interno', '').strip()
    descripcion = request.form.get('descripcion', '').strip()
    cantidad = float(request.form.get('cantidad', 0) or 0)
    precio_unit = float(request.form.get('precio_unitario', 0) or 0)
    opcion = request.form.get('opcion_reclamo', 'Cambio').strip()
    motivo = request.form.get('motivo', 'Garantía / Defecto').strip()
    
    if cantidad <= 0 or not codigo_interno:
        flash('Datos de reclamo no válidos.', 'error')
        return redirect(url_for('index', tab='pestaña10'))
        
    monto_afectado = cantidad * precio_unit
    fecha_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT estado_pago, monto_restante FROM facturas WHERE id = ?", (factura_id,))
    fac_info = cursor.fetchone()
    esta_pagada = (fac_info and (fac_info[0] == 'Pagado' or float(fac_info[1] or 0) == 0))
    
    if opcion == 'Descuento' and esta_pagada:
        conn.close()
        flash('Esta factura ya está PAGADA completa. No se puede descontar deuda. Elige Cambio o Reintegro.', 'error')
        return redirect(url_for('index', tab='pestaña10'))
        
    cursor.execute("UPDATE inventario SET stock = stock + ?, stock_en_reclamo = COALESCE(stock_en_reclamo, 0) + ? WHERE codigo_interno = ?", 
                   (cantidad, cantidad, codigo_interno))
                   
    if opcion == 'Cambio':
        cursor.execute('''
            INSERT INTO devoluciones_espera (tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, opcion_reclamo, monto_afectado, motivo, estado, fecha)
            VALUES ('Venta', ?, ?, ?, ?, ?, 'Cambio de Producto', ?, ?, 'Pendiente Entrega Cambio', ?)
        ''', (f"#{factura_id}", cliente_nombre, codigo_interno, descripcion, cantidad, monto_afectado, motivo, fecha_str))
        flash(f'¡Devolución registrada! Pieza defectuosa aislada en reclamo. Pendiente entregar reemplazo.', 'success')
        
    elif opcion == 'Reintegro':
        cursor.execute('''
            INSERT INTO devoluciones_espera (tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, opcion_reclamo, monto_afectado, motivo, estado, fecha)
            VALUES ('Venta', ?, ?, ?, ?, ?, 'Reintegro de Dinero', ?, ?, 'Pendiente Reembolso al Cliente', ?)
        ''', (f"#{factura_id}", cliente_nombre, codigo_interno, descripcion, cantidad, monto_afectado, motivo, fecha_str))
        flash(f'¡Devolución registrada! Pendiente reembolsar $ {monto_afectado:,.2f} COP al cliente.', 'success')
        
    elif opcion == 'Descuento':
        deuda_act = float(fac_info[1] or 0)
        nueva_deuda = max(0.0, deuda_act - monto_afectado)
        nuevo_est = 'Pagado' if nueva_deuda == 0 else 'Pendiente'
        cursor.execute("UPDATE facturas SET monto_restante = ?, estado_pago = ? WHERE id = ?", (nueva_deuda, nuevo_est, factura_id))
            
        cursor.execute('''
            INSERT INTO devoluciones_espera (tipo_flujo, referencia_factura, tercero, codigo_interno, descripcion, cantidad, opcion_reclamo, monto_afectado, motivo, estado, fecha)
            VALUES ('Venta', ?, ?, ?, ?, ?, 'Descontado de Deuda', ?, ?, 'Descontado de Factura', ?)
        ''', (f"#{factura_id}", cliente_nombre, codigo_interno, descripcion, cantidad, monto_afectado, motivo, fecha_str))
        flash(f'¡Se descontaron $ {monto_afectado:,.2f} COP de la deuda del cliente en la Factura #{factura_id}!', 'success')

    conn.commit()
    conn.close()
    return redirect(url_for('index', tab='pestaña10'))


@app.route('/resolver_devolucion/<int:dev_id>/<accion>')
def resolver_devolucion(dev_id, accion):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT tipo_flujo, codigo_interno, cantidad FROM devoluciones_espera WHERE id = ?", (dev_id,))
    dev = cursor.fetchone()
    if dev:
        tipo_flujo, cod_int, cant = dev[0], dev[1], float(dev[2] or 0)
        if accion == 'recibido_cambio':
            cursor.execute("UPDATE inventario SET stock_en_reclamo = MAX(0.0, stock_en_reclamo - ?) WHERE codigo_interno = ?", (cant, cod_int))
            cursor.execute("UPDATE devoluciones_espera SET estado = 'Cambio Recibido (Liberado)' WHERE id = ?", (dev_id,))
            flash('¡Mercancía de reemplazo liberada al inventario disponible para la venta!', 'success')
        elif accion == 'reintegro_recibido':
            cursor.execute("UPDATE devoluciones_espera SET estado = 'Dinero Reintegrado' WHERE id = ?", (dev_id,))
            flash('¡Reintegro marcado como cobrado exitosamente!', 'success')
        elif accion == 'entregar_cambio_cliente':
            cursor.execute("UPDATE inventario SET stock = MAX(0.0, stock - ?) WHERE codigo_interno = ?", (cant, cod_int))
            cursor.execute("UPDATE devoluciones_espera SET estado = 'Cambio Entregado a Cliente' WHERE id = ?", (dev_id,))
            flash('¡Cambio entregado al cliente y descontado del stock disponible!', 'success')
        elif accion == 'eliminar':
            cursor.execute("DELETE FROM devoluciones_espera WHERE id = ?", (dev_id,))
            flash('Registro eliminado del historial de reclamos.', 'success')
        conn.commit()
    conn.close()
    return redirect(url_for('index', tab='pestaña10'))


@app.route('/eliminar_repuesto/<int:id>')
def eliminar_repuesto(id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM inventario WHERE id = ?", (id,))
    conn.commit()
    conn.close()
    flash('¡Repuesto eliminado del inventario correctamente!', 'success')
    return redirect(url_for('index', tab='pestaña2'))


@app.route('/actualizar_precio/<int:id>', methods=['POST'])
def actualizar_precio(id):
    try: nuevo_precio = float(request.form.get('nuevo_precio', 0))
    except: nuevo_precio = 0.0
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE inventario SET precio_venta = ? WHERE id = ?", (nuevo_precio, id))
    conn.commit()
    conn.close()
    flash('¡Precio de venta actualizado correctamente!', 'success')
    return redirect(url_for('index', tab='pestaña2'))


@app.route('/marcar_pagado_cxp/<int:cxp_id>')
def marcar_pagado_cxp(cxp_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT total_factura FROM cuentas_por_pagar WHERE id = ?", (cxp_id,))
    cxp = cursor.fetchone()
    if cxp:
        total = float(cxp[0] or 0.0)
        cursor.execute("UPDATE cuentas_por_pagar SET monto_abonado = ?, monto_restante = 0.0, estado_pago = 'Pagado' WHERE id = ?", (total, cxp_id))
        conn.commit()
        flash('¡Cuenta por Pagar marcada como Pagada!', 'success')
    conn.close()
    return redirect(url_for('index', tab='pestaña5'))


@app.route('/registrar_abono_cxp/<int:cxp_id>', methods=['POST'])
def registrar_abono_cxp(cxp_id):
    try: nuevo_abono = float(request.form.get('monto_abono', 0))
    except: nuevo_abono = 0.0
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT total_factura, monto_abonado, monto_restante FROM cuentas_por_pagar WHERE id = ?", (cxp_id,))
    cxp = cursor.fetchone()
    if cxp:
        total = float(cxp[0] or 0.0)
        abonado_actual = float(cxp[1] or 0.0)
        total_abonado_nuevo = abonado_actual + nuevo_abono
        total_restante_nuevo = max(0.0, total - total_abonado_nuevo)
        estado = 'Pagado' if total_restante_nuevo == 0 else 'Pendiente'
        cursor.execute('UPDATE cuentas_por_pagar SET monto_abonado = ?, monto_restante = ?, estado_pago = ? WHERE id = ?', 
                       (total_abonado_nuevo, total_restante_nuevo, estado, cxp_id))
        conn.commit()
        flash('¡Abono registrado en la Cuenta por Pagar!', 'success')
    conn.close()
    return redirect(url_for('index', tab='pestaña5'))


@app.route('/verificar_cliente', methods=['POST'])
def verificar_cliente():
    doc = request.form.get('documento', '').strip()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT nombre, telefono, direccion, compras_count FROM clientes WHERE documento = ?", (doc,))
    cliente = cursor.fetchone()
    conn.close()
    if cliente:
        return {'encontrado': True, 'nombre': cliente[0], 'telefono': cliente[1], 'direccion': cliente[2], 'compras': int(cliente[3] or 1) + 1}
    else:
        return {'encontrado': False, 'compras': 1}


@app.route('/buscar_repuestos_json')
def buscar_repuestos_json():
    q = request.args.get('q', '').strip()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    like_q = f"%{q}%"
    cursor.execute("SELECT id, codigo_interno, descripcion, precio_venta, (stock - COALESCE(stock_en_reclamo, 0)) as stock_disp, costo FROM inventario WHERE codigo_interno LIKE ? OR descripcion LIKE ? OR codigo_barras LIKE ? LIMIT 10", 
                   (like_q, like_q, like_q))
    rows = cursor.fetchall()
    conn.close()
    results = []
    for r in rows:
        results.append({
            'id': r[0], 'codigo': r[1], 'descripcion': r[2],
            'precio': float(r[3] if r[3] is not None else 0.0), 
            'stock': max(0.0, float(r[4] if r[4] is not None else 0.0)), 
            'costo': float(r[5] if r[5] is not None else 0.0),
            'es_servicio': False
        })
    return jsonify(results)


@app.route('/buscar_repuestos_inventario_json')
def buscar_repuestos_inventario_json():
    q = request.args.get('q', '').strip()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    like_q = f"%{q}%"
    cursor.execute("SELECT codigo_interno, descripcion, categoria, codigo_barras, codigo_opcional, precio_venta, costo, proveedor FROM inventario WHERE descripcion LIKE ? OR codigo_interno LIKE ? LIMIT 10", 
                   (like_q, like_q))
    rows = cursor.fetchall()
    conn.close()
    results = []
    for r in rows:
        results.append({
            'codigo_interno': r[0], 'descripcion': r[1], 'categoria': r[2],
            'codigo_barras': r[3] or '', 'codigo_opcional': r[4] or '',
            'precio_venta': float(r[5] if r[5] is not None else 0.0), 
            'costo': float(r[6] if r[6] is not None else 0.0), 
            'proveedor': r[7] or ''
        })
    return jsonify(results)


# ENTRADA DE FACTURA PROVEEDOR (CON CÁLCULO DE RETEFUENTE 2.5% SI DECLARA RENTA)
@app.route('/registrar_factura_proveedor', methods=['POST'])
def registrar_factura_proveedor():
    config = get_config()
    valor_uvt = float(config.get('valor_uvt', 52374))
    tope_uvt = valor_uvt * 10.0
    
    nro_factura = request.form.get('nro_factura_prov', '').strip() or 'S/N'
    nro_control = request.form.get('nro_control', '').strip() or 'S/N'
    proveedor = request.form.get('proveedor', 'Proveedor General').strip()
    aplica_iva = request.form.get('aplica_iva', 'No')
    declara_renta = request.form.get('declara_renta', 'Si')
    try: dias_credito = int(request.form.get('dias_credito', 0))
    except: dias_credito = 0
    carrito_json = request.form.get('carrito_compra_json', '[]')
    try: carrito = json.loads(carrito_json)
    except: carrito = []
    
    if not carrito:
        flash('El carrito de entrada de mercancía está vacío.', 'error')
        return redirect(url_for('index', tab='pestaña1'))
        
    subtotal_factura = sum(float(item['costo']) * float(item['cantidad']) for item in carrito)
    monto_iva = subtotal_factura * 0.19 if aplica_iva == 'Si' else 0.0
    
    # ReteFuente 2.5% sobre la BASE imponible si supera 10 UVT ($523.740) y declara renta
    monto_retefuente = (subtotal_factura * 0.025) if (declara_renta == 'Si' and subtotal_factura >= tope_uvt) else 0.0
    
    # El valor neto a pagarle al proveedor se reduce por la ReteFuente
    total_con_iva_neto = (subtotal_factura + monto_iva) - monto_retefuente
    
    fecha_dt = datetime.now()
    fecha_str = fecha_dt.strftime('%d/%m/%Y')
    fecha_venc = (fecha_dt + timedelta(days=dias_credito)).strftime('%d/%m/%Y')
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    productos_resumen_str = ", ".join([f"{i['descripcion']} ({i['cantidad']} un.)" for i in carrito])
    
    cursor.execute('''
        INSERT INTO cuentas_por_pagar 
        (nro_factura, nro_control, proveedor, fecha_emision, dias_plazo, fecha_vencimiento, total_factura, monto_abonado, monto_restante, estado_pago, productos_resumen, aplica_iva, monto_iva, monto_retencion, iva_declarado, proveedor_declara_renta, monto_retefuente, base_imponible)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'No', ?, ?, ?)
    ''', (nro_factura, nro_control, proveedor, fecha_str, dias_credito, fecha_venc, total_con_iva_neto, 0.0, total_con_iva_neto, 'Pendiente', productos_resumen_str, aplica_iva, monto_iva, monto_retefuente, declara_renta, monto_retefuente, subtotal_factura))
    
    if monto_retefuente > 0:
        cursor.execute('''
            INSERT INTO retenciones (nro_factura, nro_control, proveedor, fecha, base_imponible, monto_iva, monto_retencion, tipo_retencion)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'ReteFuente 2.5%')
        ''', (nro_factura, nro_control, proveedor, fecha_str, subtotal_factura, monto_iva, monto_retefuente))
        
    for item in carrito:
        cursor.execute('''
            INSERT INTO inventario 
            (codigo_interno, descripcion, categoria, codigo_barras, codigo_opcional, serial, nro_factura_prov, nro_control, costo, precio_venta, stock, proveedor, dias_credito, garantia, imagen, imagen2, imagen3, fecha_registro, stock_en_reclamo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', '', '', ?, 0)
        ''', (item['codigo_interno'], item['descripcion'], item['categoria'], item.get('codigo_barras', ''), item.get('codigo_opcional', ''), item.get('serial', ''), nro_factura, nro_control, float(item['costo']), float(item['precio_venta']), float(item['cantidad']), proveedor, dias_credito, item.get('garantia', 'Sin garantía'), fecha_str))
        
    conn.commit()
    conn.close()
    flash(f'¡Factura Proveedor #{nro_factura} registrada! Se descontó ReteFuente 2.5% ($ {monto_retefuente:,.2f}) de la cuenta por pagar.', 'success')
    return redirect(url_for('index', tab='pestaña1'))


@app.route('/exportar_clientes_excel')
def exportar_clientes_excel():
    config = get_config()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT codigo_interno, descripcion, categoria, codigo_barras, precio_venta, stock, COALESCE(stock_en_reclamo, 0) FROM inventario ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Catálogo Disponible"
    ws.views.sheetView[0].showGridLines = True
    
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    ws.append(["CÓDIGO INTERNO", "DESCRIPCIÓN", "CATEGORÍA", "CÓDIGO BARRAS", "PRECIO VENTA", "STOCK DISPONIBLE"])
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for r in rows:
        stk_disp = max(0.0, float(r[5] or 0) - float(r[6] or 0))
        ws.append([r[0], r[1], r[2], r[3], float(r[4] or 0), stk_disp])
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"Catalogo_{config.get('nombre_negocio')}_{datetime.now().strftime('%Y%m%d')}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/generar_presupuesto_pdf', methods=['POST'])
def generar_presupuesto_pdf():
    config = get_config()
    cliente_nombre = request.form.get('presupuesto_cliente', 'Cliente General')
    rif = request.form.get('presupuesto_rif', 'S/N')
    carrito_json = request.form.get('presupuesto_carrito_json', '[]')
    try: carrito = json.loads(carrito_json)
    except: carrito = []
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    elements = []
    styles = getSampleStyleSheet()
    
    title_bold = ParagraphStyle('TBold', parent=styles['Normal'], fontSize=11, leading=13, fontName='Helvetica-Bold')
    normal_sm = ParagraphStyle('NSm', parent=styles['Normal'], fontSize=8, leading=10)
    normal_sm_bold = ParagraphStyle('NSmBold', parent=styles['Normal'], fontSize=8, leading=10, fontName='Helvetica-Bold')
    center_sm = ParagraphStyle('CSm', parent=styles['Normal'], fontSize=7.5, leading=9.5, alignment=1)
    
    header_table_data = [
        [
            Paragraph(f"<b>{config.get('nombre_negocio')}</b><br/>NIT: {config.get('nit_negocio')}<br/>{config.get('direccion_negocio')}<br/>TEL: {config.get('telefono_negocio')}<br/>{config.get('ciudad_negocio')}", normal_sm),
            Paragraph(f"<b>COTIZACIÓN / PRESUPUESTO</b><br/><b>FECHA:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}<br/><b>CLIENTE:</b> {cliente_nombre}<br/><b>NIT / CC:</b> {rif}", normal_sm_bold)
        ]
    ]
    t_head = Table(header_table_data, colWidths=[330, 230])
    t_head.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(t_head)
    elements.append(Spacer(1, 8))
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1E3A8A'), spaceAfter=8))
    
    table_items = [["ITEM", "Código", "Cant.", "Descripción", "Vr. Unit (COP)", "Vr. Parcial (COP)"]]
    item_num = 1
    subtotal_general = 0.0
    for d in carrito:
        sub = float(d.get('cantidad', 1)) * float(d.get('precio', 0))
        subtotal_general += sub
        table_items.append([
            str(item_num),
            str(d.get('codigo', 'REP')),
            str(d.get('cantidad', 1)),
            str(d.get('descripcion', '')),
            f"$ {float(d.get('precio', 0)):,.2f}",
            f"$ {sub:,.2f}"
        ])
        item_num += 1
        
    t_prod = Table(table_items, colWidths=[35, 75, 45, 230, 85, 90])
    t_prod.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ALIGN', (0,0), (2,-1), 'CENTER'),
        ('ALIGN', (4,0), (-1,-1), 'RIGHT'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_prod)
    elements.append(Spacer(1, 10))
    
    base_imp = subtotal_general
    iva_calc = base_imp * 0.19
    total_neto = base_imp + iva_calc
    
    tot_data = [
        [
            Paragraph("Precios sujetos a cambio sin previo aviso.<br/>Validez de esta cotización: 5 días hábiles.", normal_sm),
            Paragraph(f"<b>Subtotal Base:</b><br/><b>IVA (19%):</b><br/><b>TOTAL EN PESOS (COP):</b>", normal_sm),
            Paragraph(f"$ {base_imp:,.2f}<br/>$ {iva_calc:,.2f}<br/><b>$ {total_neto:,.2f}</b>", ParagraphStyle('Rgt', parent=normal_sm, alignment=2))
        ]
    ]
    t_tot = Table(tot_data, colWidths=[280, 140, 140])
    t_tot.setStyle(TableStyle([
        ('GRID', (1,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (1,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(t_tot)
    
    doc.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Presupuesto_{cliente_nombre}.pdf", mimetype='application/pdf')


@app.route('/descargar_pdf_retencion/<int:ret_id>')
def descargar_pdf_retencion(ret_id):
    config = get_config()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT nro_factura, nro_control, proveedor, fecha, base_imponible, monto_iva, monto_retencion, COALESCE(tipo_retencion, 'ReteFuente 2.5%') FROM retenciones WHERE id = ?", (ret_id,))
    ret = cursor.fetchone()
    
    if not ret: 
        conn.close()
        return "Comprobante no encontrado", 404
        
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=14, textColor=colors.HexColor('#1E3A8A'), alignment=1)
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#555555'), alignment=1)
    
    elements.append(Paragraph(config.get('nombre_negocio'), title_style))
    elements.append(Paragraph(f"NIT: {config.get('nit_negocio')} | {config.get('direccion_negocio')} | TEL: {config.get('telefono_negocio')}", sub_style))
    elements.append(Paragraph("<b>CERTIFICADO DE RETENCIÓN EN LA FUENTE</b>", title_style))
    elements.append(Spacer(1, 15))
    
    bi, miva, mret, t_ret = float(ret[4] or 0), float(ret[5] or 0), float(ret[6] or 0), ret[7]
    table_data = [
        ["Concepto Tributario", "Base Imponible (COP)", "Tarifa Aplicada", "Monto Retenido"],
        [t_ret, f"$ {bi:,.2f}", "2.5% Renta", f"$ {mret:,.2f}"]
    ]
    t_det = Table(table_data, colWidths=[180, 120, 120, 120])
    t_det.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ALIGN', (1,0), (-1,-1), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(t_det)
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"Expedido a favor de: <b>{ret[2]}</b> correspondiente a la Factura N° {ret[0]}", sub_style))
    elements.append(Spacer(1, 35))
    
    t_firma = Table([["____________________________"], [f"Firma Autorizada: {config.get('encargada_nombre')}"]], colWidths=[260])
    t_firma.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    elements.append(t_firma)
    
    doc.build(elements)
    
    cursor.execute("DELETE FROM retenciones WHERE id = ?", (ret_id,))
    conn.commit()
    conn.close()
    
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=f"Certificado_Retencion_{ret[0]}.pdf", mimetype='application/pdf')


if __name__ == '__main__':
    print("Iniciando sistema AutoRepuestos MERP para Dayana Rujano...")
    app.run(debug=True, port=5000)