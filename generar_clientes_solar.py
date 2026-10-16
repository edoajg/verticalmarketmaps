# -*- coding: utf-8 -*-
"""Clientes de la vista Solar: propietarios de las centrales fotovoltaicas por región.

Lee el Reporte de Capacidad Instalada de Generación de la CNE (hoja «base»),
agrupa cada sociedad titular en su grupo económico y escribe el bloque
CLIENTES de index.html.

    python generar_clientes_solar.py                          # imprime el bloque
    python generar_clientes_solar.py index.html               # lo reemplaza entre los marcadores
    python generar_clientes_solar.py index.html --plantilla   # además rehace distribuidores.csv

Distribuidores: distribuidores.csv tiene una fila por región y cliente. Se
completa la columna «distribuidor» y, al volver a correr el script con
index.html, esas asignaciones pasan al bloque DISTRIB de la página. El CSV
solo se escribe si no existe o con --plantilla, que lo rehace con los
clientes vigentes y conserva lo ya escrito en «distribuidor» y «observaciones».

La CNE registra la sociedad dueña de cada central, casi siempre un SPV. El paso
a grupo económico es elaboración propia: GRUPO para las sociedades conocidas,
REGLAS para los prefijos que delatan a la matriz y FAMILIAS para series de SPV
con nombre común cuya matriz aún no está identificada.
"""
import csv, io, os, openpyxl, re, sys
from collections import Counter, defaultdict
sys.stdout.reconfigure(encoding='utf-8')

ARCHIVO = 'CNE_Capacidad_Instalada_Generacion_2026-10.xlsx'
CSV = 'distribuidores.csv'
CORTE = 'oct-2026'
TOP = 12            # clientes por región con detalle de centrales

REG = {'Región de Arica y Parinacota': 15, 'Región de Tarapacá': 1,
       'Región de Antofagasta': 2, 'Región de Atacama': 3, 'Región de Coquimbo': 4,
       'Región de Valparaíso': 5, 'Región Metropolitana de Santiago': 13,
       'Región del Libertador Gral. Bernardo O’Higgins': 6, 'Región del Maule': 7,
       'Región de Ñuble': 16, 'Región del Biobío': 8, 'Región de La Araucanía': 9,
       'Región de Los Ríos': 14, 'Región de Los Lagos': 10,
       'Región Aisén del Gral.Carlos Ibáñez del Campo': 11,
       'Región de Magallanes y de la Antártica Chilena': 12}
CODS = [15, 1, 2, 3, 4, 5, 13, 6, 7, 16, 8, 9, 14, 10, 11, 12]

# clave -> (nombre, matriz o aclaración)
GRUPOS = {
 'enel': ('Enel Green Power', 'Enel Chile'),
 'aes': ('AES Andes', 'AES Corp.'),
 'sonnedix': ('Sonnedix', 'J.P. Morgan AM'),
 'atlas': ('Atlas Renewable Energy', ''),
 'acciona': ('Acciona Energía', ''),
 'engie': ('Engie Chile', ''),
 'gm': ('Generadora Metropolitana', 'AME + EDF'),
 'metlen': ('Metlen', 'ex Mytilineos'),
 'grenergy': ('Grenergy', ''),
 'contour': ('ContourGlobal', 'KKR'),
 'cvcdif': ('CVC DIF', ''),
 'andesren': ('Andes Renovables', 'Mainstream'),
 'colbun': ('Colbún', ''),
 'copec': ('Copec', ''),
 'cerro': ('Grupo Cerro', 'Cerro Dominador'),
 'total': ('TotalEnergies', 'TSGF, 50%'),
 'pacific': ('Pacific Hydro', 'SPIC'),
 'innergex': ('Innergex', ''),
 'opdenergy': ('Opdenergy', ''),
 'xelio': ('X-Elio', ''),
 'naturgy': ('Naturgy', 'GPG'),
 'toesca': ('Toesca', ''),
 'weg': ('WEG Capital', 'fondo WEG-4'),
 'pattern': ('Pattern Energy', ''),
 'sojitz': ('Sojitz', 'con Shikoku y Eiffage'),
 'actis': ('Actis', ''),
 'rijn': ('Rijn Capital', ''),
 'repsol': ('Repsol · Ibereólica', ''),
 'solarpack': ('Solarpack', ''),
 'solek': ('Solek', ''),
 'cve': ('CVE', ''),
 'reden': ('Reden', ''),
 'enlasa': ('Enlasa', ''),
 'farmdo': ('Farmdo Energy', ''),
 'enfragen': ('EnfraGen', 'Prime Energía'),
}

# Sociedad titular (columna «propietario») -> (grupo, confianza).
# v: confirmado con fuente reciente · p: probable (fuente antigua o indirecta)
GRUPO = {
 'ENEL GREEN POWER CHILE S.A.': ('enel', 'v'),
 'ARCADIA GENERACIÓN SOLAR S.A.': ('sonnedix', 'v'),        # comprada a Enel, oct-2023
 'ATACAMA SOLAR SPA': ('sonnedix', 'v'),
 'TERCERA REGIÓN SOLAR SPA': ('sonnedix', 'p'),             # Meseta de los Andes (SEA 2020)
 'AES ANDES S.A.': ('aes', 'v'),
 'ANDES SOLAR SPA': ('aes', 'v'), 'ANDES SOLAR II SPA': ('aes', 'v'),
 'ANDES SOLAR III SPA': ('aes', 'v'), 'ANDES SOLAR IV SPA': ('aes', 'v'),
 'BOLERO SPA': ('aes', 'v'),                                # comprada a EDF y Marubeni, 2023
 'PARQUE SOLAR FOTOVOLTAICO SOL DEL DESIERTO SPA': ('atlas', 'v'),
 'ESTEPA SOLAR SPA': ('atlas', 'v'),
 'JAVIERA SPA': ('atlas', 'v'),
 'CHUNGUNGO S.A.': ('atlas', 'p'),                          # Quilapilún
 'ACCIONA ENERGÍA CHILE HOLDINGS S.A.': ('acciona', 'v'),
 'ALMEYDA SPA': ('acciona', 'v'),
 'ENGIE ENERGÍA CHILE S.A.': ('engie', 'v'),
 'SOLAR LOS LOROS SPA': ('engie', 'v'),
 'CEME 1 SPA': ('gm', 'v'),
 'SANTIAGO SOLAR S.A.': ('gm', 'v'),
 'PLANTA SOLAR TOCOPILLA SPA': ('metlen', 'p'),             # Taira
 'TAMARICO SOLAR DOS SPA': ('metlen', 'p'),
 'INVERSIONES FOTOVOLTAICAS SPA': ('metlen', 'v'),          # Willka
 'DOÑA ANTONIA SOLAR SPA': ('metlen', 'v'),
 'GR LENGA SPA': ('cvcdif', 'v'),                           # Gabriela, vendida por Grenergy
 'GR CHAÑAR SPA': ('contour', 'v'),                         # Víctor Jara, Oasis de Atacama
 'PARQUE FOTOVOLTAICO NUEVO QUILLAGUA SPA': ('contour', 'v'),
 'GR ALGARROBO SPA': ('grenergy', 'v'), 'GR LIUN SPA': ('grenergy', 'v'),
 'AR ESCONDIDO SPA': ('andesren', 'v'), 'AR VALLE ESCONDIDO SPA': ('andesren', 'v'),
 'AR PAMPA SPA': ('andesren', 'v'),
 'COLBÚN S.A.': ('colbun', 'v'),
 'COPEC RENOVABLES SPA': ('copec', 'v'), 'COPEC RENOVABLES II SPA': ('copec', 'v'),
 'CERRO DOMINADOR PV S.A.': ('cerro', 'v'),
 'TSGF SPA': ('total', 'v'),
 'COPIAPÓ SOLAR SPA': ('pacific', 'p'),
 'PV SALVADOR S.A.': ('innergex', 'v'), 'SAN ANDRÉS SPA': ('innergex', 'v'),
 'SOL DE LOS ANDES SPA': ('opdenergy', 'v'),
 'FOTOVOLTAICA NORTE GRANDE 1 SPA': ('xelio', 'v'),
 'FOTOVOLTAICA NORTE GRANDE 5 SPA': ('xelio', 'v'),
 'GPG SOLAR CHILE 2017 SPA': ('naturgy', 'v'),
 'GPG GENERACIÓN DISTRIBUIDA SPA': ('naturgy', 'v'),
 'PARQUE SOLAR FOTOVOLTAICO LUZ DEL NORTE SPA': ('toesca', 'v'),
 'GENERACIÓN SOLAR SPA': ('weg', 'v'),
 'CONEJO SOLAR SPA': ('pattern', 'p'),
 'AUSTRIANSOLAR CHILE CUATRO SPA': ('sojitz', 'p'),
 'EL PELÍCANO SOLAR COMPANY SPA': ('actis', 'p'),
 'PLANTA SOLAR SAN PEDRO III SPA': ('rijn', 'p'),
 'SOLAR ELENA SPA': ('repsol', 'p'),
 'POZO ALMONTE SOLAR 1 S.A.': ('solarpack', 'p'),
 'POZO ALMONTE SOLAR 2 S.A.': ('solarpack', 'p'),
 'POZO ALMONTE SOLAR 3 S.A.': ('solarpack', 'p'),
 'PARQUE SOLAR LEYDA SPA': ('solek', 'p'),
 'ENLASA GENERACIÓN CHILE S.A.': ('enlasa', 'v'),
 'FARMDO ENERGY CHILE SPA': ('farmdo', 'v'),
 'ENFRAGEN CHILE SOLAR SPA': ('enfragen', 'v'),
}

# prefijos de SPV que identifican a la matriz
REGLAS = [
 (r'^SONNEDIX ', 'sonnedix', 'v'),
 (r'^PARSOSY ', 'sonnedix', 'p'),
 (r'^CVE PROYECTO ', 'cve', 'v'),
 (r'^REDEN ', 'reden', 'v'),
 (r'^SLK ', 'solek', 'p'),
 (r'^GR ', 'grenergy', 'p'),          # convención de nombres de Grenergy (árboles nativos)
]

# series de SPV con un nombre común: probable matriz única, aún sin identificar
FAMILIAS = [
 (r'^PFV (EL|LA|LAS|LOS) ', 'pfv-aves', 'Serie «PFV + ave»'),
 (r'^PV (EL|LA|LAS|LOS) ', 'pv-aves', 'Serie «PV + ave»'),
 (r' DEL? VERANO SPA$', 'verano', 'Serie «de Verano»'),
 (r'^SOLAR TI ', 'solar-ti', 'Serie «Solar TI»'),
 (r'^ANDINA SOLAR ', 'andina-solar', 'Serie «Andina Solar»'),
 (r'^PSF ', 'psf', 'Serie «PSF»'),
 (r' SUNLIGHT SPA$', 'sunlight', 'Serie «Sunlight»'),
 (r'^TEDLAR ', 'tedlar', 'Serie «Tedlar»'),
 (r'^IMPULSO SOLAR ', 'impulso', 'Serie «Impulso Solar»'),
 (r'^BLUE SOLAR ', 'blue-solar', 'Serie «Blue Solar»'),
 (r'^CHESTER SOLAR ', 'chester', 'Serie «Chester Solar»'),
 (r'^CHUNGUNGO SOLAR ', 'chungungo-solar', 'Serie «Chungungo Solar»'),
 (r'^EMPRESA ELÉCTRICA ', 'ee-arboles', 'Serie «Empresa Eléctrica»'),
]
NOTA_FAM = {'verano': 'originadas por Verano Energy; parte de la cartera se vendió a Matrix Renewables'}

BAJA = {'de', 'del', 'la', 'las', 'el', 'los', 'y', 'e'}
ROM = re.compile(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)(-[A-Z])?$')
SIGLAS = {'GR', 'PSF', 'CEME', 'SLK', 'CB', 'CVE', 'EA', 'SF', 'TI', 'PVP', 'PAS1', 'FV', 'PFV', 'PV'}


def titulo(s):
    out = []
    for j, w in enumerate(s.split()):
        lw = w.lower()
        if ROM.match(w) or w in SIGLAS or re.search(r'\d', w) and len(w) <= 4:
            out.append(w)
        elif lw in BAJA and j > 0:
            out.append(lw)
        elif lw in ('spa', 's.a.', 'ltda.'):
            out.append({'spa': 'SpA', 's.a.': 'S.A.', 'ltda.': 'Ltda.'}[lw])
        else:
            out.append(lw[:1].upper() + lw[1:])
    return ' '.join(out)


def central(n):
    n = re.sub(r'\s+', ' ', n).strip()
    csp = n.startswith('CSP ')
    n = re.sub(r'^(PMGD |PMG )?(PFV |PV |CSP )?(FV )?', '', n)
    return titulo(n) + (' CSP' if csp else '')


def slug(s):
    s = s.lower()
    for a, b in zip('áéíóúñü', 'aeiounu'):
        s = s.replace(a, b)
    s = re.sub(r'\s(spa|s\.a\.|ltda\.?)$', '', s)
    return re.sub(r'[^a-z0-9]+', '-', s).strip('-')


def cliente(prop):
    """sociedad titular -> (clave, confianza, tipo); tipo g=grupo f=serie s=sociedad"""
    if prop in GRUPO:
        k, c = GRUPO[prop]
        return k, c, 'g'
    for rx, k, c in REGLAS:
        if re.search(rx, prop):
            return k, c, 'g'
    for rx, k, _ in FAMILIAS:
        if re.search(rx, prop):
            return 'f:' + k, 'p', 'f'
    return 's:' + slug(prop), 'v', 's'


wb = openpyxl.load_workbook(ARCHIVO, read_only=True, data_only=True)
filas = list(wb['base'].iter_rows(values_only=True))
I = {h: i for i, h in enumerate(filas[1]) if h}
centrales = []
for r in filas[2:]:
    if not r or not r[I['central']] or 'solar' not in str(r[I['tipo_de_energia']]).lower():
        continue
    if r[I['central']].startswith('PE '):            # PE La Cabaña: eólica mal clasificada
        continue
    if r[I['comuna_nombre']] == 'Isla de Pascua':    # fuera del mapa
        continue
    prop = re.sub(r'\s+', ' ', r[I['propietario']]).strip()
    k, conf, tipo = cliente(prop)
    f = r[I['fecha_puesta_servicio_central']]
    centrales.append(dict(cod=REG[r[I['region_nombre']]], k=k, conf=conf, tipo=tipo, prop=prop,
                          n=central(r[I['central']]), c=r[I['comuna_nombre']],
                          mw=float(r[I['potencia_neta_mw']] or 0), a=f.year if f else None,
                          pmgd=r[I['medio_generacion']] == 'PMGD'))

nombres = {k: v for k, v in GRUPOS.items()}
for p in centrales:
    if p['tipo'] == 'f':
        fk = p['k'][2:]
        et = next(e for _, kk, e in FAMILIAS if kk == fk)
        nombres[p['k']] = (et, NOTA_FAM.get(fk, ''))
    elif p['tipo'] == 's':
        nombres[p['k']] = (titulo(p['prop']), '')


def num(x):
    return ('%.1f' % x).rstrip('0').rstrip('.')


def js(s):
    return "'" + s.replace('\\', '\\\\').replace("'", "\\'") + "'"


out = ['/* >>> CLIENTES · generado por generar_clientes_solar.py a partir de %s */' % ARCHIVO,
       'const CLIENTES = {', "  corte:'%s'," % CORTE]

# nacional
gk = defaultdict(lambda: [0.0, 0, set()])
for p in centrales:
    gk[p['k']][0] += p['mw']; gk[p['k']][1] += 1; gk[p['k']][2].add(p['cod'])
tot = sum(p['mw'] for p in centrales)
top = sorted(gk.items(), key=lambda x: -x[1][0])[:10]
out.append('  pais:{mw:%s, centrales:%d, titulares:%d, clientes:%d,' % (
    num(tot), len(centrales), len({p['prop'] for p in centrales}), len(gk)))
out.append('    top:[' + ','.join("['%s',%s,%d,[%s]]" % (
    k, num(v[0]), v[1], ','.join(str(c) for c in CODS if c in v[2]))
    for k, v in top) + ']},')

usados = {k for k, _ in top} | set(GRUPOS)     # todos los grupos: también los que solo tienen obras
reg = defaultdict(list)
for p in centrales:
    reg[p['cod']].append(p)
cuerpo = []
for cod in CODS:
    ps = reg.get(cod, [])
    if not ps:
        cuerpo.append('    %d:{mw:0, centrales:0, clientes:0, top:[], resto:[0,0,0]}' % cod)
        continue
    cl = defaultdict(list)
    for p in ps:
        cl[p['k']].append(p)
    orden = sorted(cl.items(), key=lambda x: -sum(p['mw'] for p in x[1]))
    vis, resto = orden[:TOP], orden[TOP:]
    filas_cl = []
    for k, L in vis:
        usados.add(k)
        L.sort(key=lambda p: -p['mw'])
        mw = sum(p['mw'] for p in L)
        dudoso = sum(p['mw'] for p in L if p['conf'] == 'p') > mw / 2
        filas_cl.append("      ['%s',%s,%s,[%s]]" % (
            k, num(mw), "'p'" if dudoso and L[0]['tipo'] == 'g' else "''",
            ','.join('[%s,%s,%s,%s,%d]' % (js(p['n']), js(p['c']), num(p['mw']),
                                           p['a'] or 0, p['pmgd']) for p in L)))
    rmw = sum(p['mw'] for _, L in resto for p in L)
    cuerpo.append('    %d:{mw:%s, centrales:%d, clientes:%d,\n     top:[\n%s],\n     resto:[%d,%d,%s]}' % (
        cod, num(sum(p['mw'] for p in ps)), len(ps), len(cl), ',\n'.join(filas_cl),
        len(resto), sum(len(L) for _, L in resto), num(rmw)))

out.append('  g:{')
out.append(',\n'.join("    %s:[%s,%s]" % (js(k) if not re.match(r'^[a-z]+$', k) else k,
                                          js(nombres[k][0]), js(nombres[k][1]))
                      for k in sorted(usados, key=lambda k: (k[:2] in ('f:', 's:'), k))))
out.append('  },')
out.append('  reg:{')
out.append(',\n'.join(cuerpo))
out.append('  }')
out.append('};')
out.append('/* <<< CLIENTES */')
bloque = '\n'.join(out)

destino = next((a for a in sys.argv[1:] if not a.startswith('--')), None)
if not destino:
    print(bloque)
    sys.exit()


def reemplazar(html, marca, nuevo):
    ini, fin = html.index('/* >>> ' + marca), html.index('/* <<< %s */' % marca) + len('/* <<< %s */' % marca)
    return html[:ini] + nuevo + html[fin:]


# ---------- plantilla de distribuidores: una fila por región y cliente ----------
html = open(destino, encoding='utf-8').read()
NOMREG = {int(c): n for c, n in re.findall(r'"cod":(\d+),"nombre":"([^"]+)"', html)}

# proyectos del catastro SOLAR que aún no operan y ya tienen cliente (campo k)
OBRAS = defaultdict(lambda: defaultdict(list))
sol = html[html.index('const SOLAR = {'):]
sol, cod = sol[:sol.index('\n};')], None
for linea in sol.splitlines():
    m = re.match(r'\s*(\d+)\s*:\{\s*mw:', linea)
    if m:
        cod = int(m.group(1))
    for o in re.finditer(r"\{n:'((?:[^'\\]|\\.)*)'[^{}]*?mw:([\d.]+)[^{}]*?e:'(\w+)'[^{}]*?k:'([\w:-]+)'\}", linea):
        OBRAS[cod][o.group(4)].append((o.group(1), float(o.group(2)), o.group(3)))
ETAPA = {'co': 'en construcción', 'pr': 'en pruebas', 'rca': 'RCA aprobada', 'ev': 'en calificación'}
COLS = ['id', 'region_cod', 'region', 'ranking', 'en_tarjeta', 'cliente', 'matriz', 'tipo',
        'mw_operacion', 'centrales', 'mw_en_obra', 'centrales_principales', 'distribuidor', 'observaciones']


def coma(x):
    return num(x).replace('.', ',')


def tipo(k, L):
    if k.startswith('f:'):
        return 'Serie de SPV (matriz por identificar)'
    if k.startswith('s:'):
        return 'Sociedad sin grupo identificado'
    mw = sum(p['mw'] for p in L)
    return 'Grupo por confirmar' if sum(p['mw'] for p in L if p['conf'] == 'p') > mw / 2 else 'Grupo'


filas_csv = []
for cod in CODS:
    cl = defaultdict(list)
    for p in reg.get(cod, []):
        cl[p['k']].append(p)
    ob = OBRAS.get(cod, {})
    tarjeta = {k for k, _ in sorted(cl.items(), key=lambda x: -sum(p['mw'] for p in x[1]))[:TOP]} | set(ob)
    op = lambda k: sum(p['mw'] for p in cl.get(k, []))
    obra = lambda k: sum(mw for _, mw, e in ob.get(k, []) if e in ('co', 'pr'))
    for i, k in enumerate(sorted(set(cl) | set(ob), key=lambda k: -(op(k) + obra(k))), 1):
        L = sorted(cl.get(k, []), key=lambda p: -p['mw'])
        princ = ['%s (%s)' % (p['n'], coma(p['mw'])) for p in L[:3]]
        if len(L) > 3:
            princ.append('+%d más' % (len(L) - 3))
        princ += ['%s (%s, %s)' % (n, coma(mw), ETAPA[e]) for n, mw, e in ob.get(k, [])]
        filas_csv.append({'id': '%d:%s' % (cod, k), 'region_cod': cod, 'region': NOMREG[cod],
                          'ranking': i, 'en_tarjeta': 'sí' if k in tarjeta else 'no',
                          'cliente': nombres[k][0], 'matriz': nombres[k][1], 'tipo': tipo(k, L),
                          'mw_operacion': coma(op(k)), 'centrales': len(L), 'mw_en_obra': coma(obra(k)),
                          'centrales_principales': ' · '.join(princ),
                          'distribuidor': '', 'observaciones': ''})
vigentes = {f['id'] for f in filas_csv}


def leer_csv(ruta):
    for enc in ('utf-8-sig', 'cp1252'):          # «CSV UTF-8» o «CSV» a secas de Excel
        try:
            txt = open(ruta, encoding=enc).read()
            break
        except UnicodeDecodeError:
            continue
    cab = txt.splitlines()[0]
    sep = ';' if cab.count(';') >= cab.count(',') else ','
    return {f['id'].strip(): f for f in csv.DictReader(io.StringIO(txt), delimiter=sep) if (f.get('id') or '').strip()}


previo = leer_csv(CSV) if os.path.exists(CSV) else {}
asig = {i: f['distribuidor'].strip() for i, f in previo.items() if (f.get('distribuidor') or '').strip()}

dist = ['/* >>> DISTRIB · cargado desde %s por generar_clientes_solar.py */' % CSV]
if asig:
    dist.append('const DISTRIB = {')
    dist.append(',\n'.join('  %s:%s' % (js(i), js(d)) for i, d in sorted(asig.items(), key=lambda x: (int(x[0].split(':')[0]), x[0]))))
    dist.append('};')
else:
    dist.append('const DISTRIB = {};')
dist.append('/* <<< DISTRIB */')

html = reemplazar(reemplazar(html, 'CLIENTES', bloque), 'DISTRIB', '\n'.join(dist))
open(destino, 'w', encoding='utf-8', newline='').write(html)
print('CLIENTES y DISTRIB actualizados en', destino, '·', len(centrales), 'centrales ·',
      len(asig), 'asignaciones')
for d, n in Counter(asig.values()).most_common():
    print('   %3d  %s' % (n, d))
huerfanos = sorted(set(previo) - vigentes, key=lambda i: (int(i.split(':')[0]), i))
if huerfanos:
    print('Aviso: %d filas del CSV ya no corresponden a un cliente vigente:' % len(huerfanos), ', '.join(huerfanos))

if not previo or '--plantilla' in sys.argv:
    for f in filas_csv:                            # conserva lo que ya se escribió a mano
        if f['id'] in previo:
            f['distribuidor'] = (previo[f['id']].get('distribuidor') or '').strip()
            f['observaciones'] = (previo[f['id']].get('observaciones') or '').strip()
    for i in huerfanos:                            # y no pierde asignaciones de clientes que salieron
        f = previo[i]
        if (f.get('distribuidor') or '').strip() or (f.get('observaciones') or '').strip():
            filas_csv.append({c: f.get(c, '') for c in COLS} | {'tipo': 'Sin centrales en el registro vigente'})
    try:
        with open(CSV, 'w', encoding='utf-8-sig', newline='') as fh:
            w = csv.DictWriter(fh, fieldnames=COLS, delimiter=';')
            w.writeheader()
            w.writerows(filas_csv)
        print(CSV, 'escrito ·', len(filas_csv), 'filas')
    except PermissionError:
        print('No pude escribir %s: ciérralo en Excel y vuelve a correr con --plantilla.' % CSV)
elif len(vigentes - set(previo)):
    print('Hay %d clientes nuevos que no están en %s: corre con --plantilla para sumarlos.' % (
        len(vigentes - set(previo)), CSV))
