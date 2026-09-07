# -*- coding: utf-8 -*-
import openpyxl, re, sys
from collections import defaultdict, Counter
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('Chile_IIR Projects.xlsx', data_only=True)
ws = wb['Export']; rows = list(ws.iter_rows(values_only=True))
hdr = list(rows[0]); idx = {h: i for i, h in enumerate(hdr)}
data = [r for r in rows[1:] if r[idx['Project Name']]]

REG = {'Region Metropolitana':13,'Santiago Metro':13,'Valparaiso':5,'Bio-Bio':8,
       'Coquimbo':4,'Magallanes y de la Antartica Chilena':12,'Atacama':3}

MW = {301203903:19.2,301203925:9.6,301203931:9.6,301203941:9.6,
 301214278:80,301214685:60,301214714:60,301214708:100,301038236:7,301038274:33,
 301169804:4.8,301169815:4.8,301169847:4.8,301169853:4.8,301066808:4.8,301130047:4.8,
 301169307:4.8,301169346:4.8,301169349:4.8,301169350:4.8,300700646:6,300700647:6,
 300851211:2,300725734:4.8,300726124:18,301002994:3.5,301003015:1,
 301091350:45,301091378:45,301090101:10,300962956:10,
 301114874:25,301114929:25,301035882:16,301114405:50,301114926:50,301214747:80,
 300775126:5,300775224:10,300775399:5,300775217:10}

TIPO = {'Grassroot':'gr','Plant Expansion':'ex','Equipment Addition':'eq',
        'Brownfield':'br','Upgrade':'up','Electricity Transmission':'tr'}

ROM = {'I','II','III','IV','V','VI','VII','VIII','IX','X','XI','XII'}
BAJA = {'de','del','la','el','y','los','las','a','en'}

def titulo(s):
    out = []
    for j, w in enumerate(s.split()):
        core = re.sub(r'[^A-Za-z0-9&]', '', w)
        if not core:
            out.append(w); continue
        if any(c.isdigit() for c in core) or core.upper() in ROM or core.upper() in ('MW','DC','IT','DH'):
            out.append(w)
        elif core.lower() in BAJA and j > 0:
            out.append(w.lower())
        else:
            out.append(re.sub(r'[A-Za-z]+', lambda m: m.group(0).capitalize(), w.lower()))
    return ' '.join(out)

TILDES = [('Valparaiso', 'Valparaíso'), ('Quinenco', 'Quiñenco'),
          ('Amunategui', 'Amunátegui'), ('Colina el Pino', 'Colina El Pino'),
          ('Centurylink', 'CenturyLink'), ('Substation And', 'Substation and'),
          ('Data Center NO.2', 'Data Center No. 2'),
          ('Data Center No 2', 'Data Center No. 2'),
          ('Data Center Data Hall Quilicura', 'Data Center · Data Hall')]

def acentos(s):
    for a, b in TILDES:
        s = s.replace(a, b)
    return s

OP = {'Scala Chile Data Centers SpA':'Scala Data Centers',
 'Servicios Amazon Data Service Chile SpA':'Amazon Web Services',
 'Ascenty Chile Spa':'Ascenty (Digital Realty)','DC Terranova SpA':'DC Terranova',
 'Equinix Chile Spa':'Equinix','Odata Chile SA':'Odata (Aligned)',
 'Odata Colocation':'Odata (Aligned)','GR Huina SpA':'GR Huina',
 'EdgeConnex Chile II SpA':'EdgeConneX','EdgeConnex Chile IV SpA':'EdgeConneX',
 'Google Incorporated':'Google','Cirion Technologies Chile SA':'Cirion Technologies',
 'TECfusions Incorporated':'TECfusions','Grupo Energy Service SpA':'Grupo Energy Service',
 'Association of Universities for Research in Astronomy':'AURA (observatorios)',
 'Axxa Group':'Axxa Group',
 'Comunicacion y Telefonia Rural SA':'Comunicación y Telefonía Rural',
 'Microsoft Chile SA':'Microsoft','Huawei Technologies Co Ltd':'Huawei',
 'Alibaba (China) Company Limited':'Alibaba Cloud','Zelestra Chile SAS':'Zelestra'}

# Etapa del ciclo de vida IIR: viene embebida en el Scope como
# "performs <actividad> for <nombre del proyecto>". No hay campo estructurado.
ETAPA = {'Completion':'co', 'Final Commissioning':'fc', 'Construction':'cn',
 'Purchasing':'pu', 'Planning and Scheduling':'pl', 'Permitting':'pe',
 'Detailed Design':'dd', 'Preliminary Engineering':'pi', 'Preliminary Design':'dp',
 'Capital Approval':'ca', 'Project Justification':'ju', 'Project Scope':'al',
 'Market Analysis':'ma'}
BOLSA = {'co':'fin','fc':'fin','cn':'obr','pu':'ing','pl':'ing','pe':'ing','dd':'ing',
 'pi':'est','dp':'est','ca':'est','ju':'est','al':'est','ma':'est','nd':'nd'}
BOLSAS = [('fin','Terminado o en comisionamiento'), ('obr','En construcción'),
 ('ing','Compras, permisos e ingeniería'), ('est','Estudio previo'),
 ('nd','Sin etapa declarada')]
RXET = re.compile(r'performs\s+([A-Za-z ]+?)\s+for\s', re.I)

def etapa(sc):
    m = RXET.search(sc or '')
    return ETAPA.get(m.group(1).strip(), 'nd') if m else 'nd'

COM = {'Valparaiso': 'Valparaíso'}
MES = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']

def fecha(s):
    y, m = s.split('-')
    return '%s-%s' % (MES[int(m) - 1], y)

proys = []
for r in data:
    pid = r[idx['Project #']]
    c = r[idx['Plant City']]
    proys.append(dict(id=pid, cod=REG[r[idx['Plant State']]],
        n=acentos(titulo(r[idx['Project Name']])), c=COM.get(c, c),
        p=OP[r[idx['Project Owner']]],
        i=round((r[idx['Total Investment Value']] or 0) / 1e6, 1),
        mw=MW.get(pid), a=r[idx['Kickoff Date']], e=TIPO[r[idx['Project Type']]],
        s=etapa(r[idx['Scope']])))

TXT = {
 13: ("Concentra casi toda la cartera del país: 71 de los 79 proyectos y el 91% de la inversión anunciada.",
      "Doce comunas del Gran Santiago absorben la totalidad de la cartera regional. Quilicura (12 proyectos), Buin, Huechuraba y Lampa (10 cada una) forman el eje del clúster. El campus de Alto Jahuel de GR Huina —cuatro edificios, 300 MW y US$ 2.100 millones— es la mayor apuesta individual del catastro."),
 3: ("Una sola operación, pero la segunda mayor del país: el hiperescala de Zelestra en el desierto de Atacama.",
     "US$ 600 millones y 80 MW en Tierra Amarilla, con arranque previsto en septiembre de 2028. Es el primer intento de acercar la carga de cómputo al lugar donde se concentra el recurso solar: la región tiene 4.036 MW de ERNC instalada y vertió unos 775 GWh en el primer semestre de 2026."),
 5: ("Cuatro fases de un mismo campus de Scala Data Centers en la ciudad de Valparaíso.",
     "CUR01 a CUR04 suman US$ 122 millones y 30 MW declarados entre 2022 y 2027. Junto a la Metropolitana, es la única región con un campus de fases sucesivas en lugar de una instalación aislada."),
 12: ("Un data center regional de conectividad en Punta Arenas, sin escala hiperescala.",
      "US$ 20 millones con arranque en julio de 2022. Su alcance no declara capacidad en MW: responde a necesidades de conectividad austral más que a demanda de cómputo."),
 4: ("Único proyecto de la región: el centro de datos astronómicos de AURA en La Serena.",
     "US$ 5 millones (2018) para procesar y almacenar la producción de los observatorios del Norte Chico. Es el proyecto más antiguo del catastro y el único que no responde a demanda comercial de cómputo."),
 8: ("Un proyecto pequeño en Coronel, asociado al polo industrial y energético local.",
     "US$ 4 millones con arranque en noviembre de 2026. Es la única presencia del catastro al sur del Maule y su alcance no declara capacidad en MW."),
 15: ("Sin cartera de data centers registrada.",
      "No aparece en el catastro. La demanda de cómputo sigue anclada al eje Santiago–Valparaíso, a más de 2.000 km de aquí."),
 1: ("Sin cartera de data centers registrada.",
     "Sin proyectos en el catastro. El tamaño de la red local y la distancia a los puntos de intercambio de tráfico dejan a la región fuera del mapa."),
 2: ("Sin cartera de data centers registrada, pese a ser la mayor región eléctrica renovable del país.",
     "Concentra el 36% de la capacidad ERNC de Chile —7.596 MW— y el mayor vertimiento nacional, pero no registra un solo proyecto de data center. La carga de cómputo no ha seguido al recurso energético."),
 6: ("Sin cartera de data centers registrada.",
     "Sin proyectos, pese a limitar con el eje metropolitano donde se concentra el 91% de la inversión anunciada."),
 7: ("Sin cartera de data centers registrada.",
     "Sin proyectos en el catastro, aunque su parque ERNC (982 MW) supera al de varias regiones del norte."),
 16: ("Sin cartera de data centers registrada.", "Sin proyectos en el catastro."),
 9: ("Sin cartera de data centers registrada.",
     "Sin proyectos en el catastro. Su matriz renovable es mayoritariamente eólica y de biomasa."),
 14: ("Sin cartera de data centers registrada.", "Sin proyectos en el catastro."),
 10: ("Sin cartera de data centers registrada.",
      "Sin proyectos en el catastro. El data center más austral del país está en Punta Arenas, 1.300 km al sur."),
 11: ("Sin cartera de data centers registrada.",
      "Sin proyectos en el catastro. La región tampoco está conectada al Sistema Eléctrico Nacional."),
}
CODS = [15, 1, 2, 3, 4, 5, 13, 6, 7, 16, 8, 9, 14, 10, 11, 12]

reg = defaultdict(list)
for p in proys:
    reg[p['cod']].append(p)

def num(x):
    return ('%.1f' % x).rstrip('0').rstrip('.') if x % 1 else str(int(x))

def js(s):
    return "'" + s.replace('\\', '\\\\').replace("'", "\\'") + "'"

out = []
out.append('const DC = {')
out.append("  corte:'catastro IIR · arranques 2018-2037',")

tot = round(sum(p['i'] for p in proys), 1)
totmw = round(sum(p['mw'] or 0 for p in proys), 1)
nmw = sum(1 for p in proys if p['mw'])
ser = defaultdict(float)
for p in proys:
    ser[int(p['a'][:4])] += p['i']
serie = [[y, round(ser.get(y, 0), 1)] for y in range(2018, 2032)]
tipos = defaultdict(lambda: [0, 0.0])
for p in proys:
    tipos[p['e']][0] += 1
    tipos[p['e']][1] += p['i']
otros_n = tipos['br'][0] + tipos['up'][0] + tipos['tr'][0]
otros_i = round(tipos['br'][1] + tipos['up'][1] + tipos['tr'][1], 1)
ops = Counter()
for p in proys:
    ops[p['p']] += p['i']

out.append('  pais:{')
out.append('    inv:%s, n:%d, mw:%s, nmw:%d, ops:%d, comunas:%d, regiones:%d, tardio:65,' % (
    tot, len(proys), totmw, nmw, len(ops), len(set(p['c'] for p in proys)), len(reg)))
out.append("    tipos:[['Obra nueva · grassroot',%d,%s],['Ampliación de planta',%d,%s],"
           "['Adición de equipamiento',%d,%s],['Otros · brownfield y transmisión',%d,%s]]," % (
    tipos['gr'][0], round(tipos['gr'][1], 1), tipos['ex'][0], round(tipos['ex'][1], 1),
    tipos['eq'][0], round(tipos['eq'][1], 1), otros_n, otros_i))
et = defaultdict(lambda: [0, 0.0])
for p in proys:
    b = BOLSA[p['s']]
    et[b][0] += 1
    et[b][1] += p['i']
out.append('    etapas:[' + ','.join("['%s','%s',%d,%s]" % (k, n, et[k][0], round(et[k][1], 1))
                                     for k, n in BOLSAS if et[k][0]) + '],')
out.append('    top:[' + ','.join("['%s',%s]" % (k, round(v, 1)) for k, v in ops.most_common(5)) + '],')
out.append('    serie:[' + ','.join('[%d,%s]' % (y, v) for y, v in serie) + ']')
out.append('  },')
out.append('  reg:{')
for cod in CODS:
    ps = sorted(reg.get(cod, []), key=lambda p: (-p['i'], p['a']))
    inv = round(sum(p['i'] for p in ps), 1)
    mw = round(sum(p['mw'] or 0 for p in ps), 1)
    o = Counter()
    for p in ps:
        o[p['p']] += p['i']
    perfil, nota = TXT[cod]
    out.append('    %d:{inv:%s, n:%d, mw:%s, nmw:%d, comunas:%d,' % (
        cod, inv, len(ps), mw, sum(1 for p in ps if p['mw']), len(set(p['c'] for p in ps))))
    eb = defaultdict(lambda: [0, 0.0])
    for p in ps:
        b = BOLSA[p['s']]
        eb[b][0] += 1
        eb[b][1] += p['i']
    out.append('       etapas:{' + ','.join("%s:[%d,%s]" % (k, eb[k][0], round(eb[k][1], 1))
                                            for k, _ in BOLSAS if eb[k][0]) + '},')
    out.append('       ops:[' + ','.join("['%s',%s]" % (k, round(v, 1)) for k, v in o.most_common()) + '],')
    out.append('       perfil:%s,' % js(perfil))
    out.append('       nota:%s,' % js(nota))
    if ps:
        out.append('       obras:[')
        for k, p in enumerate(ps):
            mwtxt = ('mw:%s, ' % num(p['mw'])) if p['mw'] else ''
            out.append('         {n:%s, c:%s, p:%s, i:%s, %sa:%s, e:%s, s:%s}%s' % (
                js(p['n']), js(p['c']), js(p['p']), p['i'], mwtxt, js(fecha(p['a'])), js(p['e']),
                js(p['s']), ',' if k < len(ps) - 1 else ']},'))
    else:
        out.append('       obras:[]},')
out[-1] = out[-1].rstrip(',')
out.append('  }')
out.append('};')
print('\n'.join(out))
