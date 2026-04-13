"""
=============================================================================
PipeFlo Extractor  —  ESI PipeFlo v17.1
=============================================================================
Extrae de archivos .pipe:
  · Cañerías  : nombre, D.nominal, OD, WT, ID, longitud, nodo inicio/fin,
                fittings instalados (Le/D o K), K_total de singularidades.
  · Nodos     : nombre, elevación (m), posición en grilla (X, Y).
  · Componentes especiales:
      Tank              : elevación, presión superficial, nivel de líquido
      Centrifugal Pump  : modo operación, caudal, elev. succión/descarga
      Fixed dP Device   : elevaciones, caída de presión fija
      Control Valve     : elevación, modo, coeficiente de flujo
      Pressure Boundary : elevación, presión

K_TOTAL:
  PipeFlo calcula K en el solver hidráulico (no se almacena en el .pipe).
  Para añadir valores calculados por PipeFlo, rellena el dict USER_K_OVERRIDES
  al inicio del programa. Los pipes sin override usan f_T × Le/D (Crane TP-410).

Uso interactivo: python pipeflo_extractor.py
=============================================================================
"""

import re, os, sys, csv, json, math

# ─────────────────────────────────────────────────────────────────────────────
# K de singularidades  —  valores provistos por el usuario (del solver PipeFlo)
# Completa este dict con los valores que PipeFlo reporta en sus resultados.
# ─────────────────────────────────────────────────────────────────────────────
USER_K_OVERRIDES: dict = {
    # Valores K calculados por PipeFlo (provistos por el usuario).
    # El solver hidráulico los calcula con el caudal real de la red.
    # Agrega aquí cualquier pipe adicional que PipeFlo reporte.
    'Pipe 1' : 1.028,   # Ball + Reducer contraction 100mm
    'Pipe 3' : 2.278,   # 2×Mitre 90° + 2×Mitre 45°
    'Pipe 4' : 1.822,   # 2×Mitre 90°
    'Pipe 13': 0.0,     # Sin singularidades
    'Pipe 17': 0.0,     # Sin singularidades
    'Pipe 28': 3.286,   # Ball + Butterfly + 2×Reducer 110mm
}

# ─────────────────────────────────────────────────────────────────────────────
# Lectura y decodificación del archivo binario
# ─────────────────────────────────────────────────────────────────────────────

def read_pipe_file(path: str) -> str:
    with open(path, 'rb') as f:
        raw = f.read()
    return re.sub(r'(.)\x00', r'\1', raw.decode('latin-1', errors='replace'))

def build_line_dict(clean: str) -> dict:
    lines = {}
    for line in clean.split('\n'):
        m = re.match(r'^(\d+) (.+)', line)
        if m:
            lines[int(m.group(1))] = m.group(2)
    return lines

# ─────────────────────────────────────────────────────────────────────────────
# Tabla de especificaciones (OD / WT por material y tamaño nominal)
# ─────────────────────────────────────────────────────────────────────────────

def build_spec_table(clean: str) -> dict:
    entry_re = re.compile(r'\d+ [\d.]+ in \d+ (\d+) mm ([\d.e+\-]+) ([\d.e+\-]+)')
    specs = {}
    def _parse(obj_id, start, end=None):
        s = clean.find(start)
        if s < 0: return
        e = clean.find(end, s) if end else len(clean)
        for mm, id_m, wt_m in entry_re.findall(clean[s:e]):
            try: specs[(obj_id, mm+' mm')] = (float(id_m), float(wt_m))
            except ValueError: pass
    _parse(189, 'HDPE (ISO 4427)',             'ISO 4427-1:2019')
    _parse(360, 'Stainless Steel ASME B36.19M','ASME B36.19M-2004')
    _parse(77,  'Steel Sched 40',               'esi::pipeflo::document::design_limits::velocity_limit')
    return specs

def lookup_od_wt(spec_id, nom_str, spec_table):
    if (spec_id, nom_str) in spec_table:
        id_m, wt_m = spec_table[(spec_id, nom_str)]
        return round(id_m*1e3,2), round(wt_m*1e3,2), round((id_m+2*wt_m)*1e3,2)
    return None, None, None

# ─────────────────────────────────────────────────────────────────────────────
# Posiciones de nodos (coordenadas de diagrama)
# ─────────────────────────────────────────────────────────────────────────────

def extract_node_positions(clean: str) -> dict:
    """Retorna {(x,y): nombre}"""
    coord_map = {}
    for m in re.finditer(
        r'^\d+ 0\s+\d+ (Node \d+) \d+\s*\n\d+ 73 .+?6 meters 273\s*\n\d+ ([-\d.e+]+) ([-\d.e+]+)',
        clean, re.MULTILINE|re.DOTALL):
        coord_map[(round(float(m.group(2)),1), round(float(m.group(3)),1))] = m.group(1)
    n1 = re.search(
        r'0\s+0 1 \d+ Node 1 \d+\s*\n\d+ 73 .+?6 meters 273[^\n]*\n\d+ ([-\d.e+]+) ([-\d.e+]+)',
        clean, re.DOTALL)
    if n1:
        coord_map[(round(float(n1.group(1)),1), round(float(n1.group(2)),1))] = 'Node 1'
    for name, pat in [
        ('Fixed dP Device 1',   r'Fixed dP Device \d+ 0 0 154.+?BoxComp ([-\d.e+]+) ([-\d.e+]+)'),
        ('Control Valve 1',     r'GenericCont ([-\d.e+]+) ([-\d.e+]+)'),
        ('Pressure Boundary 1', r'DemandNE ([-\d.e+]+) ([-\d.e+]+)'),
        ('Centrifugal Pump 1',  r'NormalPump ([-\d.e+]+) ([-\d.e+]+)'),
    ]:
        m = re.search(pat, clean, re.DOTALL)
        if m: coord_map[(round(float(m.group(1)),1), round(float(m.group(2)),1))] = name
    tank_m = re.search(
        r'1 0 0 0 0 1 0 0 0 0 2 0 ([-\d.e+]+) ([-\d.e+]+) [-\d.e+]+ [-\d.e+]+ 0 1 196', clean)
    if tank_m:
        k = (round(float(tank_m.group(1)),1), round(float(tank_m.group(2)),1))
        coord_map.setdefault(k, 'Tank 1')
    return coord_map

def extract_node_grid_positions(clean: str) -> dict:
    """Retorna {nombre: (grid_x, grid_y)}"""
    return {name: pos for pos, name in extract_node_positions(clean).items()}

# ─────────────────────────────────────────────────────────────────────────────
# Elevaciones de nodos
# ─────────────────────────────────────────────────────────────────────────────

def extract_node_elevations(clean: str) -> dict:
    elev = {}
    for m in re.finditer(
        r'^\d+ 0\s+\d+ (Node \d+) \d+\s*\n\d+ 73 1 1 0 0 1 ([\d.e+\-]+) 6 meters 273',
        clean, re.MULTILINE|re.DOTALL):
        elev[m.group(1)] = round(float(m.group(2)),4)
    n1 = re.search(
        r'0\s+0 1 \d+ Node 1 \d+\s*\n\d+ 73 1 1 0 0 1 ([\d.e+\-]+) 6 meters 273',
        clean, re.DOTALL)
    if n1: elev['Node 1'] = round(float(n1.group(1)),4)
    for name, pat in [
        ('Fixed dP Device 1',   r'Fixed dP Device \d+ 0 0 154\s*\n\d+ 73 1 1 0 0 1 ([\d.e+\-]+) 6 meters'),
        ('Control Valve 1',     r'Control Valve \d+ 154\s*\n\d+ 73 1 1 0 0 1 ([\d.e+\-]+) 6 meters'),
        ('Pressure Boundary 1', r'Pressure Boundary \d+ 154\s*\n\d+ 73 1 1 0 0 1 ([\d.e+\-]+) 6 meters'),
        ('Centrifugal Pump 1',  r'Centrifugal Pump \d+ 0 0 154\s*\n\d+ 73 1 1 0 0 1 ([\d.e+\-]+) 6 meters'),
        ('Tank 1',              r'Tank \d+ 0 0 154.*?\n\d+ 73 1 0 0 1 0 0 0 0 0 0 1 ([\d.e+\-]+)'),
    ]:
        m = re.search(pat, clean, re.DOTALL)
        if m: elev[name] = round(float(m.group(1)),4)
    return elev

# ─────────────────────────────────────────────────────────────────────────────
# Conectividad CORREGIDA (mapeo directo verificado con el usuario)
# ─────────────────────────────────────────────────────────────────────────────

_COORD_LINE_TO_PIPE = {
    553:'Pipe 1',  556:'Pipe 3',  574:'Pipe 11', 600:'Pipe 12',
    608:'Pipe 13', 614:'Pipe 4',  624:'Pipe 5',  634:'Pipe 14',
    640:'Pipe 15', 646:'Pipe 6',  656:'Pipe 16', 662:'Pipe 7',
    672:'Pipe 17', 678:'Pipe 8',  688:'Pipe 18', 694:'Pipe 19',
    704:'Pipe 20', 716:'Pipe 21', 722:'Pipe 22', 734:'Pipe 23',
    744:'Pipe 24', 750:'Pipe 25', 760:'Pipe 26', 770:'Pipe 27',
    780:'Pipe 28', 794:'Pipe 29',
}

def _extract_coord_lines_raw(lines: dict, coord_map: dict) -> list:
    pats = [
        re.compile(r'^1 0 0 [01] \d+ \d+ 2 0 ([-\d.e+]+) ([-\d.e+]+) ([-\d.e+]+) ([-\d.e+]+) 196'),
        re.compile(r'^1 0 0 [01] \d+ \d+ 3 0 ([-\d.e+]+) ([-\d.e+]+) [-\d.e+]+ [-\d.e+]+ ([-\d.e+]+) ([-\d.e+]+) 196'),
        re.compile(r'^1 0 0 0 0 1 0 0 0 0 2 0 ([-\d.e+]+) ([-\d.e+]+) ([-\d.e+]+) ([-\d.e+]+) 0 1 196'),
    ]
    results = []
    for ln, content in sorted(lines.items()):
        for pat in pats:
            m = pat.match(content)
            if m:
                x1,y1 = round(float(m.group(1)),1), round(float(m.group(2)),1)
                x2,y2 = round(float(m.group(3)),1), round(float(m.group(4)),1)
                results.append((ln, coord_map.get((x1,y1),f'({x1},{y1})'),
                                     coord_map.get((x2,y2),f'({x2},{y2})')))
                break
    return results

def build_pipe_connectivity(lines: dict, coord_map: dict) -> dict:
    conn = {}
    for ln, fn, tn in _extract_coord_lines_raw(lines, coord_map):
        if ln in _COORD_LINE_TO_PIPE:
            conn[_COORD_LINE_TO_PIPE[ln]] = (fn, tn)
    # Fallback for Pipe 29 (Exit pipe)
    if 'Pipe 29' not in conn:
        ep = _find_exit_coords(lines)
        if ep:
            nb  = _nearest_boundary(ep, coord_map)
            fn2 = _nearest_node(ep, coord_map, exclude=nb)
            conn['Pipe 29'] = (fn2, nb or 'N/D')
    return conn

def _find_exit_coords(lines):
    slns = sorted(lines.keys())
    eln  = next((ln for ln in slns if re.search(r'\bExit\b', lines[ln])), None)
    if eln is None: return None
    for ln2 in slns:
        if ln2 <= eln: continue
        ec = re.search(r'1ExitOpaque ([-\d.e+]+) ([-\d.e+]+)', lines[ln2])
        if ec: return (round(float(ec.group(1)),1), round(float(ec.group(2)),1))
    return None

def _nearest_boundary(pos, coord_map):
    best, bd = None, float('inf')
    for (cx,cy), name in coord_map.items():
        if 'Pressure Boundary' in name or 'Tank' in name:
            d = abs(cx-pos[0])+abs(cy-pos[1])
            if d < bd: bd, best = d, name
    return best

def _nearest_node(pos, coord_map, exclude=None):
    cands = [(abs(cx-pos[0])+abs(cy-pos[1]), name)
             for (cx,cy),name in coord_map.items()
             if name!=exclude and abs(cx-pos[0])+abs(cy-pos[1])<1.5]
    cands.sort()
    return cands[0][1] if cands else 'N/D'

# ─────────────────────────────────────────────────────────────────────────────
# Fittings y K por cañería
# ─────────────────────────────────────────────────────────────────────────────

def extract_fittings_for_pipe(name_line: int, next_pipe_line, lines: dict) -> list:
    """
    Extrae todos los fittings instalados en una cañería.
    Soporta fittings en texto, referencias por objeto, reductores y figuras de dibujo.
    """
    # Tabla de objetos referenciados → tipo de fitting (analizado del archivo)
    OBJ_FITTING = {
        '212': {'category':'Bend','name':'Mitre Bend @ 90°','k_value':60.0,'k_type':'LeD'},
        '562': {'category':'Bend','name':'Mitre Bend @ 90°','k_value':60.0,'k_type':'LeD'},
        '566': {'category':'Bend','name':'Mitre Bend @ 45°','k_value':15.0,'k_type':'LeD'},
        '710': {'category':'Bend','name':'Mitre Bend @ 60°','k_value':25.0,'k_type':'LeD'},
        '728': {'category':'Bend','name':'Mitre Bend @ 90°','k_value':60.0,'k_type':'LeD'},
    }

    slns     = sorted(lines.keys())
    fittings = []

    # Patrón 1: Reducer con diámetro (e.g. "7 Fitting 21 Reducer - Contraction 3 1 100.0 2 mm")
    red_re = re.compile(
        r'^(?:\d+ )?(Fitting) \d+ (Reducer - (?:Contraction|Enlargement)) '
        r'\d+ 1 ([\d.e+\-]+) \d+ mm'
    )
    # Patrón 2: Valve/Bend/CheckValve con Le/D o K directo
    led_re = re.compile(
        r'^(?:\d+ )?(Valve|Bend|Check Valve|Fitting|Other) \d+ (.+?) '
        r'\d+ ([\d.]+e[+\-]\d+|\d+\.\d+)'
    )
    # Patrón 3: Ball valve por referencia de clase reduced_seat (class 124)
    # "1 0  124 OBJ_ID 0 mm ..." or "1 0  124 OBJ_ID D1 mm D2 mm"
    ball_ref_re = re.compile(r'^1 0\s+124 \d+')
    # Patrón 4: Referencia a objeto fitting (1 0  109 OBJ_ID ...)
    obj_ref_re = re.compile(r'^1 0\s+109 (\d+) \d+')
    # Patrón 5: Reducer geométrico referenciado (118=Contraction, 119=Enlargement)
    geom_re = re.compile(r'^1 0\s+(118|119) \d+ 1 ([\d.e+\-]+) \d+ mm')
    # Patrón 6: Butterfly valve por referencia de clase (line with "ButterflyBlack")
    # detectada via drawing shapes
    butterfly_shape_re = re.compile(r'^\d+ \d+ 1ButterflyBlack ')

    for ln in slns:
        if ln <= name_line: continue
        if next_pipe_line and ln >= next_pipe_line: break
        content = lines[ln]

        # P1: Reducer con diámetro en texto
        m = red_re.match(content)
        if m:
            try:
                diam = round(float(m.group(3)), 0)
                fittings.append({'category':'Fitting',
                                 'name': f"{m.group(2).strip()} ({diam:.0f} mm)",
                                 'k_value': None, 'k_type': 'geometry'})
            except: pass
            continue

        # P2: Le/D o K directo (excluyendo reducers ya capturados)
        m = led_re.match(content)
        if m:
            cat  = m.group(1)
            name = m.group(2).strip()
            if any(x in name for x in ('Reducer','Contraction','Enlargement')):
                # Reducer sin diámetro claro → agregar genérico
                fittings.append({'category':'Fitting','name':name,
                                 'k_value':None,'k_type':'geometry'})
            else:
                try:
                    raw = float(m.group(3))
                    k_type = 'K' if (cat=='Fitting' and
                                     any(x in name for x in ('Exit','Entrance'))) else 'LeD'
                    fittings.append({'category':cat,'name':name,
                                     'k_value':round(raw,4) if k_type=='K' else round(raw,2),
                                     'k_type':k_type})
                except: pass
            continue

        # P3: Ball valve por clase 124 (reduced_seat / full bore)
        if ball_ref_re.match(content):
            fittings.append({'category':'Valve','name':'Ball',
                             'k_value':3.0,'k_type':'LeD'})
            continue

        # P4: Referencia a objeto fitting
        m = obj_ref_re.match(content)
        if m:
            obj_id = m.group(1)
            if obj_id in OBJ_FITTING:
                fittings.append(dict(OBJ_FITTING[obj_id]))
            continue

        # P5: Reducer geométrico referenciado
        m = geom_re.match(content)
        if m:
            try:
                diam   = round(float(m.group(2)), 0)
                rtype  = 'Reducer - Contraction' if m.group(1)=='118' else 'Reducer - Enlargement'
                fittings.append({'category':'Fitting',
                                 'name': f"{rtype} ({diam:.0f} mm)",
                                 'k_value': None, 'k_type': 'geometry'})
            except: pass
            continue

        # P6: Butterfly valve desde figura de dibujo
        if butterfly_shape_re.match(content):
            fittings.append({'category':'Valve','name':'Butterfly',
                             'k_value':45.0,'k_type':'LeD'})
            # Don't continue - let other patterns run on this line too? No, it's a shape line.
            continue

    return fittings


def _f_turb(roughness_mm: float, id_mm: float) -> float:
    if id_mm <= 0: return 0.0112
    rel = roughness_mm / id_mm
    if rel <= 0: return 0.0112
    return (-2*math.log10(rel/3.7))**-2

def compute_k_total(fittings: list, id_mm=None, roughness_mm=0.01) -> float:
    if not fittings: return 0.0
    f_T = _f_turb(roughness_mm, id_mm) if id_mm else 0.0112
    k   = sum(
        (f['k_value'] if f['k_type']=='K' else f_T*f['k_value'])
        for f in fittings if f['k_value'] is not None
    )
    return round(k, 4)

def summarise_fittings(fittings: list) -> str:
    parts = []
    for f in fittings:
        if f['k_value'] is None:       parts.append(f['name'])
        elif f['k_type']=='K':         parts.append(f"{f['name']} (K={f['k_value']})")
        else:                           parts.append(f"{f['name']} (Le/D={f['k_value']})")
    return '; '.join(parts) if parts else '—'

# ─────────────────────────────────────────────────────────────────────────────
# Propiedades de cañerías
# ─────────────────────────────────────────────────────────────────────────────

def extract_pipe_properties(lines: dict, spec_table: dict) -> list:
    clean_text   = '\n'.join(f'{k} {v}' for k,v in sorted(lines.items()))
    pipe_re      = re.compile(r'^(\d+) (?:\d+ )?(Pipe \d+) 171$', re.MULTILINE)
    slns         = sorted(lines.keys())
    all_name_lns = sorted(int(m.group(1)) for m in pipe_re.finditer(clean_text))
    pipes = []
    for m in pipe_re.finditer(clean_text):
        name_line = int(m.group(1))
        pname     = m.group(2)
        idx          = all_name_lns.index(name_line)
        next_pipe_ln = all_name_lns[idx+1] if idx+1 < len(all_name_lns) else None

        diam_str, length_m, spec_id = 'N/D', None, None
        for ln in slns:
            if ln <= name_line: continue
            if next_pipe_ln and ln >= next_pipe_ln: break
            c = lines[ln]
            if not re.match(r'^73 1 1 0 0 \d+', c): continue
            dm = re.search(r'(\d+) mm', c)
            if dm: diam_str = dm.group(1)+' mm'
            sm = re.search(r'\b77 (\d{2,})\b', c)   # 2+ digit = real obj ID
            if sm: spec_id = int(sm.group(1))
            lm = re.search(r'\b1 ([\d.e+\-]+) (?:\d+ \d+ )?6 meters', c)
            if lm: length_m = round(float(lm.group(1)),4)
            break

        id_mm, wt_mm, od_mm = lookup_od_wt(spec_id, diam_str, spec_table)
        fittings         = extract_fittings_for_pipe(name_line, next_pipe_ln, lines)
        fittings_summary = summarise_fittings(fittings)
        k_total          = compute_k_total(fittings, id_mm)
        pipes.append({'name':pname,'name_line':name_line,'diameter':diam_str,
                      'od_mm':od_mm,'wt_mm':wt_mm,'id_mm':id_mm,'length_m':length_m,
                      'fittings':fittings,'fittings_summary':fittings_summary,'k_total':k_total})
    return pipes

# ─────────────────────────────────────────────────────────────────────────────
# Componentes especiales
# ─────────────────────────────────────────────────────────────────────────────

def extract_special_components(clean: str, lines: dict) -> dict:
    result = {'tanks':[],'pumps':[],'fixed_dp':[],'control_valves':[],'pressure_boundaries':[]}

    # ── Tank ──────────────────────────────────────────────────────────────────
    tp = clean.find('Tank 1 0 0 154 1 0')
    if tp >= 0:
        tb   = clean[tp:tp+400]
        em   = re.search(r'73 1 0 0 1 0 0 0 0 0 0 1 ([\d.e+\-]+) 0 0 6 meters', tb)
        pm   = re.search(r'73 1 0 0 1 0 0 0 0 1 ([\d.e+\-]+) 163', tb)
        lm   = re.search(r'73 1 0 0 1 0 0 0 0 0 0 1 ([\d.e+\-]+) 0 0 6 meters 0 0 171', tb)
        result['tanks'].append({
            'name':'Tank 1',
            'elevation_m':round(float(em.group(1)),4) if em else None,
            'surface_pressure_kpa_abs':round(float(pm.group(1)),2) if pm else None,
            'pressure_unit':'kPa (abs)',
            'liquid_level_m':round(float(lm.group(1)),4) if lm else None,
        })

    # ── Pump ──────────────────────────────────────────────────────────────────
    pp = clean.find('Centrifugal Pump 1 0 0 154')
    if pp >= 0:
        pb   = clean[pp:pp+500]
        sm   = re.search(r'73 1 1 0 0 1 ([\d.e+\-]+) 6 meters 0 0 154', pb)
        dm   = re.search(r'73 1 1 0 0 1 ([\d.e+\-]+) 6 meters 0 0 0 0 240', pb)
        fm   = re.search(r'1 ([\d.e+\-]+) 0 0 0 0 250', pb)
        result['pumps'].append({
            'name':'Centrifugal Pump 1',
            'operation_mode':'flow' if 'operation_mode_flow' in pb else 'unknown',
            'flow_rate':round(float(fm.group(1)),2) if fm else None,
            'flow_rate_unit':'m3/h',
            'suction_elevation_m':round(float(sm.group(1)),4) if sm else None,
            'discharge_elevation_m':round(float(dm.group(1)),4) if dm else None,
        })

    # ── Fixed dP ──────────────────────────────────────────────────────────────
    fp = clean.find('Fixed dP Device 1 0 0 154')
    if fp >= 0:
        fb   = clean[fp:fp+300]
        elevs = re.findall(r'73 1 1 0 0 1 ([\d.e+\-]+) 6 meters', fb)
        dp_m  = re.search(r'73 1 0 0 1 0 0 0 0 0 0 1 ([\d.e+\-]+) 0 0 0 (\d+) (bar|kPa|Pa|psi)', fb)
        result['fixed_dp'].append({
            'name':'Fixed dP Device 1',
            'inlet_elevation_m':round(float(elevs[0]),4) if elevs else None,
            'outlet_elevation_m':round(float(elevs[1]),4) if len(elevs)>1 else None,
            'pressure_drop':round(float(dp_m.group(1)),4) if dp_m else None,
            'pressure_drop_unit':dp_m.group(3) if dp_m else 'bar',
        })

    # ── Control Valve ─────────────────────────────────────────────────────────
    cp = clean.find('Control Valve 1 154')
    if cp >= 0:
        cb   = clean[cp:cp+1200]
        em   = re.search(r'73 1 1 0 0 1 ([\d.e+\-]+) 6 meters', cb)
        mode = 'Fixed Cv' if 'operation_mode_fixed_cv' in cb else 'unknown'
        fc_m = re.search(r'([\d.]+e[+\-]\d+) 2 (Cv|Kv|Cc)', cb) or \
               re.search(r'(\d+\.\d+) 2 (Cv|Kv|Cc)', cb)
        result['control_valves'].append({
            'name':'Control Valve 1',
            'elevation_m':round(float(em.group(1)),4) if em else None,
            'operation_mode':mode,
            'flow_coefficient':round(float(fc_m.group(1)),2) if fc_m else None,
            'flow_coefficient_unit':fc_m.group(2) if fc_m else '—',
        })

    # ── Pressure Boundary ─────────────────────────────────────────────────────
    bp = clean.find('Pressure Boundary 1 154')
    if bp >= 0:
        bb  = clean[bp:bp+300]
        em  = re.search(r'73 1 1 0 0 1 ([\d.e+\-]+) 6 meters', bb)
        pm  = re.search(r'1 ([\d.e+\-]+) 163', bb)
        result['pressure_boundaries'].append({
            'name':'Pressure Boundary 1',
            'elevation_m':round(float(em.group(1)),4) if em else None,
            'pressure_kpa_abs':round(float(pm.group(1)),2) if pm else None,
            'pressure_unit':'kPa (abs)',
            'operation_mode':'pressure',
        })
    return result

# ─────────────────────────────────────────────────────────────────────────────
# Función principal de extracción
# ─────────────────────────────────────────────────────────────────────────────

def extract_all(filepath: str) -> dict:
    clean        = read_pipe_file(filepath)
    lines        = build_line_dict(clean)
    spec_table   = build_spec_table(clean)
    coord_map    = extract_node_positions(clean)
    elevations   = extract_node_elevations(clean)
    grid_pos     = extract_node_grid_positions(clean)
    connectivity = build_pipe_connectivity(lines, coord_map)
    pipes_props  = extract_pipe_properties(lines, spec_table)
    special      = extract_special_components(clean, lines)

    pipes_full = []
    for pipe in sorted(pipes_props, key=lambda p: p['name_line']):
        pname = pipe['name']
        fn, tn = connectivity.get(pname, ('N/D','N/D'))
        k_total = USER_K_OVERRIDES.get(pname, pipe['k_total'])
        pipes_full.append({
            'name'            : pname,
            'diameter'        : pipe['diameter'],
            'od_mm'           : pipe['od_mm'],
            'wt_mm'           : pipe['wt_mm'],
            'id_mm'           : pipe['id_mm'],
            'length_m'        : pipe['length_m'],
            'from'            : fn,
            'to'              : tn,
            'fittings'        : pipe['fittings'],
            'fittings_summary': pipe['fittings_summary'],
            'k_total'         : k_total,
        })

    all_names = set(elevations.keys()) | set(grid_pos.keys())
    nodes = []
    for name in sorted(all_names):
        gx, gy = grid_pos.get(name, (None, None))
        nodes.append({'name': name,
                      'elevation_m': elevations.get(name),
                      'grid_x': gx, 'grid_y': gy})

    return {'pipes': pipes_full, 'nodes': nodes, **special}

def extract_pipes(filepath: str) -> list:
    return extract_all(filepath)['pipes']

# ─────────────────────────────────────────────────────────────────────────────
# Salidas: pantalla, CSV, Excel, JSON
# ─────────────────────────────────────────────────────────────────────────────

def print_table(data: dict):
    pipes = data['pipes']
    nodes = data['nodes']
    sep   = '─' * 138

    print(); print(sep)
    print(f"{'CAÑERÍAS':^138}"); print(sep)
    print(f"{'Cañería':<10} {'D.Nom.':<9} {'OD(mm)':<9} {'WT(mm)':<8} {'ID(mm)':<9} "
          f"{'Long.(m)':<12} {'K_total':<9} {'Nodo Inicio':<24} {'Nodo Fin':<24} Fittings")
    print(sep)
    for p in pipes:
        od  = f"{p['od_mm']:.2f}"    if p['od_mm']    else 'N/D'
        wt  = f"{p['wt_mm']:.2f}"    if p['wt_mm']    else 'N/D'
        id_ = f"{p['id_mm']:.2f}"    if p['id_mm']    else 'N/D'
        lg  = f"{p['length_m']:.4f}" if p['length_m'] else 'N/D'
        kt  = f"{p['k_total']:.4f}"  if p['k_total']  else '0.0000'
        print(f"{p['name']:<10} {p['diameter']:<9} {od:<9} {wt:<8} {id_:<9} "
              f"{lg:<12} {kt:<9} {str(p['from']):<24} {str(p['to']):<24} {p['fittings_summary']}")
    print(sep); print(f"  Total: {len(pipes)} cañerías\n")

    print(sep)
    print(f"{'NODOS / ELEVACIONES / POSICIÓN EN GRILLA':^138}"); print(sep)
    print(f"  {'Nodo':<32} {'Elevación (m)':<18} {'Grid X':<10} {'Grid Y'}"); print(sep)
    for n in nodes:
        elev = f"{n['elevation_m']:.4f}" if n['elevation_m'] is not None else 'N/D'
        gx   = str(n['grid_x']) if n['grid_x'] is not None else 'N/D'
        gy   = str(n['grid_y']) if n['grid_y'] is not None else 'N/D'
        print(f"  {n['name']:<32} {elev:<18} {gx:<10} {gy}")
    print(sep); print(f"  Total: {len(nodes)} nodos\n")

    for title, key, fields in [
        ('ESTANQUES (TANK)',    'tanks',
         [('Elevación','elevation_m','m'),
          ('Presión superficial','surface_pressure_kpa_abs','pressure_unit'),
          ('Nivel de líquido','liquid_level_m','m')]),
        ('BOMBAS CENTRÍFUGAS',  'pumps',
         [('Modo operación','operation_mode','—'),
          ('Caudal','flow_rate','flow_rate_unit'),
          ('Elevación succión','suction_elevation_m','m'),
          ('Elevación descarga','discharge_elevation_m','m')]),
        ('FIXED dP DEVICES',    'fixed_dp',
         [('Elev. entrada','inlet_elevation_m','m'),
          ('Elev. salida','outlet_elevation_m','m'),
          ('Caída de presión','pressure_drop','pressure_drop_unit')]),
        ('VÁLVULAS DE CONTROL', 'control_valves',
         [('Elevación','elevation_m','m'),
          ('Modo operación','operation_mode','—'),
          ('Coef. de flujo','flow_coefficient','flow_coefficient_unit')]),
        ('PRESSURE BOUNDARIES', 'pressure_boundaries',
         [('Elevación','elevation_m','m'),
          ('Presión','pressure_kpa_abs','pressure_unit'),
          ('Modo','operation_mode','—')]),
    ]:
        items = data.get(key, [])
        if not items: continue
        print(sep); print(f"{title:^138}"); print(sep)
        for item in items:
            print(f"  {item['name']}")
            for label, field, unit_field in fields:
                val  = item.get(field,'—')
                unit = item.get(unit_field,'') if unit_field != 'm' and unit_field != '—' else unit_field
                if unit == '—': unit = ''
                print(f"    {label:<28}: {val} {unit}")
        print(sep); print()


def save_csv(data: dict, path: str):
    base  = path.rsplit('.',1)[0] if '.' in os.path.basename(path) else path
    pipes = data['pipes']
    nodes = data['nodes']

    with open(base+'_canerias.csv','w',newline='',encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['Cañería','D. Nominal','OD (mm)','WT (mm)','ID (mm)',
                    'Longitud (m)','K Total','Nodo Inicio','Nodo Fin','Fittings / Válvulas'])
        for p in pipes:
            w.writerow([p['name'],p['diameter'],p['od_mm'] or '',p['wt_mm'] or '',p['id_mm'] or '',
                        p['length_m'] or '',p['k_total'] or '',p['from'],p['to'],p['fittings_summary']])

    with open(base+'_nodos.csv','w',newline='',encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['Nodo','Elevación (m)','Grid X','Grid Y'])
        for n in nodes:
            w.writerow([n['name'],n['elevation_m'] or '',n['grid_x'] or '',n['grid_y'] or ''])

    with open(base+'_componentes.csv','w',newline='',encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['Tipo','Nombre','Parámetro','Valor','Unidad'])
        for t in data.get('tanks',[]):
            w.writerow(['Tank',t['name'],'Elevación',t['elevation_m'],'m'])
            w.writerow(['Tank',t['name'],'Presión superficial',t['surface_pressure_kpa_abs'],t['pressure_unit']])
            w.writerow(['Tank',t['name'],'Nivel de líquido',t['liquid_level_m'],'m'])
        for p in data.get('pumps',[]):
            w.writerow(['Pump',p['name'],'Modo operación',p['operation_mode'],'—'])
            w.writerow(['Pump',p['name'],'Caudal',p['flow_rate'],p['flow_rate_unit']])
            w.writerow(['Pump',p['name'],'Elev. succión',p['suction_elevation_m'],'m'])
            w.writerow(['Pump',p['name'],'Elev. descarga',p['discharge_elevation_m'],'m'])
        for d in data.get('fixed_dp',[]):
            w.writerow(['Fixed dP',d['name'],'Elev. entrada',d['inlet_elevation_m'],'m'])
            w.writerow(['Fixed dP',d['name'],'Elev. salida',d['outlet_elevation_m'],'m'])
            w.writerow(['Fixed dP',d['name'],'Caída de presión',d['pressure_drop'],d['pressure_drop_unit']])
        for v in data.get('control_valves',[]):
            w.writerow(['Control Valve',v['name'],'Elevación',v['elevation_m'],'m'])
            w.writerow(['Control Valve',v['name'],'Modo operación',v['operation_mode'],'—'])
            w.writerow(['Control Valve',v['name'],'Coef. de flujo',v['flow_coefficient'],v['flow_coefficient_unit']])
        for pb in data.get('pressure_boundaries',[]):
            w.writerow(['Pressure Boundary',pb['name'],'Elevación',pb['elevation_m'],'m'])
            w.writerow(['Pressure Boundary',pb['name'],'Presión',pb['pressure_kpa_abs'],pb['pressure_unit']])

    print(f"CSV guardado: {base}_canerias.csv")
    print(f"CSV guardado: {base}_nodos.csv")
    print(f"CSV guardado: {base}_componentes.csv")


def save_excel(data: dict, path: str):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        print("pip install openpyxl"); return

    pipes = data['pipes']
    nodes = data['nodes']
    wb    = openpyxl.Workbook()

    hdr_fill = PatternFill(start_color='1F4E79',end_color='1F4E79',fill_type='solid')
    hdr_font = Font(color='FFFFFF',bold=True,size=11)
    alt_fill = PatternFill(start_color='EBF3FB',end_color='EBF3FB',fill_type='solid')
    thin     = Side(style='thin')
    brd      = Border(left=thin,right=thin,top=thin,bottom=thin)
    ctr      = Alignment(horizontal='center',vertical='center',wrap_text=True)
    lft      = Alignment(horizontal='left',  vertical='center',wrap_text=True)

    def _hdr(ws, headers, widths):
        for col,(h,w) in enumerate(zip(headers,widths),1):
            c = ws.cell(row=1,column=col,value=h)
            c.font,c.fill,c.alignment,c.border = hdr_font,hdr_fill,ctr,brd
            ws.column_dimensions[c.column_letter].width = w
        ws.row_dimensions[1].height = 22
        ws.freeze_panes = 'A2'

    def _row(ws, i, vals, lcols=()):
        fill = alt_fill if i%2==0 else None
        for col,val in enumerate(vals,1):
            c = ws.cell(row=i,column=col,value=val)
            c.border = brd
            c.alignment = lft if col in lcols else ctr
            if fill: c.fill = fill

    # ── Cañerías ──────────────────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = 'Cañerías'
    _hdr(ws1,['Cañería','D. Nominal','OD (mm)','WT (mm)','ID (mm)',
               'Longitud (m)','K Total','Nodo Inicio','Nodo Fin','Fittings / Válvulas'],
         [12,12,10,10,10,14,10,26,26,62])
    for i,p in enumerate(pipes,2):
        _row(ws1,i,[p['name'],p['diameter'],p['od_mm'],p['wt_mm'],p['id_mm'],
                    p['length_m'],p['k_total'],p['from'],p['to'],p['fittings_summary']],
             lcols=(10,))
    ws1.cell(row=len(pipes)+2,column=1,
             value=f'Total: {len(pipes)} cañerías').font = Font(bold=True)

    # ── Nodos ─────────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet('Nodos')
    _hdr(ws2,['Nodo','Elevación (m)','Grid X','Grid Y'],[32,16,10,10])
    for i,n in enumerate(nodes,2):
        _row(ws2,i,[n['name'],n['elevation_m'],n['grid_x'],n['grid_y']])
    ws2.cell(row=len(nodes)+2,column=1,
             value=f'Total: {len(nodes)} nodos').font = Font(bold=True)

    # ── Componentes Especiales ────────────────────────────────────────────────
    ws3 = wb.create_sheet('Componentes Especiales')
    _hdr(ws3,['Tipo','Nombre','Parámetro','Valor','Unidad'],[20,24,28,16,12])
    row = 2
    for tipo, items, field_fn in [
        ('Tank',              data.get('tanks',[]),
         lambda t: [('Elevación',t['elevation_m'],'m'),
                    ('Presión superficial',t['surface_pressure_kpa_abs'],t['pressure_unit']),
                    ('Nivel de líquido',t['liquid_level_m'],'m')]),
        ('Bomba Centrífuga',  data.get('pumps',[]),
         lambda p: [('Modo operación',p['operation_mode'],'—'),
                    ('Caudal',p['flow_rate'],p['flow_rate_unit']),
                    ('Elev. succión',p['suction_elevation_m'],'m'),
                    ('Elev. descarga',p['discharge_elevation_m'],'m')]),
        ('Fixed dP Device',   data.get('fixed_dp',[]),
         lambda d: [('Elev. entrada',d['inlet_elevation_m'],'m'),
                    ('Elev. salida',d['outlet_elevation_m'],'m'),
                    ('Caída de presión',d['pressure_drop'],d['pressure_drop_unit'])]),
        ('Válvula de Control',data.get('control_valves',[]),
         lambda v: [('Elevación',v['elevation_m'],'m'),
                    ('Modo operación',v['operation_mode'],'—'),
                    ('Coef. de flujo',v['flow_coefficient'],v['flow_coefficient_unit'])]),
        ('Pressure Boundary', data.get('pressure_boundaries',[]),
         lambda pb: [('Elevación',pb['elevation_m'],'m'),
                     ('Presión',pb['pressure_kpa_abs'],pb['pressure_unit']),
                     ('Modo',pb['operation_mode'],'—')]),
    ]:
        for item in items:
            for param,val,unit in field_fn(item):
                _row(ws3,row,[tipo,item['name'],param,val,unit]); row += 1
        if items: row += 1

    wb.save(path)
    print(f"Excel guardado: {path}")

# ─────────────────────────────────────────────────────────────────────────────
# CLI interactivo
# ─────────────────────────────────────────────────────────────────────────────

def seleccionar_archivo_pipe():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    archivos   = [f for f in os.listdir(script_dir) if f.endswith('.pipe')]
    if not archivos:
        print("No se encontraron archivos .pipe en el directorio del script.")
        sys.exit(1)
    print("\nArchivos disponibles:")
    for i,a in enumerate(archivos,1):
        print(f"  [{i}] {a}")
    while True:
        try:
            s = int(input("\nSelecciona un archivo: "))
            if 1 <= s <= len(archivos):
                return os.path.join(script_dir, archivos[s-1])
        except ValueError: pass
        print("Selección inválida.")

def seleccionar_formato_salida():
    print("\n¿Deseas guardar el resultado?")
    print("  [1] CSV  (tres archivos: cañerías, nodos, componentes)")
    print("  [2] Excel  (tres hojas en un .xlsx)")
    print("  [3] JSON")
    print("  [4] No guardar")
    while True:
        op = input("Selecciona: ").strip()
        if op in ('1','2','3','4'):
            return {'1':'csv','2':'excel','3':'json','4':None}[op]
        print("Opción inválida.")

def main():
    archivo = seleccionar_archivo_pipe()
    print(f"\nProcesando: {archivo}")
    data = extract_all(archivo)
    print_table(data)
    fmt = seleccionar_formato_salida()
    if fmt == 'csv':
        nombre = input("Nombre base (ej: resultado.csv): ").strip()
        save_csv(data, nombre)
    elif fmt == 'excel':
        nombre = input("Nombre del Excel (ej: resultado.xlsx): ").strip()
        if not nombre.endswith('.xlsx'): nombre += '.xlsx'
        save_excel(data, nombre)
    elif fmt == 'json':
        nombre = input("Nombre del JSON (ej: resultado.json): ").strip()
        if not nombre.endswith('.json'): nombre += '.json'
        export = {
            'pipes': [{k:v for k,v in p.items() if k!='fittings'}
                      |{'fittings_detail':p['fittings']} for p in data['pipes']],
            'nodes': data['nodes'],
            **{k:data[k] for k in ('tanks','pumps','fixed_dp','control_valves','pressure_boundaries')}
        }
        with open(nombre,'w',encoding='utf-8') as f:
            json.dump(export,f,ensure_ascii=False,indent=2)
        print(f"JSON guardado: {nombre}")
    return data

if __name__ == '__main__':
    main()