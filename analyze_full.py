# -*- coding: utf-8 -*-
"""
Script de procesamiento y análisis cuantitativo y cualitativo integral
para el Programa de Acción Climática (PAC) 2026 - Línea Verde
"""

import openpyxl
import json
import unicodedata
import re
from collections import defaultdict, Counter

def clean_str(s):
    if s is None:
        return ''
    return str(s).strip()

def strip_accents(s):
    if not s:
        return ''
    return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn').strip().lower()

# Diccionarios de nombres chilenos frecuentes para inferencia de género
FEMALE_NAMES = {
    'amanda', 'danna', 'isabella', 'paulina', 'sophia', 'xochtiel', 'abigail', 'ayleen', 
    'estefania', 'eysa', 'fabiola', 'gladys', 'mariana', 'matilde', 'matilda', 'monserrat',
    'pamela', 'valentina', 'dafne', 'lia', 'leah', 'julieta', 'kamila', 'janely', 'isidora',
    'salome', 'sara', 'daniela', 'fabrianny', 'fernanda', 'susy', 'monica', 'carol', 'pia',
    'guisselle', 'jocelyn', 'rosmery', 'javiera', 'mariel', 'kathleen', 'martina', 'ignacia',
    'constanza', 'catalina', 'camila', 'antonia', 'francisca', 'florencia', 'lucia', 'valeria',
    'sofia', 'emilia', 'trinidad', 'maite', 'mia', 'zoe', 'antonella', 'dominic', 'breymarw',
    'gioberlys', 'nahuel'
}

MALE_NAMES = {
    'agustin', 'axel', 'deiby', 'ignacio', 'raul', 'benjamin', 'demian',
    'matias', 'maximiliano', 'pablo', 'emilio', 'tomas', 'gian',
    'gabriel', 'carlos', 'mateo', 'lucas', 'joaquin', 'vicente', 'martin', 'santiago',
    'alonso', 'cristobal', 'bruno', 'dilan', 'ian', 'alvaro', 'rodrigo', 'pedro',
    'nicolas', 'sebastian', 'felipe', 'diego', 'claudio'
}

def infer_gender(full_name):
    first = strip_accents(full_name).split()[0] if full_name else ''
    if first in FEMALE_NAMES:
        return 'Femenino'
    if first in MALE_NAMES:
        return 'Masculino'
    if first.endswith('a'):
        return 'Femenino'
    if first.endswith('o') or first.endswith('el') or first.endswith('an') or first.endswith('or'):
        return 'Masculino'
    return 'Femenino' if first.endswith(('i', 'y')) else 'Masculino'

# Claves de respuestas correctas
KEYS = {
    'Hito 3': ['d', 'a', 'b', 'c', 'c'],
    'Hito 4': ['b', 'c', 'a', 'b', 'b']
}

# Clasificación cualitativa de respuestas abiertas (rúbrica de progresión)
KEYWORDS_SCIENTIFIC_H3 = ['ecosistema', 'biodiversidad', 'gaviotin', '50 gr', '50 gramos', 'microorganismo', 'flora', 'fauna', 'salina', 'desierto costero', 'cadena trofica', 'humedal', 'reserva', 'portada', 'endemica', 'migratoria']
KEYWORDS_BASIC_H3 = ['nada', 'no se', 'bonito', 'lindo', 'agua', 'poza', 'hoyo', 'paloma', 'gaviota']

KEYWORDS_SCIENTIFIC_H4 = ['atmosfera', 'escudo', 'capa de ozono', 'gases', 'efecto invernadero', 'radiacion', 'uv', 'contaminacion luminica', 'astronomia', 'telescopio', 'paranal', 'luz azul', 'cielos limpios', 'patrimonio']
KEYWORDS_BASIC_H4 = ['nada', 'no se', 'bonito', 'estrellas', 'noche', 'luz', 'calor']

def evaluate_open_text(text, hito):
    if not text or len(text.strip()) < 3:
        return {'level': 0, 'label': 'Sin respuesta', 'text': ''}
    t_clean = strip_accents(text)
    if any(b in t_clean for b in ['no se', 'nada', 'ninguna', 'no me acuerdo', 'nose']):
        return {'level': 1, 'label': 'Desconocimiento / Básico', 'text': text}
    
    keywords = KEYWORDS_SCIENTIFIC_H3 if hito == 'Hito 3' else KEYWORDS_SCIENTIFIC_H4
    matches = sum(1 for kw in keywords if kw in t_clean)
    
    if matches >= 2 or len(t_clean.split()) >= 10:
        return {'level': 3, 'label': 'Apropiación Científica / Detallado', 'text': text}
    elif matches == 1 or len(t_clean.split()) >= 4:
        return {'level': 2, 'label': 'Comprensión Inicial / Intuitivo', 'text': text}
    else:
        return {'level': 1, 'label': 'Desconocimiento / Básico', 'text': text}

def main():
    excel_path = r'E:\Habilidades Verdes\datos_evaluacion\Participantes PAC 2026.xlsx'
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    # 1. Metadatos de Establecimientos y Líderes
    schools = {}
    s_part = wb['Participantes 2026']
    for r in range(2, s_part.max_row + 1):
        sch = clean_str(s_part.cell(r, 2).value)
        if sch and sch not in schools:
            schools[sch] = {
                'nombre': sch,
                'equipo': s_part.cell(r, 1).value,
                'direccion': clean_str(s_part.cell(r, 3).value),
                'dependencia': clean_str(s_part.cell(r, 4).value) or 'Municipal',
                'certificacion': clean_str(s_part.cell(r, 5).value) or 'No',
                'participacion_previa': clean_str(s_part.cell(r, 6).value) or 'No',
                'lider1': clean_str(s_part.cell(r, 8).value),
                'email1': clean_str(s_part.cell(r, 9).value),
                'tel1': clean_str(s_part.cell(r, 10).value),
                'lider2': clean_str(s_part.cell(r, 11).value),
                'email2': clean_str(s_part.cell(r, 12).value),
                'tel2': clean_str(s_part.cell(r, 13).value),
            }
            
    # Resumen Líderes
    s_lid = wb['Resumen Líderes']
    for r in range(2, s_lid.max_row + 1):
        sch = clean_str(s_lid.cell(r, 2).value)
        if sch in schools:
            if not schools[sch]['lider1']:
                schools[sch]['lider1'] = clean_str(s_lid.cell(r, 3).value)
                schools[sch]['tel1'] = clean_str(s_lid.cell(r, 4).value)
                schools[sch]['email1'] = clean_str(s_lid.cell(r, 5).value)
            if not schools[sch]['lider2']:
                schools[sch]['lider2'] = clean_str(s_lid.cell(r, 6).value)
                schools[sch]['tel2'] = clean_str(s_lid.cell(r, 7).value)
                schools[sch]['email2'] = clean_str(s_lid.cell(r, 8).value)

    # 2. Procesamiento de Hitos
    hitos_data = {}
    all_participants = {}

    for hito_name, sheet_name in [('Hito 3', 'Encuestas Hito 3_Verde'), ('Hito 4', 'Encuestas Hito 4_Verde')]:
        ws = wb[sheet_name]
        key = KEYS[hito_name]
        q_titles = [clean_str(ws.cell(1, c).value) for c in range(5, 10)]
        open_q_titles = [clean_str(ws.cell(1, c).value) for c in range(10, 14)]
        desafio_name = "Los otros vecinos de Antofagasta" if hito_name == 'Hito 3' else "Proteger nuestro escudo invisible"
        
        # Agrupar registros por participante
        p_dict = defaultdict(dict)
        for r in range(2, ws.max_row + 1):
            name = clean_str(ws.cell(r, 1).value)
            if not name:
                continue
            role = clean_str(ws.cell(r, 2).value) or 'Estudiante'
            school = clean_str(ws.cell(r, 3).value)
            pp = clean_str(ws.cell(r, 4).value).upper()
            if pp not in ['PRE', 'POST']:
                continue
            
            mc_ans = [clean_str(ws.cell(r, c).value).lower() for c in range(5, 10)]
            open_ans = [clean_str(ws.cell(r, c).value) for c in range(10, 14)]
            
            correct = [1 if mc_ans[i] == key[i] else 0 for i in range(5)]
            score = sum(correct)
            
            p_dict[name]['name'] = name
            p_dict[name]['role'] = role
            if school:
                p_dict[name]['school'] = school
            p_dict[name][pp] = {
                'row': r,
                'mc': mc_ans,
                'correct': correct,
                'score': score,
                'open': open_ans
            }
            
        # Evaluar participantes con PRE y POST
        paired_list = []
        for name, d in p_dict.items():
            gender = infer_gender(name)
            sch = d.get('school', 'Sin Asignar')
            # Normalizar nombre escuela si es Liceo Sabella
            if 'Sabella' in sch:
                sch = 'Liceo Andrés Sabella'
            
            # Registrar participante general
            if name not in all_participants:
                all_participants[name] = {
                    'name': name,
                    'role': d.get('role', 'Estudiante'),
                    'school': sch,
                    'gender': gender,
                    'hitos': {}
                }
            
            has_pre = 'PRE' in d
            has_post = 'POST' in d
            
            if has_pre and has_post:
                pre_score = d['PRE']['score']
                post_score = d['POST']['score']
                delta = post_score - pre_score
                status = 'Mejoró' if delta > 0 else ('Mantuvo' if delta == 0 else 'Retrocedió')
                
                # Progresión por pregunta cerrada
                q_evo = []
                for i in range(5):
                    cp = d['PRE']['correct'][i]
                    cpo = d['POST']['correct'][i]
                    if cp == 0 and cpo == 1:
                        q_evo.append('Mejoró')
                    elif cp == 1 and cpo == 1:
                        q_evo.append('Mantuvo Correcto')
                    elif cp == 0 and cpo == 0:
                        q_evo.append('Mantuvo Incorrecto')
                    else:
                        q_evo.append('Retrocedió')
                        
                # Análisis de preguntas abiertas
                open_analysis = []
                for i in range(4):
                    t_pre = d['PRE']['open'][i]
                    t_post = d['POST']['open'][i]
                    ev_pre = evaluate_open_text(t_pre, hito_name)
                    ev_post = evaluate_open_text(t_post, hito_name)
                    open_analysis.append({
                        'question': open_q_titles[i],
                        'pre_text': t_pre,
                        'post_text': t_post,
                        'pre_eval': ev_pre,
                        'post_eval': ev_post,
                        'level_delta': ev_post['level'] - ev_pre['level']
                    })
                
                p_obj = {
                    'name': name,
                    'role': d.get('role', 'Estudiante'),
                    'school': sch,
                    'gender': gender,
                    'pre_score': pre_score,
                    'post_score': post_score,
                    'delta': delta,
                    'status': status,
                    'q_evo': q_evo,
                    'pre_mc': d['PRE']['mc'],
                    'post_mc': d['POST']['mc'],
                    'open_analysis': open_analysis
                }
                paired_list.append(p_obj)
                all_participants[name]['hitos'][hito_name] = p_obj
                
        # Métricas del Hito
        N = len(paired_list)
        status_c = Counter(p['status'] for p in paired_list)
        avg_pre = sum(p['pre_score'] for p in paired_list) / N if N else 0
        avg_post = sum(p['post_score'] for p in paired_list) / N if N else 0
        avg_delta = sum(p['delta'] for p in paired_list) / N if N else 0
        
        # Ganancia de Hake g = (post - pre) / (5 - pre) para pre < 5
        g_list = [(p['post_score'] - p['pre_score']) / (5.0 - p['pre_score']) for p in paired_list if p['pre_score'] < 5]
        hake_g = sum(g_list) / len(g_list) if g_list else 0
        
        # Preguntas cerradas métricas
        questions_metrics = []
        for i in range(5):
            pre_c = sum(p['pre_mc'][i] == key[i] for p in paired_list)
            post_c = sum(p['post_mc'][i] == key[i] for p in paired_list)
            trans_c = Counter(p['q_evo'][i] for p in paired_list)
            
            # Detalle de errores en el post
            post_distractors = Counter(p['post_mc'][i].upper() for p in paired_list if p['post_mc'][i] != key[i])
            
            questions_metrics.append({
                'num': i + 1,
                'title': q_titles[i],
                'correct_key': key[i].upper(),
                'pre_correct': pre_c,
                'pre_pct': round(pre_c / N * 100, 1) if N else 0,
                'post_correct': post_c,
                'post_pct': round(post_c / N * 100, 1) if N else 0,
                'diff': post_c - pre_c,
                'transitions': dict(trans_c),
                'post_distractors': dict(post_distractors)
            })
            
        # Agrupación por Escuela
        by_school = defaultdict(list)
        for p in paired_list:
            by_school[p['school']].append(p)
            
        school_rankings = []
        for sname, plist in by_school.items():
            s_n = len(plist)
            s_pre = sum(p['pre_score'] for p in plist) / s_n
            s_post = sum(p['post_score'] for p in plist) / s_n
            s_del = sum(p['delta'] for p in plist) / s_n
            s_st = Counter(p['status'] for p in plist)
            meta = schools.get(sname, {})
            school_rankings.append({
                'school': sname,
                'n': s_n,
                'pre_avg': round(s_pre, 2),
                'post_avg': round(s_post, 2),
                'delta_avg': round(s_del, 2),
                'status_counts': dict(s_st),
                'pct_mejora': round(s_st['Mejoró'] / s_n * 100, 1),
                'lider1': meta.get('lider1', ''),
                'email1': meta.get('email1', ''),
                'tel1': meta.get('tel1', ''),
                'dependencia': meta.get('dependencia', 'Municipal'),
                'certificacion': meta.get('certificacion', 'No')
            })
        school_rankings.sort(key=lambda x: -x['delta_avg'])
        
        # Agrupación por Género
        by_gender = defaultdict(list)
        for p in paired_list:
            by_gender[p['gender']].append(p)
            
        gender_metrics = {}
        for gname, plist in by_gender.items():
            g_n = len(plist)
            g_pre = sum(p['pre_score'] for p in plist) / g_n
            g_post = sum(p['post_score'] for p in plist) / g_n
            g_del = sum(p['delta'] for p in plist) / g_n
            g_st = Counter(p['status'] for p in plist)
            gender_metrics[gname] = {
                'n': g_n,
                'pre_avg': round(g_pre, 2),
                'post_avg': round(g_post, 2),
                'delta_avg': round(g_del, 2),
                'status_counts': dict(g_st),
                'pct_mejora': round(g_st['Mejoró'] / g_n * 100, 1)
            }
            
        # Agrupación por Rol (Estudiante vs Líder)
        by_role = defaultdict(list)
        for p in paired_list:
            by_role[p['role']].append(p)
        role_metrics = {}
        for rname, plist in by_role.items():
            r_n = len(plist)
            r_pre = sum(p['pre_score'] for p in plist) / r_n
            r_post = sum(p['post_score'] for p in plist) / r_n
            r_del = sum(p['delta'] for p in plist) / r_n
            role_metrics[rname] = {
                'n': r_n,
                'pre_avg': round(r_pre, 2),
                'post_avg': round(r_post, 2),
                'delta_avg': round(r_del, 2)
            }

        # Casos con retroceso o sin mejora
        critical_cases = [p for p in paired_list if p['status'] in ['Retrocedió', 'Mantuvo']]
        
        hitos_data[hito_name] = {
            'desafio': desafio_name,
            'n_participants': N,
            'avg_pre': round(avg_pre, 2),
            'avg_post': round(avg_post, 2),
            'avg_delta': round(avg_delta, 2),
            'hake_gain': round(hake_g, 3),
            'status_counts': dict(status_c),
            'questions': questions_metrics,
            'open_questions': open_q_titles,
            'schools': school_rankings,
            'gender': gender_metrics,
            'role': role_metrics,
            'critical_cases_count': len(critical_cases),
            'participants': paired_list
        }

    # 3. Exportar JSON completo para el Dashboard
    output_data = {
        'metadata': {
            'programa': 'Programa de Acción Climática (PAC) 2026',
            'linea': 'Línea Verde',
            'fecha_analisis': '2026-09-30',
            'total_participantes_registrados': len(all_participants)
        },
        'schools_directory': schools,
        'hitos': hitos_data,
        'all_participants': list(all_participants.values())
    }
    
    output_json_path = r'E:\Habilidades Verdes\dashboard_evaluacion\data.json'
    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
        
    print(f'Procesamiento completado con éxito. Archivo generado en: {output_json_path}')

if __name__ == '__main__':
    main()
