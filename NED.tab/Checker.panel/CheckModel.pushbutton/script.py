# -*- coding: utf-8 -*-
__title__ = 'Opening\nChecker'
__doc__ = 'Check presence and size of openings for MEP elements in structural models'
__author__ = 'NED DC'

import clr
import os
import datetime
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, CheckBox, TextBox, Button,
    FlowLayoutPanel, GroupBox, Panel, RadioButton,
    ComboBox, ComboBoxStyle,
    DataGridView, DataGridViewTextBoxColumn,
    DataGridViewSelectionMode, DataGridViewAutoSizeColumnsMode,
    DataGridViewColumnHeadersHeightSizeMode, DataGridViewColumnSortMode,
    FormBorderStyle, DialogResult, DockStyle, BorderStyle,
    FlowDirection, AnchorStyles,
    FolderBrowserDialog, MessageBox, MessageBoxButtons,
    MessageBoxIcon, FormStartPosition,
    ScrollBars
)
from System.Drawing import Point, Size, Font, FontStyle, Color

from pyrevit import revit, DB, script, forms

doc = revit.doc

# =============================================
# CONSTANTS
# =============================================
FT_TO_MM = 304.8

# Ключевые слова для определения бетона: иврит, английский, русский
CONCRETE_KW = [
    u'בטון',   # בטון (иврит)
    u'Concrete', u'concrete', u'CONCRETE',
    u'бетон', u'Бетон',
]

# Коды дисциплин на позиции [2] в имени файла: S-HA-[КОД]-[КОМПАНИЯ]-[ЛОКАЦИЯ]-RVT2X
STRUCTURAL_FILE_CODES = ['AR', 'S', 'ST', 'STR', 'O', 'OP']
MEP_FILE_CODES        = ['H', 'P', 'E', 'F', 'T', 'HV', 'PL', 'EL', 'FL']
SKIP_FILE_CODES       = ['TR', 'SI', 'CO', 'CR', 'G', 'Z', 'B', 'FU', 'ID']

STATUS_NO_OPENING = 'No Opening'
STATUS_OK         = 'Opening OK'
STATUS_UNDERSIZED = 'Undersized'
STATUS_EMPTY      = 'Empty'


# =============================================
# LINK CLASSIFICATION
# =============================================
def classify_link(link_name):
    """Определяет тип модели по коду дисциплины в имени файла.

    Основной формат: S-HA-[КОД]-... → код на позиции [2].
    Альтернативный: SHA-OP-... → код на позиции [1] (сводная модель отверстий).
    """
    parts = link_name.replace('.rvt', '').replace('.RVT', '').split('-')

    SKIP_UP       = [c.upper() for c in SKIP_FILE_CODES]
    STRUCTURAL_UP = [c.upper() for c in STRUCTURAL_FILE_CODES]
    MEP_UP        = [c.upper() for c in MEP_FILE_CODES]

    # Проверяем позиции [2] и [1] (для моделей с укороченным префиксом типа SHA-OP-...)
    for idx in [2, 1]:
        if len(parts) > idx:
            d = parts[idx].upper()
            if d in SKIP_UP:       return 'skip'
            if d in STRUCTURAL_UP: return 'structural'
            if d in MEP_UP:        return 'mep'

    # Запасной путь: сканируем все части
    for part in parts:
        if part.upper() in SKIP_UP:       return 'skip'
    for part in parts:
        if part.upper() in MEP_UP:        return 'mep'
    for part in parts:
        if part.upper() in STRUCTURAL_UP: return 'structural'
    return 'unknown'


def get_all_revit_links():
    """Получает все подключённые Revit Links из текущего документа."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()
    links = []
    for inst in collector:
        ltype = doc.GetElement(inst.GetTypeId())
        if ltype is None:
            continue
        name = ltype.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        if not name:
            name = inst.Name
        links.append({'name': name, 'instance': inst, 'category': classify_link(name)})
    return links


# =============================================
# UTILITY — GEOMETRY
# =============================================
def ft_to_mm(ft):
    return ft * FT_TO_MM


def transform_bbox(bbox, transform):
    """Трансформирует BoundingBox из координат линка в координаты хоста."""
    corners = [
        DB.XYZ(x, y, z)
        for x in [bbox.Min.X, bbox.Max.X]
        for y in [bbox.Min.Y, bbox.Max.Y]
        for z in [bbox.Min.Z, bbox.Max.Z]
    ]
    pts = [transform.OfPoint(c) for c in corners]
    result = DB.BoundingBoxXYZ()
    result.Min = DB.XYZ(min(p.X for p in pts), min(p.Y for p in pts), min(p.Z for p in pts))
    result.Max = DB.XYZ(max(p.X for p in pts), max(p.Y for p in pts), max(p.Z for p in pts))
    return result


def bboxes_intersect(b1, b2):
    return (b1.Min.X <= b2.Max.X and b1.Max.X >= b2.Min.X and
            b1.Min.Y <= b2.Max.Y and b1.Max.Y >= b2.Min.Y and
            b1.Min.Z <= b2.Max.Z and b1.Max.Z >= b2.Min.Z)


# =============================================
# UTILITY — ELEMENT PROPERTIES
# =============================================
def get_level_name(elem, link_doc):
    try:
        lid = elem.LevelId
        if lid and lid != DB.ElementId.InvalidElementId:
            lv = link_doc.GetElement(lid)
            if lv:
                return lv.Name
    except Exception:
        pass
    return 'Unknown'


def _get_elem_type(elem, link_doc):
    """Возвращает объект типа элемента несколькими способами."""
    # Wall — прямое свойство, не требует поиска в документе
    wt = getattr(elem, 'WallType', None)
    if wt is not None:
        return wt
    # Floor и прочие — через GetTypeId()
    try:
        tid = elem.GetTypeId()
        if tid and tid.IntegerValue != -1:
            t = link_doc.GetElement(tid)
            if t is not None:
                return t
    except Exception:
        pass
    return None


def get_type_name(elem, link_doc):
    try:
        t = _get_elem_type(elem, link_doc)
        if t is not None:
            return t.Name or 'Unknown'
    except Exception:
        pass
    return 'Unknown'


def is_concrete(elem, link_doc):
    """Определяет бетонный ли элемент по ключевым словам в имени типа.

    Проверка по слоям материалов намеренно отключена — GetCompoundStructure()
    на элементах из Revit Link нестабилен и вызывает краш Revit.
    """
    try:
        t = _get_elem_type(elem, link_doc)
        if t is None:
            return False
        return any(kw in (t.Name or '') for kw in CONCRETE_KW)
    except Exception:
        return False


def get_thickness_mm(elem, elem_type, link_doc):
    """Возвращает толщину стены или перекрытия в мм."""
    try:
        if elem_type == 'Wall':
            return ft_to_mm(elem.Width)
        else:  # Floor — используем параметр, не GetCompoundStructure (нестабилен в Links)
            p = elem.get_Parameter(DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
            if p and p.AsDouble() > 0:
                return ft_to_mm(p.AsDouble())
    except Exception:
        pass
    return 0.0


def get_mep_size_mm(elem):
    """Возвращает (ширина_мм, высота_мм) MEP элемента."""
    try:
        # Труба
        p = elem.get_Parameter(DB.BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
        if p and p.AsDouble() > 0:
            d = ft_to_mm(p.AsDouble())
            return d, d
        # Кондуит
        p = elem.get_Parameter(DB.BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
        if p and p.AsDouble() > 0:
            d = ft_to_mm(p.AsDouble())
            return d, d
        # Прямоугольный воздуховод / кабельный лоток
        pw = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        ph = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        if pw and ph and pw.AsDouble() > 0:
            return ft_to_mm(pw.AsDouble()), ft_to_mm(ph.AsDouble())
        # Круглый воздуховод
        p = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
        if p and p.AsDouble() > 0:
            d = ft_to_mm(p.AsDouble())
            return d, d
    except Exception:
        pass
    return 0.0, 0.0


def get_opening_dims_mm(o_bbox):
    """Возвращает (ширина_мм, высота_мм) отверстия из его BoundingBox.

    Ширина — наибольшее из горизонтальных измерений,
    высота — вертикальное (ось Z).
    """
    dx = ft_to_mm(o_bbox.Max.X - o_bbox.Min.X)
    dy = ft_to_mm(o_bbox.Max.Y - o_bbox.Min.Y)
    dz = ft_to_mm(o_bbox.Max.Z - o_bbox.Min.Z)
    width  = max(dx, dy)   # горизонтальная сторона (зависит от ориентации стены)
    height = dz             # вертикаль всегда по Z
    return width, height


def get_elev_from_level_mm(mep_bbox, struct_elem, struct_doc):
    """Высота низа MEP элемента от уровня конструктивного элемента (мм)."""
    try:
        lid = struct_elem.LevelId
        if lid and lid != DB.ElementId.InvalidElementId:
            lv = struct_doc.GetElement(lid)
            if lv:
                return ft_to_mm(mep_bbox.Min.Z - lv.Elevation)
    except Exception:
        pass
    return 0.0


def get_cat_name(elem):
    try:
        return elem.Category.Name
    except Exception:
        return 'MEP'


# =============================================
# DATA COLLECTION FROM LINKED DOCUMENTS
# =============================================
def diagnose_wall_types(selected_structural, output):
    """Выводит уникальные имена типов стен из конструктивных линков (без слоёв)."""
    output.print_md('## Wall type diagnostics')
    for s_link in selected_structural:
        inst = s_link['instance']
        link_doc = inst.GetLinkDocument()
        if link_doc is None:
            continue
        output.print_md('### {}'.format(s_link['name']))
        type_names = set()
        walls = DB.FilteredElementCollector(link_doc)\
            .OfClass(DB.Wall)\
            .WhereElementIsNotElementType()\
            .ToElements()
        for wall in walls:
            try:
                wt = getattr(wall, 'WallType', None)
                if wt is not None:
                    type_names.add(wt.Name or '(empty)')
            except Exception:
                continue
        output.print_md('**Wall types ({} unique):**'.format(len(type_names)))
        for n in sorted(type_names):
            output.print_md('- `{}`'.format(n))
    output.print_md('---')


def get_mep_elements(link_doc):
    """Получает все MEP кривые (трубы, воздуховоды, лотки, кондуиты) из линка."""
    try:
        return list(
            DB.FilteredElementCollector(link_doc)
            .OfClass(DB.MEPCurve)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return []


def get_struct_elements(link_doc):
    """Получает стены и перекрытия из линка."""
    results = []
    try:
        for w in DB.FilteredElementCollector(link_doc).OfClass(DB.Wall).WhereElementIsNotElementType().ToElements():
            results.append((w, 'Wall'))
    except Exception:
        pass
    try:
        for f in DB.FilteredElementCollector(link_doc).OfClass(DB.Floor).WhereElementIsNotElementType().ToElements():
            results.append((f, 'Floor'))
    except Exception:
        pass
    return results


def get_openings(link_doc):
    """Получает элементы-отверстия из линка.

    Собирает три типа:
    - DB.Opening — стандартные вырезы Revit (Wall Opening, Floor Opening)
    - DB.DirectShape — IFC-импортированные элементы
    - Generic Model family instances — семейства-маркеры отверстий
    """
    results = []
    # Стандартные вырезы Revit
    try:
        for el in DB.FilteredElementCollector(link_doc)\
                .OfClass(DB.Opening)\
                .WhereElementIsNotElementType()\
                .ToElements():
            results.append(el)
    except Exception:
        pass
    # IFC / облегчённая геометрия
    try:
        for el in DB.FilteredElementCollector(link_doc)\
                .OfClass(DB.DirectShape)\
                .WhereElementIsNotElementType()\
                .ToElements():
            results.append(el)
    except Exception:
        pass
    # Семейства категории Generic Model
    try:
        for el in DB.FilteredElementCollector(link_doc)\
                .OfCategory(DB.BuiltInCategory.OST_GenericModel)\
                .WhereElementIsNotElementType()\
                .ToElements():
            results.append(el)
    except Exception:
        pass
    return results


# =============================================
# STEP 2 — INTERSECTION CHECK
# =============================================
def run_check(selected_structural, selected_mep, gap_mm, output):
    """Находит пересечения MEP элементов с конструктивом и определяет статус отверстий."""
    results = []

    # --- Загрузка конструктивных элементов ---
    output.print_md('**Loading structural elements...**')
    struct_index  = []   # список dict с данными каждого конструктивного элемента
    openings_host = []   # [(opening, bbox_in_host, struct_doc)] для всех линков

    for s_link in selected_structural:
        inst = s_link['instance']
        link_doc = inst.GetLinkDocument()
        if link_doc is None:
            output.print_md('- {} — *not loaded, skipped*'.format(s_link['name']))
            continue

        transform = inst.GetTotalTransform()
        struct_elems = get_struct_elements(link_doc)
        openings = get_openings(link_doc)

        count = 0
        for elem, etype in struct_elems:
            try:
                raw_bb = elem.get_BoundingBox(None)
                if raw_bb is None:
                    continue
                struct_index.append({
                    'elem':         elem,
                    'etype':        etype,
                    'bb':           transform_bbox(raw_bb, transform),
                    'link_doc':     link_doc,
                    'transform':    transform,
                    'link_name':    s_link['name'],
                    'thickness_mm': get_thickness_mm(elem, etype, link_doc),
                    'is_concrete':  is_concrete(elem, link_doc),
                    'level':        get_level_name(elem, link_doc),
                    'type_name':    get_type_name(elem, link_doc),
                    'openings':     openings,
                })
                count += 1
            except Exception:
                continue

        for op in openings:
            try:
                ob = op.get_BoundingBox(None)
                if ob:
                    openings_host.append((op, transform_bbox(ob, transform), link_doc))
            except Exception:
                continue

        output.print_md('- {} — {} walls/floors'.format(s_link['name'], count))

    # Если конструктивных элементов не найдено — режим OP-модели
    if not struct_index:
        output.print_md('*No walls/floors found — switching to **Opening Model Mode***')
        return run_check_opening_model(selected_structural, selected_mep, gap_mm, output)

    # Быстрый поиск конструктивного элемента по ID
    struct_by_id = {se['elem'].Id.IntegerValue: se for se in struct_index}

    # --- Проверка MEP элементов ---
    output.print_md('**Checking MEP elements...**')
    used_opening_ids = set()   # ID отверстий через которые прошёл MEP

    for m_link in selected_mep:
        inst = m_link['instance']
        link_doc = inst.GetLinkDocument()
        if link_doc is None:
            output.print_md('- {} — *not loaded, skipped*'.format(m_link['name']))
            continue

        transform = inst.GetTotalTransform()
        mep_elems = get_mep_elements(link_doc)
        output.print_md('- {} — {} MEP curves'.format(m_link['name'], len(mep_elems)))

        for mep in mep_elems:
            try:
                raw_bb = mep.get_BoundingBox(None)
                if raw_bb is None:
                    continue
                mep_bb   = transform_bbox(raw_bb, transform)
                mep_w, mep_h = get_mep_size_mm(mep)
                cat_name = get_cat_name(mep)
            except Exception:
                continue

            for se in struct_index:
                if not bboxes_intersect(mep_bb, se['bb']):
                    continue

                # Пересечение найдено — ищем отверстие
                found_op    = None
                found_ob    = None

                for op, ob_host, op_doc in openings_host:
                    if not bboxes_intersect(mep_bb, ob_host):
                        continue
                    # Отверстие должно пересекаться и с конструктивным элементом
                    if not bboxes_intersect(ob_host, se['bb']):
                        continue
                    # Если отверстие из того же линка что и конструктив —
                    # проверяем что оно принадлежит именно этой стене/перекрытию.
                    # Если из отдельной модели отверстий — разрешаем без host-проверки.
                    try:
                        if op_doc is se['link_doc']:
                            host = op.Host
                            if host and host.Id != se['elem'].Id:
                                continue
                    except Exception:
                        pass
                    found_op = op
                    found_ob = ob_host
                    used_opening_ids.add(op.Id.IntegerValue)
                    break

                if found_op is None:
                    status       = STATUS_NO_OPENING
                    opening_size = '-'
                else:
                    o_w, o_h = get_opening_dims_mm(found_ob)
                    opening_size = '{}x{} mm'.format(int(round(o_w)), int(round(o_h)))
                    if o_w >= mep_w + gap_mm * 2 and o_h >= mep_h + gap_mm * 2:
                        status = STATUS_OK
                    else:
                        status = STATUS_UNDERSIZED

                results.append({
                    'status':       status,
                    'level':        se['level'],
                    'mep_system':   m_link['name'],
                    'mep_type':     cat_name,
                    'mep_id':       mep.Id.IntegerValue,
                    'mep_w_mm':     mep_w,
                    'mep_h_mm':     mep_h,
                    'struct_type':  se['etype'],
                    'is_concrete':  se['is_concrete'],
                    'type_name':    se['type_name'],
                    'thickness_mm': se['thickness_mm'],
                    'opening_size': opening_size,
                    'elevation_mm': int(round(get_elev_from_level_mm(mep_bb, se['elem'], se['link_doc']))),
                    'struct_id':    se['elem'].Id.IntegerValue,
                })

    # --- Пустые отверстия: открытия без MEP ---
    for op, ob_host, op_doc in openings_host:
        if op.Id.IntegerValue in used_opening_ids:
            continue
        try:
            host = op.Host
            if host is None:
                continue
            se = struct_by_id.get(host.Id.IntegerValue)
            if se is None:
                continue
            o_w, o_h = get_opening_dims_mm(ob_host)
            results.append({
                'status':       STATUS_EMPTY,
                'level':        se['level'],
                'mep_system':   '-',
                'mep_type':     '-',
                'mep_id':       0,
                'mep_w_mm':     0,
                'mep_h_mm':     0,
                'struct_type':  se['etype'],
                'is_concrete':  se['is_concrete'],
                'type_name':    se['type_name'],
                'thickness_mm': se['thickness_mm'],
                'opening_size': '{}x{} mm'.format(int(round(o_w)), int(round(o_h))),
                'elevation_mm': 0,
                'struct_id':    host.Id.IntegerValue,
            })
        except Exception:
            continue

    return results


def run_check_opening_model(selected_structural, selected_mep, gap_mm, output):
    """Режим сводной модели отверстий (OP-модель без стен).

    Каждый элемент OP-модели — маркер отверстия (DirectShape, Generic Model и т.п.).
    Проверяет, какие MEP элементы проходят через каждое отверстие.
    """
    results = []
    output.print_md('**Mode: Opening Model** — OP model elements treated as opening markers')

    # Собираем элементы-маркеры отверстий из всех выбранных structural-линков
    op_openings = []   # (elem, bbox_host, link_doc, link_name)

    for s_link in selected_structural:
        inst = s_link['instance']
        link_doc = inst.GetLinkDocument()
        if link_doc is None:
            output.print_md('- {} — *not loaded, skipped*'.format(s_link['name']))
            continue

        transform = inst.GetTotalTransform()

        # get_openings собирает: DB.Opening + DB.DirectShape + OST_GenericModel
        elems = get_openings(link_doc)

        # Дополнительно: FamilyInstance любой категории (кастомные семейства OP)
        try:
            fi_ids = set(e.Id.IntegerValue for e in elems)
            for el in DB.FilteredElementCollector(link_doc)\
                    .OfClass(DB.FamilyInstance)\
                    .WhereElementIsNotElementType()\
                    .ToElements():
                if el.Id.IntegerValue not in fi_ids:
                    elems.append(el)
                    fi_ids.add(el.Id.IntegerValue)
        except Exception:
            pass

        count = 0
        seen_ids = set()
        for el in elems:
            if el.Id.IntegerValue in seen_ids:
                continue
            seen_ids.add(el.Id.IntegerValue)
            try:
                bb = el.get_BoundingBox(None)
                if bb is None:
                    continue
                bb_host = transform_bbox(bb, transform)
                # Отфильтровываем микро-элементы (меньше 50 мм в любом измерении)
                dx = ft_to_mm(bb_host.Max.X - bb_host.Min.X)
                dy = ft_to_mm(bb_host.Max.Y - bb_host.Min.Y)
                dz = ft_to_mm(bb_host.Max.Z - bb_host.Min.Z)
                if max(dx, dy, dz) < 50:
                    continue
                op_openings.append((el, bb_host, link_doc, s_link['name']))
                count += 1
            except Exception:
                continue

        output.print_md('- {} — {} opening markers found'.format(s_link['name'], count))

    if not op_openings:
        output.print_md('*No opening elements found. Check model contents.*')
        return results

    # Проверяем каждый MEP элемент на пересечение с маркерами отверстий
    output.print_md('**Checking MEP elements against opening markers...**')
    used_op_ids = set()   # ID отверстий, через которые прошёл хотя бы один MEP

    for m_link in selected_mep:
        inst = m_link['instance']
        link_doc = inst.GetLinkDocument()
        if link_doc is None:
            output.print_md('- {} — *not loaded, skipped*'.format(m_link['name']))
            continue

        transform = inst.GetTotalTransform()
        mep_elems = get_mep_elements(link_doc)
        output.print_md('- {} — {} MEP curves'.format(m_link['name'], len(mep_elems)))

        for mep in mep_elems:
            try:
                raw_bb = mep.get_BoundingBox(None)
                if raw_bb is None:
                    continue
                mep_bb   = transform_bbox(raw_bb, transform)
                mep_w, mep_h = get_mep_size_mm(mep)
                cat_name = get_cat_name(mep)
            except Exception:
                continue

            for op_elem, op_bb, op_doc, op_link_name in op_openings:
                if not bboxes_intersect(mep_bb, op_bb):
                    continue

                used_op_ids.add(op_elem.Id.IntegerValue)
                o_w, o_h = get_opening_dims_mm(op_bb)
                opening_size = '{}x{} mm'.format(int(round(o_w)), int(round(o_h)))

                if mep_w > 0 and o_w >= mep_w + gap_mm * 2 and o_h >= mep_h + gap_mm * 2:
                    status = STATUS_OK
                elif mep_w > 0:
                    status = STATUS_UNDERSIZED
                else:
                    status = STATUS_OK   # размер MEP неизвестен — считаем OK

                try:
                    type_name = get_type_name(op_elem, op_doc)
                    level     = get_level_name(op_elem, op_doc)
                except Exception:
                    type_name = 'Unknown'
                    level     = 'Unknown'

                results.append({
                    'status':       status,
                    'level':        level,
                    'mep_system':   m_link['name'],
                    'mep_type':     cat_name,
                    'mep_id':       mep.Id.IntegerValue,
                    'mep_w_mm':     mep_w,
                    'mep_h_mm':     mep_h,
                    'struct_type':  'Opening',
                    'is_concrete':  False,
                    'type_name':    type_name,
                    'thickness_mm': 0,
                    'opening_size': opening_size,
                    'elevation_mm': int(round(ft_to_mm(mep_bb.Min.Z))),
                    'struct_id':    op_elem.Id.IntegerValue,
                })

    # Пустые отверстия — маркеры без MEP
    for op_elem, op_bb, op_doc, op_link_name in op_openings:
        if op_elem.Id.IntegerValue in used_op_ids:
            continue
        o_w, o_h = get_opening_dims_mm(op_bb)
        try:
            type_name = get_type_name(op_elem, op_doc)
            level     = get_level_name(op_elem, op_doc)
        except Exception:
            type_name = 'Unknown'
            level     = 'Unknown'

        results.append({
            'status':       STATUS_EMPTY,
            'level':        level,
            'mep_system':   '-',
            'mep_type':     '-',
            'mep_id':       0,
            'mep_w_mm':     0,
            'mep_h_mm':     0,
            'struct_type':  'Opening',
            'is_concrete':  False,
            'type_name':    type_name,
            'thickness_mm': 0,
            'opening_size': '{}x{} mm'.format(int(round(o_w)), int(round(o_h))),
            'elevation_mm': 0,
            'struct_id':    op_elem.Id.IntegerValue,
        })

    return results


def print_results(results, output, gap_mm):
    """Выводит итоги проверки в окно Output."""
    counts = {STATUS_NO_OPENING: 0, STATUS_OK: 0, STATUS_UNDERSIZED: 0, STATUS_EMPTY: 0}
    for r in results:
        if r['status'] in counts:
            counts[r['status']] += 1

    output.print_md('## Summary')
    output.print_md('Total intersections found: **{}**'.format(len(results)))
    output.print_md('- No Opening: **{}**'.format(counts[STATUS_NO_OPENING]))
    output.print_md('- Opening OK: **{}**'.format(counts[STATUS_OK]))
    output.print_md('- Undersized: **{}**'.format(counts[STATUS_UNDERSIZED]))
    output.print_md('- Empty (no MEP): **{}**'.format(counts[STATUS_EMPTY]))

    critical = [r for r in results
                if r['is_concrete'] and r['thickness_mm'] >= 400
                and r['status'] not in (STATUS_OK, STATUS_EMPTY)]
    if critical:
        output.print_md('**Critical (concrete >= 400 mm, no valid opening): {}**'.format(len(critical)))

    if not results:
        output.print_md('*No intersections found.*')
        return

    output.print_md('## Results table')

    # Статусные иконки
    status_icon = {
        STATUS_NO_OPENING: '[NO OPENING]',
        STATUS_OK:         '[OK]',
        STATUS_UNDERSIZED: '[UNDERSIZED]',
        STATUS_EMPTY:      '[EMPTY]',
    }

    table_data = []
    for r in results:
        mep_size = '{}x{} mm'.format(int(r['mep_w_mm']), int(r['mep_h_mm'])) if r['mep_w_mm'] > 0 else '-'
        thickness = '{} mm'.format(int(r['thickness_mm'])) if r['thickness_mm'] > 0 else '-'
        elev = '{} mm'.format(r['elevation_mm']) if r['elevation_mm'] != 0 else '-'
        mep_id_str   = str(r['mep_id'])   if r['mep_id']   != 0 else '-'
        struct_id_str = str(r['struct_id'])

        table_data.append([
            status_icon.get(r['status'], r['status']),
            r['level'],
            r['mep_system'],
            r['mep_type'],
            mep_size,
            r['struct_type'],
            'Yes' if r['is_concrete'] else 'No',
            thickness,
            r['opening_size'],
            elev,
            mep_id_str,
            struct_id_str,
        ])

    output.print_table(
        table_data,
        title='Opening Check Results (clearance = {} mm)'.format(gap_mm),
        columns=[
            'Status', 'Level', 'MEP Model', 'MEP Type', 'MEP Size',
            'Struct', 'Concrete', 'Thickness',
            'Opening Size', 'Elevation', 'MEP ID', 'Struct ID'
        ]
    )


# =============================================
# STEP 4 — NAVIGATE TO ELEMENT IN 3D VIEW
# =============================================
VIEW_NAME = 'NED_OpeningChecker_View'


def navigate_to_result(result, selected_structural, selected_mep):
    """Создаёт/обновляет 3D вид с Section Box вокруг выбранного пересечения.

    Возвращает ElementId вида или None при ошибке.
    """
    bboxes = []

    # BBox MEP элемента
    if result['mep_id'] != 0:
        for m_link in selected_mep:
            link_doc = m_link['instance'].GetLinkDocument()
            if link_doc is None:
                continue
            try:
                elem = link_doc.GetElement(DB.ElementId(result['mep_id']))
                if elem is not None:
                    bb = elem.get_BoundingBox(None)
                    if bb:
                        bboxes.append(transform_bbox(bb, m_link['instance'].GetTotalTransform()))
                    break
            except Exception:
                pass

    # BBox конструктивного элемента
    for s_link in selected_structural:
        link_doc = s_link['instance'].GetLinkDocument()
        if link_doc is None:
            continue
        try:
            elem = link_doc.GetElement(DB.ElementId(result['struct_id']))
            if elem is not None:
                bb = elem.get_BoundingBox(None)
                if bb:
                    bboxes.append(transform_bbox(bb, s_link['instance'].GetTotalTransform()))
                break
        except Exception:
            pass

    if not bboxes:
        return None

    offset_ft = 1000.0 / FT_TO_MM
    min_x = min(b.Min.X for b in bboxes) - offset_ft
    min_y = min(b.Min.Y for b in bboxes) - offset_ft
    min_z = min(b.Min.Z for b in bboxes) - offset_ft
    max_x = max(b.Max.X for b in bboxes) + offset_ft
    max_y = max(b.Max.Y for b in bboxes) + offset_ft
    max_z = max(b.Max.Z for b in bboxes) + offset_ft

    section_box = DB.BoundingBoxXYZ()
    section_box.Min = DB.XYZ(min_x, min_y, min_z)
    section_box.Max = DB.XYZ(max_x, max_y, max_z)

    try:
        with revit.Transaction('NED: Update opening view'):
            view = None
            for v in DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements():
                if not v.IsTemplate and v.Name == VIEW_NAME:
                    view = v
                    break

            if view is None:
                vft = None
                for vt in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements():
                    if vt.ViewFamily == DB.ViewFamily.ThreeDimensional:
                        vft = vt
                        break
                if vft:
                    view = DB.View3D.CreateIsometric(doc, vft.Id)
                    view.Name = VIEW_NAME

            if view:
                view.SetSectionBox(section_box)
                return view.Id
    except Exception:
        pass

    return None


# =============================================
# STEP 4 — RESULTS NAVIGATOR FORM
# =============================================
_STATUS_COLORS = {
    STATUS_NO_OPENING: (Color.FromArgb(255, 80, 80),  Color.White),
    STATUS_OK:         (Color.FromArgb(120, 195, 60), Color.Black),
    STATUS_UNDERSIZED: (Color.FromArgb(255, 140, 0),  Color.White),
    STATUS_EMPTY:      (Color.FromArgb(255, 230, 0),  Color.Black),
}


class ResultsNavigatorForm(Form):

    def __init__(self, results, selected_structural, selected_mep):
        Form.__init__(self)
        self.all_results = results
        self.selected_structural = selected_structural
        self.selected_mep = selected_mep
        self.filtered = list(results)
        self.target_view_id = None
        self._init_ui()
        self._rebuild_grid()

    # ------------------------------------------------------------------
    # UI build
    # ------------------------------------------------------------------
    def _init_ui(self):
        self.Text = 'NED DC — Opening Checker: Results'
        self.Size = Size(1300, 720)
        self.MinimumSize = Size(1000, 500)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font('Segoe UI', 9)
        self.BackColor = Color.White

        # ---- Top toolbar ----
        toolbar = Panel()
        toolbar.Dock = DockStyle.Top
        toolbar.Height = 86
        toolbar.BackColor = Color.FromArgb(245, 247, 250)
        self.Controls.Add(toolbar)

        x = 10
        def lbl(text, tx, ty):
            l = Label(); l.Text = text; l.Location = Point(tx, ty); l.AutoSize = True
            toolbar.Controls.Add(l)

        # Status filter
        lbl('Status:', x, 13)
        self._cmb_status = ComboBox()
        self._cmb_status.Location = Point(x + 50, 9)
        self._cmb_status.Width = 130
        self._cmb_status.DropDownStyle = ComboBoxStyle.DropDownList
        for v in ['All', STATUS_NO_OPENING, STATUS_OK, STATUS_UNDERSIZED, STATUS_EMPTY]:
            self._cmb_status.Items.Add(v)
        self._cmb_status.SelectedIndex = 0
        self._cmb_status.SelectedIndexChanged += self._on_filter
        toolbar.Controls.Add(self._cmb_status)

        # System filter
        x2 = x + 195
        lbl('System:', x2, 13)
        systems = ['All'] + sorted(set(r['mep_system'] for r in self.all_results if r['mep_system'] != '-'))
        self._cmb_system = ComboBox()
        self._cmb_system.Location = Point(x2 + 55, 9)
        self._cmb_system.Width = 210
        self._cmb_system.DropDownStyle = ComboBoxStyle.DropDownList
        for s in systems:
            self._cmb_system.Items.Add(s)
        self._cmb_system.SelectedIndex = 0
        self._cmb_system.SelectedIndexChanged += self._on_filter
        toolbar.Controls.Add(self._cmb_system)

        # Level filter
        x3 = x2 + 280
        lbl('Level:', x3, 13)
        levels = ['All'] + sorted(set(r['level'] for r in self.all_results))
        self._cmb_level = ComboBox()
        self._cmb_level.Location = Point(x3 + 47, 9)
        self._cmb_level.Width = 160
        self._cmb_level.DropDownStyle = ComboBoxStyle.DropDownList
        for lv in levels:
            self._cmb_level.Items.Add(lv)
        self._cmb_level.SelectedIndex = 0
        self._cmb_level.SelectedIndexChanged += self._on_filter
        toolbar.Controls.Add(self._cmb_level)

        # Concrete only
        x4 = x3 + 225
        self._chk_concrete = CheckBox()
        self._chk_concrete.Text = 'Concrete only'
        self._chk_concrete.Location = Point(x4, 11)
        self._chk_concrete.AutoSize = True
        self._chk_concrete.CheckedChanged += self._on_filter
        toolbar.Controls.Add(self._chk_concrete)

        # Group-by row
        lbl('Group by:', x, 48)
        self._rb_sys = RadioButton()
        self._rb_sys.Text = 'System → Level'
        self._rb_sys.Location = Point(x + 68, 45)
        self._rb_sys.AutoSize = True
        self._rb_sys.Checked = True
        self._rb_sys.CheckedChanged += self._on_filter
        toolbar.Controls.Add(self._rb_sys)

        self._rb_lvl = RadioButton()
        self._rb_lvl.Text = 'Level → System'
        self._rb_lvl.Location = Point(x + 195, 45)
        self._rb_lvl.AutoSize = True
        self._rb_lvl.CheckedChanged += self._on_filter
        toolbar.Controls.Add(self._rb_lvl)

        # Status bar label
        self._lbl_count = Label()
        self._lbl_count.Location = Point(x + 330, 48)
        self._lbl_count.AutoSize = True
        self._lbl_count.ForeColor = Color.FromArgb(90, 90, 90)
        toolbar.Controls.Add(self._lbl_count)

        # Navigate hint label
        self._lbl_nav = Label()
        self._lbl_nav.Location = Point(x + 330, 10)
        self._lbl_nav.AutoSize = True
        self._lbl_nav.ForeColor = Color.FromArgb(30, 90, 160)
        self._lbl_nav.Font = Font('Segoe UI', 9, FontStyle.Bold)
        toolbar.Controls.Add(self._lbl_nav)

        # ---- DataGridView ----
        self._grid = DataGridView()
        self._grid.Dock = DockStyle.Fill
        self._grid.ReadOnly = True
        self._grid.AllowUserToAddRows = False
        self._grid.AllowUserToDeleteRows = False
        self._grid.MultiSelect = False
        self._grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self._grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        self._grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self._grid.ColumnHeadersHeight = 30
        self._grid.RowTemplate.Height = 22
        self._grid.BackgroundColor = Color.White
        self._grid.GridColor = Color.FromArgb(215, 215, 215)
        self._grid.BorderStyle = BorderStyle.None
        self._grid.Font = Font('Segoe UI', 8.5)
        self._grid.RowHeadersVisible = False
        self._grid.CellFormatting += self._on_cell_format
        self._grid.CellDoubleClick += self._on_dbl_click
        self.Controls.Add(self._grid)

        col_defs = [
            ('Status',        110),
            ('Level',         110),
            ('MEP System',    240),
            ('MEP Type',      130),
            ('Struct Type',    90),
            ('Concrete',       75),
            ('Thickness mm',   100),
            ('Opening Size',   115),
            ('Elevation mm',    95),
            ('MEP ID',          80),
            ('Struct ID',       80),
        ]
        for name, width in col_defs:
            col = DataGridViewTextBoxColumn()
            col.HeaderText = name
            col.Name = name
            col.Width = width
            col.SortMode = DataGridViewColumnSortMode.Automatic
            self._grid.Columns.Add(col)

    # ------------------------------------------------------------------
    # Filter & rebuild
    # ------------------------------------------------------------------
    def _on_filter(self, _s, _a):
        self._rebuild_grid()

    def _rebuild_grid(self):
        status_f  = self._cmb_status.SelectedItem  if self._cmb_status.SelectedItem  else 'All'
        system_f  = self._cmb_system.SelectedItem  if self._cmb_system.SelectedItem  else 'All'
        level_f   = self._cmb_level.SelectedItem   if self._cmb_level.SelectedItem   else 'All'
        concrete  = self._chk_concrete.Checked

        data = self.all_results
        if status_f != 'All':
            data = [r for r in data if r['status'] == status_f]
        if system_f != 'All':
            data = [r for r in data if r['mep_system'] == system_f]
        if level_f != 'All':
            data = [r for r in data if r['level'] == level_f]
        if concrete:
            data = [r for r in data if r['is_concrete']]

        if self._rb_sys.Checked:
            data = sorted(data, key=lambda r: (r['mep_system'], r['level'], r['status']))
        else:
            data = sorted(data, key=lambda r: (r['level'], r['mep_system'], r['status']))

        self.filtered = data

        self._grid.SuspendLayout()
        self._grid.Rows.Clear()
        for r in data:
            thickness = '{} mm'.format(int(r['thickness_mm'])) if r['thickness_mm'] > 0 else '-'
            elev      = '{} mm'.format(r['elevation_mm']) if r['elevation_mm'] != 0 else '-'
            mep_id    = str(r['mep_id']) if r['mep_id'] != 0 else '-'
            vals = [
                r['status'], r['level'], r['mep_system'], r['mep_type'],
                r['struct_type'], 'Yes' if r['is_concrete'] else 'No',
                thickness, r['opening_size'], elev, mep_id, str(r['struct_id']),
            ]
            row_idx = self._grid.Rows.Add()
            row = self._grid.Rows[row_idx]
            for ci, v in enumerate(vals):
                row.Cells[ci].Value = v
        self._grid.ResumeLayout()

        totals = {}
        for _st in [STATUS_NO_OPENING, STATUS_OK, STATUS_UNDERSIZED, STATUS_EMPTY]:
            totals[_st] = len([r for r in self.all_results if r['status'] == _st])
        self._lbl_count.Text = (
            'Showing {} / {}  |  No Opening: {}  OK: {}  Undersized: {}  Empty: {}'.format(
                len(data), len(self.all_results),
                totals[STATUS_NO_OPENING], totals[STATUS_OK],
                totals[STATUS_UNDERSIZED], totals[STATUS_EMPTY]
            )
        )

    # ------------------------------------------------------------------
    # Cell coloring
    # ------------------------------------------------------------------
    def _on_cell_format(self, _s, e):
        if e.RowIndex < 0 or e.RowIndex >= len(self.filtered):
            return
        r = self.filtered[e.RowIndex]

        # Status column
        if e.ColumnIndex == 0:
            colors = _STATUS_COLORS.get(r['status'])
            if colors:
                e.CellStyle.BackColor = colors[0]
                e.CellStyle.ForeColor = colors[1]

        # Thickness column
        if e.ColumnIndex == 6 and r['thickness_mm'] > 0:
            t = r['thickness_mm']
            if t >= 400:
                e.CellStyle.BackColor = Color.FromArgb(220, 50, 50)
                e.CellStyle.ForeColor = Color.White
            elif t >= 200:
                e.CellStyle.BackColor = Color.FromArgb(255, 140, 0)
                e.CellStyle.ForeColor = Color.White
            else:
                e.CellStyle.BackColor = Color.FromArgb(255, 230, 0)
                e.CellStyle.ForeColor = Color.Black

    # ------------------------------------------------------------------
    # Double-click → navigate
    # ------------------------------------------------------------------
    def _on_dbl_click(self, _s, e):
        if e.RowIndex < 0 or e.RowIndex >= len(self.filtered):
            return
        r = self.filtered[e.RowIndex]
        try:
            vid = navigate_to_result(r, self.selected_structural, self.selected_mep)
            if vid is not None:
                self.target_view_id = vid
                self._lbl_nav.Text = u'✓ {} updated — close window to switch to it'.format(VIEW_NAME)
            else:
                self._lbl_nav.Text = 'Could not locate elements.'
        except Exception as ex:
            self._lbl_nav.Text = 'Error: {}'.format(ex)


# =============================================
# PYREVIT CONFIG HELPERS
# =============================================
def get_saved_export_path():
    try:
        cfg = script.get_config()
        return cfg.get_option('export_path', '')
    except Exception:
        return ''


def save_export_path(path):
    try:
        cfg = script.get_config()
        cfg.set_option('export_path', path)
        script.save_config()
    except Exception:
        pass


# =============================================
# STEP 1 — MODEL SELECTION DIALOG
# =============================================
class ModelSelectionDialog(Form):

    def __init__(self, structural_links, mep_links, unknown_links):
        Form.__init__(self)
        self.structural_links = structural_links
        self.mep_links = mep_links
        self.unknown_links = unknown_links
        self.selected_structural = []
        self.selected_mep = []
        self.gap_mm = 50
        self.export_path = get_saved_export_path()
        self._init_ui()

    def _make_checkbox(self, text, parent):
        cb = CheckBox()
        cb.Text = text
        cb.AutoSize = True
        cb.Font = Font('Segoe UI', 9)
        cb.Margin = cb.Margin.__class__(4, 2, 4, 2)
        parent.Controls.Add(cb)
        return cb

    def _init_ui(self):
        self.Text = 'NED DC — Opening Checker'
        self.Size = Size(640, 680)
        self.MinimumSize = Size(580, 600)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.Font = Font('Segoe UI', 9)
        self.BackColor = Color.White

        title = Label()
        title.Text = 'Opening Checker'
        title.Font = Font('Segoe UI', 13, FontStyle.Bold)
        title.ForeColor = Color.FromArgb(30, 90, 160)
        title.Location = Point(16, 14)
        title.AutoSize = True
        self.Controls.Add(title)

        subtitle = Label()
        subtitle.Text = 'Select models and configure check parameters'
        subtitle.Font = Font('Segoe UI', 9)
        subtitle.ForeColor = Color.Gray
        subtitle.Location = Point(16, 40)
        subtitle.AutoSize = True
        self.Controls.Add(subtitle)

        y = 68

        grp_struct = GroupBox()
        grp_struct.Text = 'Structural models (AR / ST / Openings)'
        grp_struct.Font = Font('Segoe UI', 9, FontStyle.Bold)
        grp_struct.Location = Point(12, y)
        grp_struct.Size = Size(608, 160)
        grp_struct.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
        self.Controls.Add(grp_struct)

        self._struct_panel = FlowLayoutPanel()
        self._struct_panel.FlowDirection = FlowDirection.TopDown
        self._struct_panel.AutoScroll = True
        self._struct_panel.Location = Point(8, 20)
        self._struct_panel.Size = Size(590, 130)
        self._struct_panel.WrapContents = False
        grp_struct.Controls.Add(self._struct_panel)

        self._struct_checkboxes = []
        links_to_show = self.structural_links + self.unknown_links
        if not links_to_show:
            lbl = Label()
            lbl.Text = 'No structural models found'
            lbl.ForeColor = Color.Gray
            lbl.AutoSize = True
            self._struct_panel.Controls.Add(lbl)
        else:
            for link in links_to_show:
                cb = self._make_checkbox(link['name'], self._struct_panel)
                cb.Checked = link['category'] == 'structural'
                cb.Tag = link
                self._struct_checkboxes.append(cb)

        y += 170

        grp_mep = GroupBox()
        grp_mep.Text = 'MEP models (HVAC / Plumbing / Electrical / Fuel)'
        grp_mep.Font = Font('Segoe UI', 9, FontStyle.Bold)
        grp_mep.Location = Point(12, y)
        grp_mep.Size = Size(608, 160)
        grp_mep.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
        self.Controls.Add(grp_mep)

        self._mep_panel = FlowLayoutPanel()
        self._mep_panel.FlowDirection = FlowDirection.TopDown
        self._mep_panel.AutoScroll = True
        self._mep_panel.Location = Point(8, 20)
        self._mep_panel.Size = Size(590, 130)
        self._mep_panel.WrapContents = False
        grp_mep.Controls.Add(self._mep_panel)

        self._mep_checkboxes = []
        if not self.mep_links:
            lbl = Label()
            lbl.Text = 'No MEP models found'
            lbl.ForeColor = Color.Gray
            lbl.AutoSize = True
            self._mep_panel.Controls.Add(lbl)
        else:
            for link in self.mep_links:
                cb = self._make_checkbox(link['name'], self._mep_panel)
                cb.Checked = True
                cb.Tag = link
                self._mep_checkboxes.append(cb)

        y += 170

        grp_settings = GroupBox()
        grp_settings.Text = 'Check settings'
        grp_settings.Font = Font('Segoe UI', 9, FontStyle.Bold)
        grp_settings.Location = Point(12, y)
        grp_settings.Size = Size(608, 110)
        self.Controls.Add(grp_settings)

        lbl_gap = Label()
        lbl_gap.Text = 'Minimum clearance (mm):'
        lbl_gap.Location = Point(10, 26)
        lbl_gap.AutoSize = True
        grp_settings.Controls.Add(lbl_gap)

        self._txt_gap = TextBox()
        self._txt_gap.Text = '50'
        self._txt_gap.Location = Point(200, 23)
        self._txt_gap.Size = Size(70, 23)
        grp_settings.Controls.Add(self._txt_gap)

        lbl_gap_hint = Label()
        lbl_gap_hint.Text = 'mm on each side of MEP element'
        lbl_gap_hint.ForeColor = Color.Gray
        lbl_gap_hint.Location = Point(278, 26)
        lbl_gap_hint.AutoSize = True
        grp_settings.Controls.Add(lbl_gap_hint)

        lbl_path = Label()
        lbl_path.Text = 'Excel report folder:'
        lbl_path.Location = Point(10, 60)
        lbl_path.AutoSize = True
        grp_settings.Controls.Add(lbl_path)

        self._txt_path = TextBox()
        self._txt_path.Text = self.export_path
        self._txt_path.Location = Point(200, 57)
        self._txt_path.Size = Size(300, 23)
        self._txt_path.ScrollBars = ScrollBars.Horizontal
        grp_settings.Controls.Add(self._txt_path)

        btn_browse = Button()
        btn_browse.Text = 'Browse...'
        btn_browse.Location = Point(508, 56)
        btn_browse.Size = Size(80, 25)
        btn_browse.Click += self._on_browse
        grp_settings.Controls.Add(btn_browse)

        y += 120

        btn_run = Button()
        btn_run.Text = 'Run check'
        btn_run.Font = Font('Segoe UI', 10, FontStyle.Bold)
        btn_run.Size = Size(160, 36)
        btn_run.Location = Point(12, y + 8)
        btn_run.BackColor = Color.FromArgb(30, 90, 160)
        btn_run.ForeColor = Color.White
        btn_run.FlatStyle = btn_run.FlatStyle.__class__.Flat
        btn_run.Click += self._on_run
        self.Controls.Add(btn_run)

        btn_diag = Button()
        btn_diag.Text = 'Diagnose types'
        btn_diag.Size = Size(130, 36)
        btn_diag.Location = Point(180, y + 8)
        btn_diag.Click += self._on_diagnose
        self.Controls.Add(btn_diag)

        btn_cancel = Button()
        btn_cancel.Text = 'Cancel'
        btn_cancel.Size = Size(100, 36)
        btn_cancel.Location = Point(318, y + 8)
        btn_cancel.Click += self._on_cancel
        self.Controls.Add(btn_cancel)

        self.ClientSize = Size(640, y + 60)

    def _on_browse(self, _s, _a):
        dlg = FolderBrowserDialog()
        dlg.Description = 'Select folder for Excel report'
        if self.export_path and os.path.exists(self.export_path):
            dlg.SelectedPath = self.export_path
        if dlg.ShowDialog() == DialogResult.OK:
            self._txt_path.Text = dlg.SelectedPath

    def _on_run(self, _s, _a):
        self.selected_structural = [cb.Tag for cb in self._struct_checkboxes if cb.Checked]
        self.selected_mep        = [cb.Tag for cb in self._mep_checkboxes if cb.Checked]

        if not self.selected_structural:
            MessageBox.Show('Please select at least one structural model.',
                            'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if not self.selected_mep:
            MessageBox.Show('Please select at least one MEP model.',
                            'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        try:
            self.gap_mm = int(self._txt_gap.Text.strip())
            if self.gap_mm < 0:
                raise ValueError
        except ValueError:
            MessageBox.Show('Please enter a valid clearance value (integer >= 0).',
                            'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        self.export_path = self._txt_path.Text.strip()
        if self.export_path:
            save_export_path(self.export_path)

        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_diagnose(self, _s, _a):
        """Запускает диагностику типов стен без проведения полной проверки."""
        struct_links = [cb.Tag for cb in self._struct_checkboxes if cb.Checked]
        if not struct_links:
            MessageBox.Show('Please select at least one structural model.',
                            'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.selected_structural = struct_links
        self.DialogResult = DialogResult.Retry   # используем Retry как сигнал диагностики
        self.Close()

    def _on_cancel(self, _s, _a):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# =============================================
# STEP 5 — EXCEL EXPORT
# =============================================
def _link_code(link_name):
    """Возвращает код дисциплины из позиции [2] имени файла."""
    parts = link_name.replace('.rvt', '').replace('.RVT', '').split('-')
    return parts[2] if len(parts) >= 3 else link_name.split('.')[0]


def export_to_excel(results, export_folder, gap_mm, selected_structural, selected_mep):
    """Экспортирует результаты в xlsx файл по формату ТЗ (через xlsxwriter)."""
    try:
        import xlsxwriter
    except ImportError:
        return None, 'xlsxwriter is not available in this pyRevit installation'

    if not os.path.exists(export_folder):
        try:
            os.makedirs(export_folder)
        except Exception as e:
            return None, 'Cannot create folder: {}'.format(e)

    now = datetime.datetime.now()

    struct_tag = '+'.join(_link_code(l['name']) for l in selected_structural)
    mep_tag    = '+'.join(_link_code(l['name']) for l in selected_mep)
    filename = 'NED_OpeningCheck_{}_{}_{}.xlsx'.format(
        now.strftime('%Y-%m-%d_%H-%M'), struct_tag, mep_tag
    )
    filepath = os.path.join(export_folder, filename)

    try:
        wb = xlsxwriter.Workbook(filepath)
    except Exception as e:
        return None, str(e)

    ws = wb.add_worksheet('Opening Check Report')

    # --- Форматы ---
    BASE = {'valign': 'vcenter', 'border': 1, 'font_name': 'Calibri', 'font_size': 10}

    def fmt(extra):
        d = dict(BASE)
        d.update(extra)
        return wb.add_format(d)

    hdr_fmt = fmt({'bold': True, 'font_color': 'white', 'bg_color': '#1E5AA0',
                   'align': 'center', 'text_wrap': True})
    data_fmt = fmt({})

    # Форматы статусов (колонка Status)
    STATUS_FMTS = {
        STATUS_NO_OPENING: fmt({'bg_color': '#FF4444'}),
        STATUS_OK:         fmt({'bg_color': '#92D050'}),
        STATUS_UNDERSIZED: fmt({'bg_color': '#FF8C00'}),
        STATUS_EMPTY:      fmt({'bg_color': '#FFFF00'}),
    }

    # Форматы толщины (колонка Thickness)
    thick_red    = fmt({'bg_color': '#FF0000'})
    thick_orange = fmt({'bg_color': '#FF8C00'})
    thick_yellow = fmt({'bg_color': '#FFFF00'})

    # --- Колонки: (заголовок, ширина) ---
    columns = [
        ('Status',                    14),
        ('Level',                     12),
        ('MEP System',                30),
        ('Element Type',              16),
        ('MEP Element ID',            16),
        ('Structure Type',            14),
        ('Is Concrete',               12),
        ('Wall/Floor Type',           28),
        ('Thickness (mm)',            14),
        ('Opening Size',              16),
        ('Elevation from Level (mm)', 24),
        ('Structure Element ID',      20),
        ('Approval Status',           16),
        ('Approval Date',             14),
        ('Comment',                   30),
    ]

    # Записываем заголовки (строка 0)
    for col_idx, (header, width) in enumerate(columns):
        ws.write(0, col_idx, header, hdr_fmt)
        ws.set_column(col_idx, col_idx, width)
    ws.set_row(0, 32)

    # Записываем данные (строки 1+)
    for row_idx, r in enumerate(results, 1):
        thickness = int(round(r['thickness_mm'])) if r['thickness_mm'] > 0 else 0
        mep_id    = r['mep_id']      if r['mep_id']      != 0 else ''
        elev      = r['elevation_mm'] if r['elevation_mm'] != 0 else ''

        row_values = [
            r['status'],
            r['level'],
            r['mep_system'],
            r['mep_type'],
            mep_id,
            r['struct_type'],
            'Yes' if r['is_concrete'] else 'No',
            r['type_name'],
            thickness,
            r['opening_size'],
            elev,
            r['struct_id'],
            '',   # Approval Status — Step 3
            '',   # Approval Date   — Step 3
            '',   # Comment         — Step 3
        ]

        for col_idx, value in enumerate(row_values):
            if col_idx == 0:
                cell_fmt = STATUS_FMTS.get(r['status'], data_fmt)
            elif col_idx == 8:
                if thickness >= 400:
                    cell_fmt = thick_red
                elif thickness >= 200:
                    cell_fmt = thick_orange
                elif thickness > 0:
                    cell_fmt = thick_yellow
                else:
                    cell_fmt = data_fmt
            else:
                cell_fmt = data_fmt
            ws.write(row_idx, col_idx, value, cell_fmt)

    # Автофильтр и заморозка заголовка
    ws.autofilter(0, 0, len(results), len(columns) - 1)
    ws.freeze_panes(1, 0)

    try:
        wb.close()
        return filepath, None
    except Exception as e:
        return None, str(e)


# =============================================
# ENTRY POINT
# =============================================
def main():
    all_links = get_all_revit_links()
    if not all_links:
        forms.alert(
            'No Revit Links found in the current document.\n'
            'Please open a host model with linked files.',
            title='NED DC — Opening Checker'
        )
        return

    structural_links = [l for l in all_links if l['category'] == 'structural']
    mep_links        = [l for l in all_links if l['category'] == 'mep']
    unknown_links    = [l for l in all_links if l['category'] == 'unknown']

    dlg = ModelSelectionDialog(structural_links, mep_links, unknown_links)
    result = dlg.ShowDialog()

    if result == DialogResult.Cancel:
        return

    output = script.get_output()
    output.print_md('# NED DC — Opening Checker')

    # Режим диагностики: только вывод типов стен, без полной проверки
    if result == DialogResult.Retry:
        diagnose_wall_types(dlg.selected_structural, output)
        return
    output.print_md('**Structural:** {}'.format(
        ', '.join(l['name'] for l in dlg.selected_structural)))
    output.print_md('**MEP:** {}'.format(
        ', '.join(l['name'] for l in dlg.selected_mep)))
    output.print_md('**Clearance:** {} mm'.format(dlg.gap_mm))
    output.print_md('---')

    # Шаг 2: проверка пересечений
    results = run_check(dlg.selected_structural, dlg.selected_mep, dlg.gap_mm, output)

    output.print_md('---')
    print_results(results, output, dlg.gap_mm)

    # Шаг 5: экспорт в Excel
    if not results:
        return

    export_folder = dlg.export_path
    if not export_folder:
        # Папка не указана — предлагаем выбрать сейчас
        from System.Windows.Forms import FolderBrowserDialog
        fb = FolderBrowserDialog()
        fb.Description = 'Select folder to save Excel report'
        from System.Windows.Forms import DialogResult as DR
        if fb.ShowDialog() == DR.OK:
            export_folder = fb.SelectedPath
            save_export_path(export_folder)
        else:
            output.print_md('*Excel export skipped — no folder selected.*')
            return

    output.print_md('---')
    output.print_md('**Saving Excel report...**')
    filepath, err = export_to_excel(
        results, export_folder, dlg.gap_mm,
        dlg.selected_structural, dlg.selected_mep
    )

    if err:
        output.print_md('**Export error:** {}'.format(err))
    else:
        output.print_md('**Report saved:** `{}`'.format(filepath))
        try:
            import System.Diagnostics
            System.Diagnostics.Process.Start(filepath)
        except Exception:
            pass

    # Шаг 4: навигатор результатов
    output.print_md('---')
    output.print_md('*Opening Results Navigator...*')
    nav = ResultsNavigatorForm(results, dlg.selected_structural, dlg.selected_mep)
    nav.ShowDialog()

    # После закрытия навигатора — переключаемся на 3D вид если был создан
    if nav.target_view_id is not None:
        try:
            view = doc.GetElement(nav.target_view_id)
            if view:
                revit.uidoc.ActiveView = view
        except Exception:
            pass


main()
