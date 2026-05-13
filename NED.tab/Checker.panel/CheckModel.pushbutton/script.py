# -*- coding: utf-8 -*-
__title__ = 'Проверка отверстий'
__doc__ = 'Проверка наличия и размеров отверстий для MEP элементов в конструктивных моделях'
__author__ = 'NED DC'

import clr
import os
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form, Label, CheckBox, TextBox, Button,
    FlowLayoutPanel, GroupBox,
    FormBorderStyle, DialogResult,
    FlowDirection, AnchorStyles,
    FolderBrowserDialog, MessageBox, MessageBoxButtons,
    MessageBoxIcon, FormStartPosition,
    ScrollBars
)
from System.Drawing import (
    Point, Size, Font, FontStyle, Color
)

from pyrevit import revit, DB, script, forms

doc = revit.doc

# Коды дисциплин для классификации моделей
STRUCTURAL_CODES = ['AR', 'AR-D', 'ST', 'O', 'OP', 'S']
MEP_CODES        = ['HV', 'PL', 'EL', 'COM', 'FL', 'H', 'P', 'E', 'T', 'F', 'F-LM']
SKIP_CODES       = ['TR', 'SI', 'CO', 'CR', 'G', 'Z', 'B', 'FU', 'ID', 'L', 'Q', 'W', 'K']

STRUCTURAL_NUMBERS = [100, 150, 200]
MEP_NUMBERS        = [300, 400, 500, 550]
SKIP_NUMBERS       = [350, 600]


def classify_link(link_name):
    """Определяет тип модели по имени файла Revit Link.

    Формат имени: S-HA-[КОД]-[КОМПАНИЯ]-[ЛОКАЦИЯ]-RVT2X
    Возвращает: 'structural', 'mep' или 'skip'
    """
    name_upper = link_name.upper()
    parts = link_name.replace('.rvt', '').replace('.RVT', '').split('-')

    # Проверяем числовые коды в имени файла
    for part in parts:
        try:
            num = int(part)
            if num in STRUCTURAL_NUMBERS:
                return 'structural'
            if num in MEP_NUMBERS:
                return 'mep'
            if num in SKIP_NUMBERS:
                return 'skip'
        except ValueError:
            pass

    # Проверяем буквенные коды (ищем сегменты имени файла)
    for part in parts:
        part_clean = part.strip().upper()
        # Сначала пропускаемые коды
        if part_clean in [c.upper() for c in SKIP_CODES]:
            return 'skip'

    for part in parts:
        part_clean = part.strip().upper()
        if part_clean in [c.upper() for c in STRUCTURAL_CODES]:
            return 'structural'
        # Проверяем MEP коды (H, P, E, F — однобуквенные)
        if part_clean in [c.upper() for c in MEP_CODES]:
            return 'mep'

    # Дополнительная проверка по характерным подстрокам
    for code in SKIP_CODES:
        if ('-' + code.upper() + '-') in name_upper:
            return 'skip'
    for code in STRUCTURAL_CODES:
        if ('-' + code.upper() + '-') in name_upper:
            return 'structural'
    for code in MEP_CODES:
        if ('-' + code.upper() + '-') in name_upper:
            return 'mep'

    return 'unknown'


def get_all_revit_links():
    """Получает все подключённые Revit Links из текущего документа."""
    collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.RevitLinkInstance)\
        .ToElements()

    links = []
    for link_instance in collector:
        link_type = doc.GetElement(link_instance.GetTypeId())
        if link_type is None:
            continue
        # Получаем имя файла без полного пути
        link_name = link_type.get_Parameter(
            DB.BuiltInParameter.ALL_MODEL_TYPE_NAME
        ).AsString()
        if not link_name:
            link_name = link_instance.Name

        category = classify_link(link_name)
        links.append({
            'name': link_name,
            'instance': link_instance,
            'category': category
        })

    return links


def get_saved_export_path():
    """Читает сохранённый путь экспорта из pyRevit config."""
    try:
        cfg = script.get_config()
        return cfg.get_option('export_path', '')
    except Exception:
        return ''


def save_export_path(path):
    """Сохраняет путь экспорта в pyRevit config."""
    try:
        cfg = script.get_config()
        cfg.set_option('export_path', path)
        script.save_config()
    except Exception:
        pass


class ModelSelectionDialog(Form):
    """Диалог выбора конструктивных и MEP моделей для проверки."""

    def __init__(self, structural_links, mep_links, unknown_links):
        Form.__init__(self)
        self.structural_links = structural_links
        self.mep_links = mep_links
        self.unknown_links = unknown_links

        # Результаты выбора
        self.selected_structural = []
        self.selected_mep = []
        self.gap_mm = 50
        self.export_path = get_saved_export_path()

        self._init_ui()

    def _make_header(self, text, parent):
        """Создаёт заголовок секции."""
        lbl = Label()
        lbl.Text = text
        lbl.Font = Font('Segoe UI', 10, FontStyle.Bold)
        lbl.ForeColor = Color.FromArgb(30, 90, 160)
        lbl.AutoSize = True
        lbl.Margin.Bottom = 4
        parent.Controls.Add(lbl)
        return lbl

    def _make_checkbox(self, text, parent):
        """Создаёт чекбокс с переданным текстом."""
        cb = CheckBox()
        cb.Text = text
        cb.AutoSize = True
        cb.Font = Font('Segoe UI', 9)
        cb.Margin = cb.Margin.__class__(4, 2, 4, 2)
        parent.Controls.Add(cb)
        return cb

    def _init_ui(self):
        self.Text = 'NED DC — Проверка отверстий'
        self.Size = Size(640, 680)
        self.MinimumSize = Size(580, 600)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.Font = Font('Segoe UI', 9)
        self.BackColor = Color.White

        # --- Заголовок ---
        title = Label()
        title.Text = 'Проверка отверстий'
        title.Font = Font('Segoe UI', 13, FontStyle.Bold)
        title.ForeColor = Color.FromArgb(30, 90, 160)
        title.Location = Point(16, 14)
        title.AutoSize = True
        self.Controls.Add(title)

        subtitle = Label()
        subtitle.Text = 'Выберите модели и настройте параметры проверки'
        subtitle.Font = Font('Segoe UI', 9)
        subtitle.ForeColor = Color.Gray
        subtitle.Location = Point(16, 40)
        subtitle.AutoSize = True
        self.Controls.Add(subtitle)

        y = 68

        # --- Группа: конструктивные модели ---
        grp_struct = GroupBox()
        grp_struct.Text = 'Конструктивные модели (АР / КР / Отверстия)'
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
        links_to_show = self.structural_links + [
            l for l in self.unknown_links
        ]
        if not links_to_show:
            lbl = Label()
            lbl.Text = 'Конструктивные модели не найдены'
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

        # --- Группа: MEP модели ---
        grp_mep = GroupBox()
        grp_mep.Text = 'Инженерные модели (MEP)'
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
            lbl.Text = 'MEP модели не найдены'
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

        # --- Группа: параметры ---
        grp_settings = GroupBox()
        grp_settings.Text = 'Параметры проверки'
        grp_settings.Font = Font('Segoe UI', 9, FontStyle.Bold)
        grp_settings.Location = Point(12, y)
        grp_settings.Size = Size(608, 110)
        self.Controls.Add(grp_settings)

        # Зазор
        lbl_gap = Label()
        lbl_gap.Text = 'Минимальный зазор (мм):'
        lbl_gap.Location = Point(10, 26)
        lbl_gap.AutoSize = True
        grp_settings.Controls.Add(lbl_gap)

        self._txt_gap = TextBox()
        self._txt_gap.Text = '50'
        self._txt_gap.Location = Point(200, 23)
        self._txt_gap.Size = Size(70, 23)
        grp_settings.Controls.Add(self._txt_gap)

        lbl_gap_hint = Label()
        lbl_gap_hint.Text = 'мм с каждой стороны от MEP элемента'
        lbl_gap_hint.ForeColor = Color.Gray
        lbl_gap_hint.Location = Point(278, 26)
        lbl_gap_hint.AutoSize = True
        grp_settings.Controls.Add(lbl_gap_hint)

        # Папка экспорта
        lbl_path = Label()
        lbl_path.Text = 'Папка для Excel отчёта:'
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
        btn_browse.Text = 'Обзор...'
        btn_browse.Location = Point(508, 56)
        btn_browse.Size = Size(80, 25)
        btn_browse.Click += self._on_browse
        grp_settings.Controls.Add(btn_browse)

        y += 120

        # --- Кнопки ---
        btn_run = Button()
        btn_run.Text = 'Запустить проверку'
        btn_run.Font = Font('Segoe UI', 10, FontStyle.Bold)
        btn_run.Size = Size(200, 36)
        btn_run.Location = Point(12, y + 8)
        btn_run.BackColor = Color.FromArgb(30, 90, 160)
        btn_run.ForeColor = Color.White
        btn_run.FlatStyle = btn_run.FlatStyle.__class__.Flat
        btn_run.Click += self._on_run
        self.Controls.Add(btn_run)

        btn_cancel = Button()
        btn_cancel.Text = 'Отмена'
        btn_cancel.Size = Size(100, 36)
        btn_cancel.Location = Point(220, y + 8)
        btn_cancel.Click += self._on_cancel
        self.Controls.Add(btn_cancel)

        self.ClientSize = Size(640, y + 60)

    def _on_browse(self, _s, _a):
        """Открывает диалог выбора папки."""
        dlg = FolderBrowserDialog()
        dlg.Description = 'Выберите папку для сохранения Excel отчёта'
        if self.export_path and os.path.exists(self.export_path):
            dlg.SelectedPath = self.export_path
        result = dlg.ShowDialog()
        if result == DialogResult.OK:
            self._txt_path.Text = dlg.SelectedPath

    def _on_run(self, sender, args):
        """Собирает результаты выбора и закрывает диалог."""
        # Собираем выбранные конструктивные модели
        self.selected_structural = [
            cb.Tag for cb in self._struct_checkboxes if cb.Checked
        ]
        # Собираем выбранные MEP модели
        self.selected_mep = [
            cb.Tag for cb in self._mep_checkboxes if cb.Checked
        ]

        if not self.selected_structural:
            MessageBox.Show(
                'Выберите хотя бы одну конструктивную модель.',
                'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        if not self.selected_mep:
            MessageBox.Show(
                'Выберите хотя бы одну MEP модель.',
                'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        # Проверяем зазор
        try:
            self.gap_mm = int(self._txt_gap.Text.strip())
            if self.gap_mm < 0:
                raise ValueError
        except ValueError:
            MessageBox.Show(
                'Введите корректное значение зазора (целое число ≥ 0).',
                'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        # Сохраняем путь экспорта
        self.export_path = self._txt_path.Text.strip()
        if self.export_path:
            save_export_path(self.export_path)

        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


def main():
    # Получаем все Revit Links и классифицируем их
    all_links = get_all_revit_links()

    if not all_links:
        forms.alert(
            'В текущем документе нет подключённых Revit Links.\n'
            'Откройте рабочую модель с подключёнными файлами.',
            title='NED DC — Проверка отверстий'
        )
        return

    structural_links = [l for l in all_links if l['category'] == 'structural']
    mep_links        = [l for l in all_links if l['category'] == 'mep']
    unknown_links    = [l for l in all_links if l['category'] == 'unknown']
    # Пропускаемые модели (TR, SI и т.д.) не показываем

    # Показываем диалог выбора
    dlg = ModelSelectionDialog(structural_links, mep_links, unknown_links)
    result = dlg.ShowDialog()

    if result != DialogResult.OK:
        return

    # Выводим итог выбора в Output
    output = script.get_output()
    output.print_md('# NED DC — Проверка отверстий')
    output.print_md('## Выбранные модели')

    output.print_md('### Конструктивные:')
    for link in dlg.selected_structural:
        output.print_md('- {}'.format(link['name']))

    output.print_md('### MEP:')
    for link in dlg.selected_mep:
        output.print_md('- {}'.format(link['name']))

    output.print_md('**Зазор:** {} мм'.format(dlg.gap_mm))
    if dlg.export_path:
        output.print_md('**Папка отчёта:** {}'.format(dlg.export_path))

    output.print_md('---')
    output.print_md('_Шаг 1 завершён. Логика проверки будет реализована на следующем шаге._')


main()
