# -*- coding: utf-8 -*-
__title__ = 'Проверка отверстий'
__doc__ = 'Check presence and size of openings for MEP elements in structural models'
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

# Коды дисциплин на позиции [2] в имени файла: S-HA-[КОД]-[КОМПАНИЯ]-[ЛОКАЦИЯ]-RVT2X
STRUCTURAL_FILE_CODES = ['AR', 'S', 'ST', 'O', 'OP']
MEP_FILE_CODES        = ['H', 'P', 'E', 'F', 'T', 'HV', 'PL', 'EL', 'FL']
SKIP_FILE_CODES       = ['TR', 'SI', 'CO', 'CR', 'G', 'Z', 'B', 'FU', 'ID']


def classify_link(link_name):
    """Определяет тип модели по имени файла Revit Link.

    Формат: S-HA-[КОД]-[КОМПАНИЯ]-[ЛОКАЦИЯ]-RVT2X
    Дисциплинарный код — позиция [2] при разбивке по дефису.
    """
    name_clean = link_name.replace('.rvt', '').replace('.RVT', '')
    parts = name_clean.split('-')

    # Основной путь: берём код с позиции [2]
    if len(parts) >= 3:
        discipline = parts[2].upper()
        if discipline in [c.upper() for c in SKIP_FILE_CODES]:
            return 'skip'
        if discipline in [c.upper() for c in STRUCTURAL_FILE_CODES]:
            return 'structural'
        if discipline in [c.upper() for c in MEP_FILE_CODES]:
            return 'mep'

    # Запасной путь для нестандартных имён: сканируем все позиции
    for part in parts:
        if part.upper() in [c.upper() for c in SKIP_FILE_CODES]:
            return 'skip'
    for part in parts:
        if part.upper() in [c.upper() for c in MEP_FILE_CODES]:
            return 'mep'
    for part in parts:
        if part.upper() in [c.upper() for c in STRUCTURAL_FILE_CODES]:
            return 'structural'

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

        # --- Structural models ---
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

        # --- MEP models ---
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

        # --- Settings ---
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

        btn_cancel = Button()
        btn_cancel.Text = 'Cancel'
        btn_cancel.Size = Size(100, 36)
        btn_cancel.Location = Point(180, y + 8)
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
        self.selected_structural = [
            cb.Tag for cb in self._struct_checkboxes if cb.Checked
        ]
        self.selected_mep = [
            cb.Tag for cb in self._mep_checkboxes if cb.Checked
        ]

        if not self.selected_structural:
            MessageBox.Show(
                'Please select at least one structural model.',
                'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        if not self.selected_mep:
            MessageBox.Show(
                'Please select at least one MEP model.',
                'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        try:
            self.gap_mm = int(self._txt_gap.Text.strip())
            if self.gap_mm < 0:
                raise ValueError
        except ValueError:
            MessageBox.Show(
                'Please enter a valid clearance value (integer >= 0).',
                'NED DC', MessageBoxButtons.OK, MessageBoxIcon.Warning
            )
            return

        self.export_path = self._txt_path.Text.strip()
        if self.export_path:
            save_export_path(self.export_path)

        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_cancel(self, _s, _a):
        self.DialogResult = DialogResult.Cancel
        self.Close()


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
    if dlg.ShowDialog() != DialogResult.OK:
        return

    output = script.get_output()
    output.print_md('# NED DC — Opening Checker')
    output.print_md('## Selected models')
    output.print_md('### Structural:')
    for link in dlg.selected_structural:
        output.print_md('- {}'.format(link['name']))
    output.print_md('### MEP:')
    for link in dlg.selected_mep:
        output.print_md('- {}'.format(link['name']))
    output.print_md('**Clearance:** {} mm'.format(dlg.gap_mm))
    if dlg.export_path:
        output.print_md('**Report folder:** {}'.format(dlg.export_path))
    output.print_md('---')
    output.print_md('_Step 1 complete. Intersection logic will be implemented next._')


main()
