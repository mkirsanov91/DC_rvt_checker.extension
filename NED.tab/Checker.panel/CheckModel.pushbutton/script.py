# -*- coding: utf-8 -*-
__title__ = 'Check Model'
__doc__ = 'Проверка модели'

from pyrevit import script

output = script.get_output()
output.print_md('# Привет из NED!')
