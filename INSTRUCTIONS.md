# NED DC — pyRevit Extension: DC_rvt_checker

## О проекте
Расширение для Autodesk Revit, разработанное для компании **NED DC**.
Добавляет вкладку **NED** на ленту Revit с инструментами проверки и автоматизации BIM-моделей.

---

## Техническое окружение
- **Autodesk Revit**: 2025
- **pyRevit**: последняя версия (установлена без прав администратора)
- **Python**: IronPython 2.7 (используется внутри pyRevit)
- **Репозиторий**: https://github.com/mkirsanov91/DC_rvt_checker.extension
- **Локальный путь**: `C:\Users\Michaelkirsanov\AppData\Roaming\pyRevit\Extensions\DC_rvt_checker.extension`

---

## Структура репозитория
```
DC_rvt_checker.extension/
└── NED.tab/                        ← вкладка на ленте Revit
    └── [Название].panel/           ← панель на вкладке
        └── [Название].pushbutton/  ← кнопка
            ├── script.py           ← основной скрипт
            ├── bundle.yaml         ← название и описание кнопки
            └── icon.png            ← иконка кнопки (опционально, 32x32px)
```

### Типы bundle (папок):
- `.pushbutton` — обычная кнопка
- `.pulldown` — кнопка с выпадающим списком
- `.stack` — стек из нескольких кнопок

---

## Правила написания скриптов

### Язык
- Комментарии в коде — **на русском**
- Названия кнопок (`__title__`) — **на русском**
- Документация (`__doc__`) — **на русском**
- Перед релизом комментарии переводятся на английский

### Заголовок каждого script.py
```python
# -*- coding: utf-8 -*-
__title__ = 'Название кнопки'
__doc__ = 'Описание что делает кнопка'
__author__ = 'NED DC'
```

### Импорты
```python
# Стандартные импорты pyRevit
from pyrevit import revit, DB, UI, script, forms

# Работа с документом
doc = revit.doc
uidoc = revit.uidoc
app = revit.app
```

### Вывод результатов
```python
# Для вывода в окно pyRevit Output
output = script.get_output()
output.print_md('# Заголовок')
output.print_md('Текст результата')

# Для простых диалогов
forms.alert('Сообщение', title='NED DC')
```

### Обработка ошибок
```python
# Транзакция для изменения модели
with revit.Transaction('Название действия'):
    # код изменения элементов
    pass
```

---

## Workflow разработки

### Тестирование без перезапуска Revit
1. Отредактировал `script.py` в VS Code
2. В Revit: **pyRevit → Reload** (или Alt+F5)
3. Нажал кнопку — проверил результат

### Сохранение изменений на GitHub
```bash
# В терминале VS Code (папка расширения)
git add .
git commit -m "описание что сделано"
git push
```

### Настройка git (один раз, чтобы коммиты отображались в contributions на GitHub)
```bash
git config --global user.name "mkirsanov91"
git config --global user.email "твой@email.com"  # тот же email что в github.com → Settings → Emails
```

### Быстрые команды git
```bash
git status          # посмотреть что изменилось
git add .           # добавить все изменения
git commit -m "..."  # зафиксировать
git push            # отправить на GitHub
git pull            # получить обновления с GitHub
```

---

## Полезные ссылки
- [pyRevit документация](https://pyrevitlabs.notion.site)
- [Revit API документация](https://www.revitapidocs.com)
- [GitHub репозиторий](https://github.com/mkirsanov91/DC_rvt_checker.extension)
