# Контекст для нового чата — COREZ Word/PDF форматирование

## Проект
`/home/user/COREZ` — Flask-приложение для управления соревнованиями по рафтингу.
Ветка разработки: `main`

## Ключевые файлы
- `docx_utils.py` — все функции генерации Word/PDF
- `routes/спринт.py`, `routes/слалом.py`, `routes/документы.py` — маршруты
- `models.py` — модели данных

## Модели данных (models.py)
- `Соревнование` — соревнование: `название`, `организаторы_блок`, `даты_проведения`, `место_проведения`, `главный_судья`, `главный_секретарь`, `последовательность_видов` ("R4-м,R4-ж,R6-м,R6-ж"), `класс_судна`
- `Команда` — команда: `нагрудный_номер`, `название`, `класс_судна` (R4/R6), `пол` (м/ж), `субъект` → FK
- `Субъект` — регион: `название`, `представитель`
- `Спортсмен` — спортсмен: `фамилия_имя_отчество`, `номер_в_экипаже`

## Уже реализованные функции в docx_utils.py
| Функция | Маршрут | Описание |
|---|---|---|
| `create_старт_спринт_docx(сор, группы, filename)` | `спринт.протокол_старт_word` | Стартовый протокол Спринт |
| `create_старт_слалом_docx(сор, секции, filename)` | `слалом.протокол_старт_word` | Стартовый протокол Слалом |
| `create_результаты_слалом_docx(...)` | `слалом.протокол_результаты_word` | Протокол результатов Слалом |
| `create_нагрудные_номера_docx(сор, команды, filename)` | `документы.нагрудные_номера_word` | Протокол выдачи нагрудных номеров |
| `send_docx(html, filename, ...)` | остальные маршруты | HTML→DOCX конвертация (fallback) |

## Ключевые технические паттерны в docx_utils.py

### Страница
```python
sect.page_width    = Mm(297)   # A4 альбомная
sect.page_height   = Mm(210)
sect.left_margin   = Mm(15)
sect.right_margin  = Mm(10)
sect.top_margin    = Mm(15)
sect.bottom_margin = Mm(15)
pgSz.set(qn('w:orient'), 'landscape')
```

### Хелперы (определяются внутри каждой функции)
- `_sp(p)` — убрать отступы до/после параграфа
- `_run(p, text, pt, bold)` — добавить run с TNR
- `_para(doc, text, align, pt, bold)` — добавить параграф
- `_line_br(p)` — `<w:br/>` внутри параграфа
- `_tcpr(cell)` — получить/создать `w:tcPr`
- `_borders(cell, sz=3)` — рамка single на все стороны
- `_bottom_border(cell, sz=6)` — только нижняя черта
- `_no_border(cell)` — убрать все рамки
- `_cw(cell, w)` — ширина ячейки в dxa
- `_valign(cell, val='center')` — вертикальное выравнивание ячейки
- `_fill(cell, text, align, pt, bold)` — заполнить ячейку текстом
- `_set_grid(tbl)` — установить `w:tblGrid` по массиву GRID
- `_tbl_w(tbl)` — установить общую ширину таблицы
- `_row_h(row, cm)` — точная высота строки (hRule=exact)
- `_row_h_min(row, cm)` — минимальная высота строки (hRule=atLeast)

### btLr вертикальный текст (Очки Многоборье в Слаломе)
```python
def _vert_cell(cell):
    pr = _tcpr(cell)
    td = OxmlElement('w:textDirection')
    td.set(qn('w:val'), 'btLr')
    pr.append(td)
    _valign(cell, 'bottom')          # vAlign=bottom → текст визуально внизу
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER   # jc=center → по центру по горизонтали
```

### vMerge (объединение строк)
```python
def _vmerge_restart(cell):   # первая ячейка объединения
    pr = _tcpr(cell)
    vm = OxmlElement('w:vMerge')
    vm.set(qn('w:val'), 'restart')
    pr.append(vm)

def _vmerge_cont(cell):      # продолжение объединения
    pr = _tcpr(cell)
    vm = OxmlElement('w:vMerge')
    pr.append(vm)
```

### Секционные строки (заголовок группы)
```python
r = tbl.add_row(); _row_h(r, 0.6)
cs = list(r.cells)
cs[0].merge(cs[NC - 1]); _cw(cs[0], TW)
_borders(cs[0])
_fill(cs[0], 'Спринт – R4, Мужчины', align=C, pt=9, bold=True)
```

### Подписи (без рамок)
```python
for pos, fio in (('Главный судья', сор.главный_судья or ''),
                 ('Главный секретарь', сор.главный_секретарь or '')):
    r = tbl.add_row(); _row_h(r, 1.2)
    cs = list(r.cells)
    cs[0].merge(cs[4]); _cw(cs[0], W_SIG_L)
    cs[5].merge(cs[9]); _cw(cs[5], W_SIG_R)
    _no_border(cs[0]); _no_border(cs[5])
    _fill(cs[0], pos, pt=10)
    _fill(cs[5], fio, pt=10)
```

## Что предстоит сделать
1. Папки `образцы_*`/`образы_*` с исходными образцами документов удалены — формат уже перенесён в соответствующие функции `create_*_docx` в `docx_utils.py`. Остались нереализованными: серверный Word/Excel-экспорт отчётов (`routes/отчеты.py`, пока только HTML для печати) и сетка параллельного спринта (`routes/пс.py` — маршрут `протокол_сетка`, всё ещё на `send_docx(html)`)
2. PDF-выгрузка (пока только браузерная печать — нет серверной генерации PDF)

## Команды git
```bash
git add -p          # стейджинг
git commit -m "..."
git push -u origin main
git pull --rebase origin main  # при конфликте
```
