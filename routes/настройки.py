# routes/настройки.py — раздел «Настройки»: справочники для фильтров на главной
# (субъекты ЦФО, муниципальные образования) и импорт соревнований из Excel.
# Защищён простым паролем (SETTINGS_PASSWORD в .env) через сессию Flask.
import os
from datetime import date as _date, datetime as _datetime
from functools import wraps

from flask import (Blueprint, render_template, request, redirect, url_for,
                    flash, session)

from models import db, Соревнование, СубъектЦФО, МуниципальноеОбразование

bp_настройки = Blueprint('настройки', __name__)


def требуется_вход(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get('настройки_вход'):
            return redirect(url_for('настройки.вход', next=request.path))
        return view(*args, **kwargs)
    return wrapped


@bp_настройки.route('/вход', methods=['GET', 'POST'])
def вход():
    if session.get('настройки_вход'):
        return redirect(url_for('настройки.страница'))
    if request.method == 'POST':
        пароль_верный = os.environ.get('SETTINGS_PASSWORD', '')
        введённый = request.form.get('пароль', '')
        if not пароль_верный:
            flash('Пароль для Настроек не задан в .env (SETTINGS_PASSWORD).', 'danger')
        elif введённый == пароль_верный:
            session['настройки_вход'] = True
            return redirect(request.args.get('next') or url_for('настройки.страница'))
        else:
            flash('Неверный пароль.', 'danger')
    return render_template('настройки/вход.html')


@bp_настройки.route('/выход', methods=['POST'])
def выход():
    session.pop('настройки_вход', None)
    return redirect(url_for('main.index'))


@bp_настройки.route('/')
@требуется_вход
def страница():
    субъекты = СубъектЦФО.query.order_by(СубъектЦФО.порядок, СубъектЦФО.название).all()
    мо = МуниципальноеОбразование.query.order_by(МуниципальноеОбразование.порядок,
                                                 МуниципальноеОбразование.название).all()
    return render_template('настройки/страница.html', субъекты=субъекты, мо=мо)


@bp_настройки.route('/субъект/добавить', methods=['POST'])
@требуется_вход
def субъект_добавить():
    название = request.form.get('название', '').strip()
    if название:
        if not СубъектЦФО.query.filter_by(название=название).first():
            db.session.add(СубъектЦФО(название=название))
            db.session.commit()
            flash('Субъект добавлен.', 'success')
        else:
            flash('Такой субъект уже есть в списке.', 'warning')
    return redirect(url_for('настройки.страница'))


@bp_настройки.route('/субъект/<int:id>/удалить', methods=['POST'])
@требуется_вход
def субъект_удалить(id):
    субъект = СубъектЦФО.query.get_or_404(id)
    db.session.delete(субъект)
    db.session.commit()
    flash('Субъект удалён.', 'warning')
    return redirect(url_for('настройки.страница'))


@bp_настройки.route('/мо/добавить', methods=['POST'])
@требуется_вход
def мо_добавить():
    название = request.form.get('название', '').strip()
    if название:
        if not МуниципальноеОбразование.query.filter_by(название=название).first():
            db.session.add(МуниципальноеОбразование(название=название))
            db.session.commit()
            flash('Муниципальное образование добавлено.', 'success')
        else:
            flash('Такое муниципальное образование уже есть в списке.', 'warning')
    return redirect(url_for('настройки.страница'))


@bp_настройки.route('/мо/<int:id>/удалить', methods=['POST'])
@требуется_вход
def мо_удалить(id):
    мо = МуниципальноеОбразование.query.get_or_404(id)
    db.session.delete(мо)
    db.session.commit()
    flash('Муниципальное образование удалено.', 'warning')
    return redirect(url_for('настройки.страница'))


# ── Импорт из Excel (лист «КП» — Календарный план) ─────────────────────────

ВИД_СПОРТА_ФИЛЬТР = 'рафтинг'
RU_МЕСЯЦЫ = {
    1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля', 5: 'мая', 6: 'июня',
    7: 'июля', 8: 'августа', 9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря',
}


def _нормализовать_пробелы(текст):
    """Схлопнуть переносы строк и повторяющиеся пробелы в один пробел."""
    import re
    return re.sub(r'\s+', ' ', str(текст)).strip()


def _даты_проведения_текст(начало, окончание):
    """Собрать текст вида «11 — 12 апреля 2026 года» из дат начала/окончания."""
    if not начало:
        return ''
    if not окончание or окончание == начало:
        return f'{начало.day} {RU_МЕСЯЦЫ[начало.month]} {начало.year} года'
    if начало.month == окончание.month and начало.year == окончание.year:
        return (f'{начало.day} — {окончание.day} {RU_МЕСЯЦЫ[начало.month]} '
                f'{начало.year} года')
    return (f'{начало.day} {RU_МЕСЯЦЫ[начало.month]} — '
            f'{окончание.day} {RU_МЕСЯЦЫ[окончание.month]} {окончание.year} года')


def _разобрать_кп(путь):
    """Разобрать лист «КП» Календарного плана. Вернуть список черновиков-словарей.

    Строки ищутся по заголовку («Наименование мероприятия» во 2-м столбце),
    затем читаются до конца листа; в черновики попадают только строки с
    видом спорта «Рафтинг».
    """
    import openpyxl
    wb = openpyxl.load_workbook(путь, data_only=True)
    лист = None
    for имя in wb.sheetnames:
        if имя.strip().upper() == 'КП':
            лист = wb[имя]
            break
    if лист is None:
        лист = wb[wb.sheetnames[0]]

    заголовок_ряд = None
    for i, строка in enumerate(лист.iter_rows(min_row=1, max_row=лист.max_row, values_only=True), start=1):
        if len(строка) > 1 and строка[1] and 'наименование мероприятия' in str(строка[1]).strip().lower():
            заголовок_ряд = i
            break
    if заголовок_ряд is None:
        return []

    черновики = []
    for строка in лист.iter_rows(min_row=заголовок_ряд + 1, max_row=лист.max_row, values_only=True):
        if len(строка) < 8:
            continue
        название    = строка[1]
        вид_спорта  = строка[2]
        начало      = строка[3]
        окончание   = строка[4]
        место       = строка[6]
        участников  = строка[7]
        if not название or not str(название).strip():
            continue
        if not вид_спорта or ВИД_СПОРТА_ФИЛЬТР not in str(вид_спорта).strip().lower():
            continue

        def _к_дате(зн):
            if isinstance(зн, _datetime):
                return зн.date()
            if isinstance(зн, _date):
                return зн
            return None

        д_начало = _к_дате(начало)
        д_окончание = _к_дате(окончание)
        try:
            кол_участников = int(участников) if участников not in (None, '') else None
        except (TypeError, ValueError):
            кол_участников = None

        черновики.append({
            'название': _нормализовать_пробелы(название),
            'дата_начала': д_начало,
            'даты_проведения': _даты_проведения_текст(д_начало, д_окончание),
            'место_проведения': _нормализовать_пробелы(место) if место else '',
            'кол_участников': кол_участников,
        })
    return черновики


@bp_настройки.route('/импорт', methods=['GET', 'POST'])
@требуется_вход
def импорт():
    if request.method == 'POST':
        файл = request.files.get('файл')
        if not файл or not файл.filename:
            flash('Выберите файл .xlsx для импорта.', 'warning')
            return redirect(url_for('настройки.импорт'))
        if not файл.filename.lower().endswith(('.xlsx', '.xlsm')):
            flash('Поддерживаются только файлы .xlsx/.xlsm.', 'warning')
            return redirect(url_for('настройки.импорт'))

        import tempfile
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            файл.save(tmp.name)
            путь = tmp.name
        try:
            черновики = _разобрать_кп(путь)
        except Exception as exc:
            flash(f'Не удалось разобрать файл: {exc}', 'danger')
            return redirect(url_for('настройки.импорт'))
        finally:
            os.unlink(путь)

        if not черновики:
            flash('В файле не найдено строк с видом спорта «Рафтинг» на листе «КП».', 'warning')
            return redirect(url_for('настройки.импорт'))

        создано = 0
        for ч in черновики:
            сор = Соревнование(
                название          = ч['название'],
                дата_начала       = ч['дата_начала'],
                даты_проведения   = ч['даты_проведения'],
                место_проведения  = ч['место_проведения'],
                кол_участников    = ч['кол_участников'],
                класс_судна       = 'R4',
            )
            db.session.add(сор)
            создано += 1
        db.session.commit()
        flash(f'Импортировано черновиков соревнований: {создано}. '
              'Заполните остальные поля (статус, классы судов, возрастные группы и т.д.) '
              'на странице редактирования каждого соревнования.', 'success')
        return redirect(url_for('main.index'))

    return render_template('настройки/импорт.html')
