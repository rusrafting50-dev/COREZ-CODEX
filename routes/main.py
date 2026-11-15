# routes/main.py — главная страница
from datetime import date as _date
from flask import Blueprint, render_template, redirect, url_for, request, send_file
from models import db, Соревнование, СубъектЦФО, получить_настройки
import календарь_excel

bp_main = Blueprint('main', __name__)


@bp_main.route('/')
def index():
    все = Соревнование.query.all()

    # Главная — календарь одного года: по умолчанию текущий
    год     = request.args.get('год', type=int) or _date.today().year
    субъект = (request.args.get('субъект') or '').strip()
    мо      = (request.args.get('мо') or '').strip()

    if год:
        все = [с for с in все if с.дата_начала and с.дата_начала.year == год]
    if субъект:
        все = [с for с in все if (с.субъект_рф_название or '').strip() == субъект]
    if мо:
        все = [с for с in все if (с.муниципальное_образование or '').strip() == мо]

    # Поздние → ранние; без дата_начала — в конец
    все.sort(key=lambda с: с.дата_начала if с.дата_начала is not None else _date.min, reverse=True)
    return render_template('index.html', соревнования=все,
                           фильтр_год=год, фильтр_субъект=субъект, фильтр_мо=мо)


@bp_main.route('/экспорт-excel')
def экспорт_excel():
    """Выгрузка соревнований в Excel по форме «Предложения по включению
    спортивных мероприятий в Календарный план…»: фильтры — год, субъект РФ
    или «Всероссийские в ЦФО», с муниципальными или без; ниже — предпросмотр."""
    все = Соревнование.query.all()
    субъекты_цфо = [с.название for с in СубъектЦФО.query.order_by(СубъектЦФО.порядок, СубъектЦФО.название).all()]
    if not субъекты_цфо:  # справочник в «Настройках» не заполнен — субъекты из самих соревнований
        субъекты_цфо = sorted({(с.субъект_рф_название or '').strip() for с in все} - {''})
    годы = sorted({с.дата_начала.year for с in все if с.дата_начала}, reverse=True) or [_date.today().year]

    год = request.args.get('год', type=int) or (_date.today().year if _date.today().year in годы else годы[0])
    субъект = (request.args.get('субъект') or '').strip() or (субъекты_цфо[0] if субъекты_цфо else календарь_excel.ВСЕРОССИЙСКИЕ_В_ЦФО)
    с_муниципальными = request.args.get('муниципальные') == '1'

    выбранные = календарь_excel.отобрать(все, год, субъект, с_муниципальными, set(субъекты_цфо))

    if request.args.get('скачать') == '1':
        настр = получить_настройки()
        орг = СубъектЦФО.query.filter_by(название=субъект).first()
        файл = календарь_excel.сформировать(
            выбранные, год, субъект, (настр.вид_спорта or 'Рафтинг').strip() or 'Рафтинг',
            орг.организация_название if орг and орг.организация_название else '',
        )
        return send_file(
            файл, as_attachment=True, download_name=f'Предложения в КП {субъект} {год}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    return render_template(
        'экспорт_excel.html', годы=годы, год=год, субъекты=субъекты_цфо, субъект=субъект,
        всероссийские=календарь_excel.ВСЕРОССИЙСКИЕ_В_ЦФО, с_муниципальными=с_муниципальными,
        соревнования=выбранные, даты=календарь_excel.даты, муниципальное=календарь_excel.муниципальное,
    )


@bp_main.route('/архив')
def архив():
    текущий = _date.today().year
    исключить = {текущий - 1, текущий, текущий + 1}
    годы = sorted(
        {с.дата_начала.year for с in Соревнование.query.all() if с.дата_начала},
        reverse=True,
    )
    годы = [г for г in годы if г not in исключить]
    return render_template('архив.html', годы=годы)


@bp_main.route('/переключить-офлайн', methods=['POST'])
def переключить_офлайн():
    настр = получить_настройки()
    настр.офлайн_режим = not настр.офлайн_режим
    db.session.commit()
    назад = request.form.get('назад') or url_for('main.index')
    return redirect(назад)
