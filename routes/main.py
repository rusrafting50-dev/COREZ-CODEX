# routes/main.py — главная страница
from datetime import date as _date
from flask import Blueprint, render_template, redirect, url_for, request
from models import db, Соревнование, получить_настройки

bp_main = Blueprint('main', __name__)


@bp_main.route('/')
def index():
    все = Соревнование.query.all()

    год     = request.args.get('год', type=int)
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
