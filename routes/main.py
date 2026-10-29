# routes/main.py — главная страница
from datetime import date as _date
from flask import Blueprint, render_template, redirect, url_for, request
from models import db, Соревнование, получить_настройки

bp_main = Blueprint('main', __name__)


@bp_main.route('/')
def index():
    все = Соревнование.query.all()
    # Поздние → ранние; без дата_начала — в конец
    все.sort(key=lambda с: с.дата_начала if с.дата_начала is not None else _date.min, reverse=True)
    return render_template('index.html', соревнования=все)


@bp_main.route('/переключить-офлайн', methods=['POST'])
def переключить_офлайн():
    настр = получить_настройки()
    настр.офлайн_режим = not настр.офлайн_режим
    db.session.commit()
    назад = request.form.get('назад') or url_for('main.index')
    return redirect(назад)
