# routes/main.py — главная страница
from datetime import date as _date
from flask import Blueprint, render_template
from models import Соревнование

bp_main = Blueprint('main', __name__)


@bp_main.route('/')
def index():
    все = Соревнование.query.all()
    # Ранние → поздние; без дата_начала — в конец (date.max как ключ-заглушка)
    все.sort(key=lambda с: с.дата_начала if с.дата_начала is not None else _date.max)
    return render_template('index.html', соревнования=все)
