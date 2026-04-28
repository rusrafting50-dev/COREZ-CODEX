# routes/main.py — главная страница
from flask import Blueprint, render_template
from models import Соревнование

bp_main = Blueprint('main', __name__)


@bp_main.route('/')
def index():
    все = Соревнование.query.all()
    # Сортировка по дате начала (ранние вперёд), соревнования без даты — в конец
    все.sort(key=lambda с: (с.дата_начала is None, с.дата_начала))
    return render_template('index.html', соревнования=все)
