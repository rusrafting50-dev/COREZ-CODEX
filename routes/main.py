# routes/main.py — главная страница
from flask import Blueprint, render_template
from models import Соревнование

bp_main = Blueprint('main', __name__)


@bp_main.route('/')
def index():
    соревнования = Соревнование.query.order_by(Соревнование.создано.desc()).all()
    return render_template('index.html', соревнования=соревнования)
