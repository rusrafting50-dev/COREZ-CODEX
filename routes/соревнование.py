# routes/соревнование.py — CRUD соревнований
from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from models import db, Соревнование

bp_соревнование = Blueprint('соревнование', __name__)


@bp_соревнование.route('/')
def список():
    соревнования = Соревнование.query.order_by(Соревнование.создано.desc()).all()
    return render_template('соревнование/список.html', соревнования=соревнования)


@bp_соревнование.route('/новое', methods=['GET', 'POST'])
def новое():
    if request.method == 'POST':
        сор = Соревнование(
            организаторы                = request.form.get('организаторы', '').strip(),
            название                    = request.form['название'].strip(),
            даты_проведения             = request.form.get('даты_проведения', '').strip(),
            место_проведения            = request.form.get('место_проведения', '').strip(),
            код_субъекта_рф             = _int_or_none(request.form.get('код_субъекта_рф')),
            класс_судна                 = request.form.get('класс_судна', 'R4'),
            возрастная_группа           = request.form.get('возрастная_группа', '').strip(),
            статус                      = request.form.get('статус', '').strip(),
            уровень                     = request.form.get('уровень', '').strip(),
            код_соревнований            = request.form.get('код_соревнований', '').strip(),
            главный_судья               = request.form.get('главный_судья', '').strip(),
            главный_секретарь           = request.form.get('главный_секретарь', '').strip(),
            председатель_комиссии_допуска = request.form.get('председатель_комиссии_допуска', '').strip(),
            начальник_дистанции         = request.form.get('начальник_дистанции', '').strip(),
            начальник_спасательной_службы = request.form.get('начальник_спасательной_службы', '').strip(),
            зам_главного_секретаря      = request.form.get('зам_главного_секретаря', '').strip(),
            кол_мужских_экипажей        = _int_or_none(request.form.get('кол_мужских_экипажей')),
            кол_женских_экипажей        = _int_or_none(request.form.get('кол_женских_экипажей')),
            кол_попыток_слалом          = int(request.form.get('кол_попыток_слалом', 2)),
            кол_спортсменов_в_экипаже   = int(request.form.get('кол_спортсменов_в_экипаже', 4)),
            кол_субъектов               = _int_or_none(request.form.get('кол_субъектов')),
        )
        db.session.add(сор)
        db.session.commit()
        flash('Соревнование создано.', 'success')
        return redirect(url_for('соревнование.карточка', id=сор.id))
    return render_template('соревнование/форма.html', сор=None)


@bp_соревнование.route('/<int:id>')
def карточка(id):
    сор = Соревнование.query.get_or_404(id)
    return render_template('соревнование/карточка.html', сор=сор)


@bp_соревнование.route('/<int:id>/редактировать', methods=['GET', 'POST'])
def редактировать(id):
    сор = Соревнование.query.get_or_404(id)
    if request.method == 'POST':
        сор.организаторы                = request.form.get('организаторы', '').strip()
        сор.название                    = request.form['название'].strip()
        сор.даты_проведения             = request.form.get('даты_проведения', '').strip()
        сор.место_проведения            = request.form.get('место_проведения', '').strip()
        сор.код_субъекта_рф             = _int_or_none(request.form.get('код_субъекта_рф'))
        сор.класс_судна                 = request.form.get('класс_судна', 'R4')
        сор.возрастная_группа           = request.form.get('возрастная_группа', '').strip()
        сор.статус                      = request.form.get('статус', '').strip()
        сор.уровень                     = request.form.get('уровень', '').strip()
        сор.код_соревнований            = request.form.get('код_соревнований', '').strip()
        сор.главный_судья               = request.form.get('главный_судья', '').strip()
        сор.главный_секретарь           = request.form.get('главный_секретарь', '').strip()
        сор.председатель_комиссии_допуска = request.form.get('председатель_комиссии_допуска', '').strip()
        сор.начальник_дистанции         = request.form.get('начальник_дистанции', '').strip()
        сор.начальник_спасательной_службы = request.form.get('начальник_спасательной_службы', '').strip()
        сор.зам_главного_секретаря      = request.form.get('зам_главного_секретаря', '').strip()
        сор.кол_мужских_экипажей        = _int_or_none(request.form.get('кол_мужских_экипажей'))
        сор.кол_женских_экипажей        = _int_or_none(request.form.get('кол_женских_экипажей'))
        сор.кол_попыток_слалом          = int(request.form.get('кол_попыток_слалом', 2))
        сор.кол_спортсменов_в_экипаже   = int(request.form.get('кол_спортсменов_в_экипаже', 4))
        сор.кол_субъектов               = _int_or_none(request.form.get('кол_субъектов'))
        db.session.commit()
        flash('Данные сохранены.', 'success')
        return redirect(url_for('соревнование.карточка', id=сор.id))
    return render_template('соревнование/форма.html', сор=сор)


@bp_соревнование.route('/<int:id>/удалить', methods=['POST'])
def удалить(id):
    сор = Соревнование.query.get_or_404(id)
    db.session.delete(сор)
    db.session.commit()
    flash('Соревнование удалено.', 'warning')
    return redirect(url_for('main.index'))


def _int_or_none(val):
    try:
        return int(val) if val and str(val).strip() else None
    except (ValueError, TypeError):
        return None
