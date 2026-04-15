# routes/команды.py — CRUD экипажей (команд)
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Команда, Субъект, Соревнование

bp_команды = Blueprint('команды', __name__)


@bp_команды.route('/<int:sor_id>')
def список(sor_id):
    сор     = Соревнование.query.get_or_404(sor_id)
    команды = (Команда.query
               .filter_by(соревнование_id=sor_id)
               .order_by(Команда.нагрудный_номер)
               .all())
    return render_template('команды/список.html', сор=сор, команды=команды)


@bp_команды.route('/<int:sor_id>/добавить', methods=['GET', 'POST'])
def добавить(sor_id):
    сор      = Соревнование.query.get_or_404(sor_id)
    субъекты = Субъект.query.filter_by(соревнование_id=sor_id).order_by(Субъект.номер_делегации).all()
    if request.method == 'POST':
        команда = Команда(
            соревнование_id     = sor_id,
            субъект_id          = _int_or_none(request.form.get('субъект_id')),
            нагрудный_номер     = int(request.form['нагрудный_номер']),
            название            = request.form.get('название', '').strip(),
            класс_судна         = request.form.get('класс_судна', 'R4'),
            пол                 = request.form.get('пол', 'м'),
            участие_спринт      = bool(request.form.get('участие_спринт')),
            участие_пс          = bool(request.form.get('участие_пс')),
            участие_слалом      = bool(request.form.get('участие_слалом')),
            участие_дг          = bool(request.form.get('участие_дг')),
            участие_многоборье  = bool(request.form.get('участие_многоборье')),
        )
        db.session.add(команда)
        db.session.commit()
        flash('Команда добавлена.', 'success')
        return redirect(url_for('команды.список', sor_id=sor_id))
    следующий_номер = _следующий_нагрудный(sor_id)
    return render_template('команды/форма.html', сор=сор, команда=None,
                           субъекты=субъекты,
                           следующий_номер=следующий_номер)


@bp_команды.route('/<int:sor_id>/<int:id>/редактировать', methods=['GET', 'POST'])
def редактировать(sor_id, id):
    сор      = Соревнование.query.get_or_404(sor_id)
    команда  = Команда.query.get_or_404(id)
    субъекты = Субъект.query.filter_by(соревнование_id=sor_id).order_by(Субъект.номер_делегации).all()
    if request.method == 'POST':
        команда.субъект_id          = _int_or_none(request.form.get('субъект_id'))
        команда.нагрудный_номер     = int(request.form['нагрудный_номер'])
        команда.название            = request.form.get('название', '').strip()
        команда.класс_судна         = request.form.get('класс_судна', 'R4')
        команда.пол                 = request.form.get('пол', 'м')
        команда.участие_спринт      = bool(request.form.get('участие_спринт'))
        команда.участие_пс          = bool(request.form.get('участие_пс'))
        команда.участие_слалом      = bool(request.form.get('участие_слалом'))
        команда.участие_дг          = bool(request.form.get('участие_дг'))
        команда.участие_многоборье  = bool(request.form.get('участие_многоборье'))
        db.session.commit()
        flash('Данные сохранены.', 'success')
        return redirect(url_for('команды.список', sor_id=sor_id))
    return render_template('команды/форма.html', сор=сор, команда=команда,
                           субъекты=субъекты, следующий_номер=None)


@bp_команды.route('/<int:sor_id>/<int:id>/удалить', methods=['POST'])
def удалить(sor_id, id):
    команда = Команда.query.get_or_404(id)
    db.session.delete(команда)
    db.session.commit()
    flash('Команда удалена.', 'warning')
    return redirect(url_for('команды.список', sor_id=sor_id))


def _следующий_нагрудный(sor_id):
    """Возвращает следующий свободный нагрудный номер для соревнования."""
    последний = (Команда.query
                 .filter_by(соревнование_id=sor_id)
                 .order_by(Команда.нагрудный_номер.desc())
                 .first())
    return (последний.нагрудный_номер + 1) if последний else 1


def _int_or_none(val):
    try:
        return int(val) if val and str(val).strip() else None
    except (ValueError, TypeError):
        return None
