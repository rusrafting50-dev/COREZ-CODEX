# routes/субъекты.py — CRUD субъектов РФ (делегаций)
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Субъект, Соревнование

bp_субъекты = Blueprint('субъекты', __name__)


@bp_субъекты.route('/<int:sor_id>')
def список(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    субъекты = (Субъект.query
                .filter_by(соревнование_id=sor_id)
                .order_by(Субъект.номер_делегации)
                .all())
    return render_template('субъекты/список.html', сор=сор, субъекты=субъекты)


@bp_субъекты.route('/<int:sor_id>/добавить', methods=['GET', 'POST'])
def добавить(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    if request.method == 'POST':
        субъект = Субъект(
            соревнование_id  = sor_id,
            номер_делегации  = _int_or_none(request.form.get('номер_делегации')),
            название         = request.form['название'].strip(),
            код_субъекта_рф  = _int_or_none(request.form.get('код_субъекта_рф')),
        )
        db.session.add(субъект)
        db.session.commit()
        flash('Субъект добавлен.', 'success')
        return redirect(url_for('субъекты.список', sor_id=sor_id))
    return render_template('субъекты/форма.html', сор=сор, субъект=None)


@bp_субъекты.route('/<int:sor_id>/<int:id>/редактировать', methods=['GET', 'POST'])
def редактировать(sor_id, id):
    сор     = Соревнование.query.get_or_404(sor_id)
    субъект = Субъект.query.get_or_404(id)
    if request.method == 'POST':
        субъект.номер_делегации = _int_or_none(request.form.get('номер_делегации'))
        субъект.название        = request.form['название'].strip()
        субъект.код_субъекта_рф = _int_or_none(request.form.get('код_субъекта_рф'))
        db.session.commit()
        flash('Данные сохранены.', 'success')
        return redirect(url_for('субъекты.список', sor_id=sor_id))
    return render_template('субъекты/форма.html', сор=сор, субъект=субъект)


@bp_субъекты.route('/<int:sor_id>/<int:id>/удалить', methods=['POST'])
def удалить(sor_id, id):
    субъект = Субъект.query.get_or_404(id)
    db.session.delete(субъект)
    db.session.commit()
    flash('Субъект удалён.', 'warning')
    return redirect(url_for('субъекты.список', sor_id=sor_id))


def _int_or_none(val):
    try:
        return int(val) if val and str(val).strip() else None
    except (ValueError, TypeError):
        return None
