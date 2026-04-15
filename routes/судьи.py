# routes/судьи.py — CRUD судей
from datetime import date
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Судья, Соревнование

bp_судьи = Blueprint('судьи', __name__)

КВАЛИФИКАЦИИ  = ['ССВК', 'СС1К', 'СС2К', 'СС3К', 'ЮС']
ДОЛЖНОСТИ = [
    'Главный судья',
    'Главный секретарь',
    'Зам. главного секретаря',
    'Председатель комиссии по допуску',
    'Начальник дистанции',
    'Начальник спасательной службы',
    'Судья на старте',
    'Судья на финише',
    'Судья у ворот',
    'Судья-хронометрист',
    'Технический делегат',
]


@bp_судьи.route('/<int:sor_id>')
def список(sor_id):
    сор   = Соревнование.query.get_or_404(sor_id)
    судьи = (Судья.query
             .filter_by(соревнование_id=sor_id)
             .order_by(Судья.фамилия_имя_отчество)
             .all())
    return render_template('судьи/список.html', сор=сор, судьи=судьи)


@bp_судьи.route('/<int:sor_id>/добавить', methods=['GET', 'POST'])
def добавить(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    if request.method == 'POST':
        судья = Судья(
            номер_в_базе          = _int_or_none(request.form.get('номер_в_базе')),
            соревнование_id       = sor_id,
            фамилия_имя_отчество  = request.form['фамилия_имя_отчество'].strip(),
            квалификация          = request.form.get('квалификация', '').strip(),
            субъект_рф            = request.form.get('субъект_рф', '').strip(),
            код_субъекта_рф       = _int_or_none(request.form.get('код_субъекта_рф')),
            пол                   = request.form.get('пол', '').strip(),
            дата_рождения         = _parse_date(request.form.get('дата_рождения')),
            дата_присвоения       = _parse_date(request.form.get('дата_присвоения')),
            дата_подтверждения    = _parse_date(request.form.get('дата_подтверждения')),
            конец_аттестации      = _parse_date(request.form.get('конец_аттестации')),
            действующая_категория = request.form.get('действующая_категория') == '1',
            должность             = request.form.get('должность', '').strip(),
        )
        db.session.add(судья)
        db.session.commit()
        flash('Судья добавлен.', 'success')
        return redirect(url_for('судьи.список', sor_id=sor_id))
    return render_template('судьи/форма.html',
                           сор=сор, судья=None,
                           квалификации=КВАЛИФИКАЦИИ, должности=ДОЛЖНОСТИ)


@bp_судьи.route('/<int:sor_id>/<int:id>/редактировать', methods=['GET', 'POST'])
def редактировать(sor_id, id):
    сор   = Соревнование.query.get_or_404(sor_id)
    судья = Судья.query.get_or_404(id)
    if request.method == 'POST':
        судья.номер_в_базе          = _int_or_none(request.form.get('номер_в_базе'))
        судья.фамилия_имя_отчество  = request.form['фамилия_имя_отчество'].strip()
        судья.квалификация          = request.form.get('квалификация', '').strip()
        судья.субъект_рф            = request.form.get('субъект_рф', '').strip()
        судья.код_субъекта_рф       = _int_or_none(request.form.get('код_субъекта_рф'))
        судья.пол                   = request.form.get('пол', '').strip()
        судья.дата_рождения         = _parse_date(request.form.get('дата_рождения'))
        судья.дата_присвоения       = _parse_date(request.form.get('дата_присвоения'))
        судья.дата_подтверждения    = _parse_date(request.form.get('дата_подтверждения'))
        судья.конец_аттестации      = _parse_date(request.form.get('конец_аттестации'))
        судья.действующая_категория = request.form.get('действующая_категория') == '1'
        судья.должность             = request.form.get('должность', '').strip()
        db.session.commit()
        flash('Данные сохранены.', 'success')
        return redirect(url_for('судьи.список', sor_id=sor_id))
    return render_template('судьи/форма.html',
                           сор=сор, судья=судья,
                           квалификации=КВАЛИФИКАЦИИ, должности=ДОЛЖНОСТИ)


@bp_судьи.route('/<int:sor_id>/<int:id>/удалить', methods=['POST'])
def удалить(sor_id, id):
    судья = Судья.query.get_or_404(id)
    db.session.delete(судья)
    db.session.commit()
    flash('Судья удалён.', 'warning')
    return redirect(url_for('судьи.список', sor_id=sor_id))


def _int_or_none(val):
    try:
        return int(val) if val and str(val).strip() else None
    except (ValueError, TypeError):
        return None


def _parse_date(val):
    if not val or not str(val).strip():
        return None
    try:
        return date.fromisoformat(val.strip())
    except ValueError:
        return None
