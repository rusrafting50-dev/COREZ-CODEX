# routes/спортсмены.py — CRUD спортсменов
from datetime import date
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Спортсмен, Команда, Субъект, Соревнование

bp_спортсмены = Blueprint('спортсмены', __name__)

РАЗРЯДЫ = ['МС', 'МСМК', 'КМС', 'I', 'II', 'III', 'Iю', 'IIю', 'IIIю', 'б/р']


@bp_спортсмены.route('/<int:sor_id>')
def список(sor_id):
    сор        = Соревнование.query.get_or_404(sor_id)
    спортсмены = (Спортсмен.query
                  .join(Команда)
                  .filter(Команда.соревнование_id == sor_id)
                  .order_by(Команда.нагрудный_номер, Спортсмен.фамилия_имя_отчество)
                  .all())
    return render_template('спортсмены/список.html', сор=сор, спортсмены=спортсмены)


@bp_спортсмены.route('/<int:sor_id>/добавить', methods=['GET', 'POST'])
def добавить(sor_id):
    сор     = Соревнование.query.get_or_404(sor_id)
    команды = (Команда.query
               .filter_by(соревнование_id=sor_id)
               .order_by(Команда.нагрудный_номер)
               .all())
    субъекты = (Субъект.query
                .filter_by(соревнование_id=sor_id)
                .order_by(Субъект.название)
                .all())
    if request.method == 'POST':
        спортсмен = Спортсмен(
            номер_в_базе            = _int_or_none(request.form.get('номер_в_базе')),
            команда_id              = int(request.form['команда_id']),
            субъект_id              = _int_or_none(request.form.get('субъект_id')),
            фамилия_имя_отчество    = request.form['фамилия_имя_отчество'].strip(),
            дата_рождения           = _parse_date(request.form.get('дата_рождения')),
            пол                     = request.form.get('пол', 'м'),
            разряд                  = request.form.get('разряд', '').strip(),
            возрастная_группа       = request.form.get('возрастная_группа', '').strip(),
            принадлежность_к_орг    = request.form.get('принадлежность_к_орг', '').strip(),
            тренер_представитель    = request.form.get('тренер_представитель', '').strip(),
            номер_экипажа_многоборье = _int_or_none(request.form.get('номер_экипажа_многоборье')),
            участие_спринт          = bool(request.form.get('участие_спринт')),
            участие_пс              = bool(request.form.get('участие_пс')),
            участие_слалом          = bool(request.form.get('участие_слалом')),
            участие_дг              = bool(request.form.get('участие_дг')),
        )
        db.session.add(спортсмен)
        db.session.commit()
        flash('Спортсмен добавлен.', 'success')
        return redirect(url_for('спортсмены.список', sor_id=sor_id))
    return render_template('спортсмены/форма.html',
                           сор=сор, спортсмен=None,
                           команды=команды, субъекты=субъекты, разряды=РАЗРЯДЫ)


@bp_спортсмены.route('/<int:sor_id>/<int:id>/редактировать', methods=['GET', 'POST'])
def редактировать(sor_id, id):
    сор       = Соревнование.query.get_or_404(sor_id)
    спортсмен = Спортсмен.query.get_or_404(id)
    команды   = (Команда.query
                 .filter_by(соревнование_id=sor_id)
                 .order_by(Команда.нагрудный_номер)
                 .all())
    субъекты  = (Субъект.query
                 .filter_by(соревнование_id=sor_id)
                 .order_by(Субъект.название)
                 .all())
    if request.method == 'POST':
        спортсмен.номер_в_базе             = _int_or_none(request.form.get('номер_в_базе'))
        спортсмен.команда_id               = int(request.form['команда_id'])
        спортсмен.субъект_id               = _int_or_none(request.form.get('субъект_id'))
        спортсмен.фамилия_имя_отчество     = request.form['фамилия_имя_отчество'].strip()
        спортсмен.дата_рождения            = _parse_date(request.form.get('дата_рождения'))
        спортсмен.пол                      = request.form.get('пол', 'м')
        спортсмен.разряд                   = request.form.get('разряд', '').strip()
        спортсмен.возрастная_группа        = request.form.get('возрастная_группа', '').strip()
        спортсмен.принадлежность_к_орг     = request.form.get('принадлежность_к_орг', '').strip()
        спортсмен.тренер_представитель     = request.form.get('тренер_представитель', '').strip()
        спортсмен.номер_экипажа_многоборье = _int_or_none(request.form.get('номер_экипажа_многоборье'))
        спортсмен.участие_спринт           = bool(request.form.get('участие_спринт'))
        спортсмен.участие_пс               = bool(request.form.get('участие_пс'))
        спортсмен.участие_слалом           = bool(request.form.get('участие_слалом'))
        спортсмен.участие_дг               = bool(request.form.get('участие_дг'))
        db.session.commit()
        flash('Данные сохранены.', 'success')
        return redirect(url_for('спортсмены.список', sor_id=sor_id))
    return render_template('спортсмены/форма.html',
                           сор=сор, спортсмен=спортсмен,
                           команды=команды, субъекты=субъекты, разряды=РАЗРЯДЫ)


@bp_спортсмены.route('/<int:sor_id>/<int:id>/удалить', methods=['POST'])
def удалить(sor_id, id):
    спортсмен = Спортсмен.query.get_or_404(id)
    db.session.delete(спортсмен)
    db.session.commit()
    flash('Спортсмен удалён.', 'warning')
    return redirect(url_for('спортсмены.список', sor_id=sor_id))


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
