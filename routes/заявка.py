# routes/заявка.py — предварительные заявки на соревнования
from datetime import date
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Соревнование, ПредварительнаяЗаявка, СпортсменЗаявки

bp_заявка = Blueprint('заявка', __name__)


@bp_заявка.route('/<int:sor_id>')
def список(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    заявки = (ПредварительнаяЗаявка.query
              .filter_by(соревнование_id=sor_id)
              .order_by(ПредварительнаяЗаявка.создано)
              .all())
    return render_template('заявка/список.html', сор=сор, заявки=заявки)


@bp_заявка.route('/<int:sor_id>/новая', methods=['GET', 'POST'])
def новая(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    if request.method == 'POST':
        з = ПредварительнаяЗаявка(
            соревнование_id      = sor_id,
            субъект_рф           = request.form.get('субъект_рф', '').strip(),
            наименование_команды = request.form.get('наименование_команды', '').strip(),
            орг_наименование     = request.form.get('орг_наименование', '').strip(),
            орг_фио_руководителя = request.form.get('орг_фио_руководителя', '').strip(),
            орг_телефон          = request.form.get('орг_телефон', '').strip(),
            орг_email            = request.form.get('орг_email', '').strip(),
            тренер_фио           = request.form.get('тренер_фио', '').strip(),
            тренер_телефон       = request.form.get('тренер_телефон', '').strip(),
            тренер_email         = request.form.get('тренер_email', '').strip(),
        )
        db.session.add(з)
        db.session.flush()
        _сохранить_спортсменов(з, request.form)
        db.session.commit()
        flash('Заявка сохранена.', 'success')
        return redirect(url_for('заявка.список', sor_id=sor_id))
    return render_template('заявка/форма.html', сор=сор, з=None)


@bp_заявка.route('/<int:sor_id>/<int:зid>/редактировать', methods=['GET', 'POST'])
def редактировать(sor_id, зid):
    сор = Соревнование.query.get_or_404(sor_id)
    з = ПредварительнаяЗаявка.query.get_or_404(зid)
    if request.method == 'POST':
        з.субъект_рф           = request.form.get('субъект_рф', '').strip()
        з.наименование_команды = request.form.get('наименование_команды', '').strip()
        з.орг_наименование     = request.form.get('орг_наименование', '').strip()
        з.орг_фио_руководителя = request.form.get('орг_фио_руководителя', '').strip()
        з.орг_телефон          = request.form.get('орг_телефон', '').strip()
        з.орг_email            = request.form.get('орг_email', '').strip()
        з.тренер_фио           = request.form.get('тренер_фио', '').strip()
        з.тренер_телефон       = request.form.get('тренер_телефон', '').strip()
        з.тренер_email         = request.form.get('тренер_email', '').strip()
        for сп in з.спортсмены:
            db.session.delete(сп)
        db.session.flush()
        _сохранить_спортсменов(з, request.form)
        db.session.commit()
        flash('Заявка обновлена.', 'success')
        return redirect(url_for('заявка.список', sor_id=sor_id))
    return render_template('заявка/форма.html', сор=сор, з=з)


@bp_заявка.route('/<int:sor_id>/<int:зid>/удалить', methods=['POST'])
def удалить(sor_id, зid):
    з = ПредварительнаяЗаявка.query.get_or_404(зid)
    db.session.delete(з)
    db.session.commit()
    flash('Заявка удалена.', 'warning')
    return redirect(url_for('заявка.список', sor_id=sor_id))


@bp_заявка.route('/<int:sor_id>/<int:зid>/печать')
def печать(sor_id, зid):
    сор = Соревнование.query.get_or_404(sor_id)
    з = ПредварительнаяЗаявка.query.get_or_404(зid)
    return render_template('заявка/печать.html', сор=сор, з=з,
                           сегодня=date.today())


def _сохранить_спортсменов(з, form):
    фамилии  = form.getlist('фамилия[]')
    имена    = form.getlist('имя[]')
    отчества = form.getlist('отчество[]')
    полы     = form.getlist('пол[]')
    даты     = form.getlist('дата_рождения[]')
    квалиф   = form.getlist('квалификация[]')
    приказы  = form.getlist('номер_приказа[]')
    страховки = form.getlist('номер_страховки[]')
    for i, фам in enumerate(фамилии):
        if not фам.strip():
            continue
        др = None
        try:
            from datetime import datetime as dt
            др = dt.strptime(даты[i], '%Y-%m-%d').date() if даты[i] else None
        except (ValueError, IndexError):
            pass
        сп = СпортсменЗаявки(
            заявка_id       = з.id,
            фамилия         = фам.strip(),
            имя             = имена[i].strip()    if i < len(имена)    else '',
            отчество        = отчества[i].strip() if i < len(отчества) else '',
            пол             = полы[i]             if i < len(полы)     else '',
            дата_рождения   = др,
            квалификация    = квалиф[i].strip()   if i < len(квалиф)   else '',
            номер_приказа   = приказы[i].strip()  if i < len(приказы)  else '',
            номер_страховки = страховки[i].strip() if i < len(страховки) else '',
        )
        db.session.add(сп)
