# routes/заявка.py — предварительные заявки на соревнования
from datetime import date
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Соревнование, ПредварительнаяЗаявка, СпортсменЗаявки

bp_заявка = Blueprint('заявка', __name__)

ВОЗРАСТНЫЕ_ГРУППЫ = [
    'Мужчины', 'Женщины',
    'Юниоры (до 26 лет)', 'Юниорки (до 26 лет)',
    'Юниоры (до 24 лет)', 'Юниорки (до 24 лет)',
    'Юниоры (до 20 лет)', 'Юниорки (до 20 лет)',
    'Юноши (до 16 лет)', 'Девушки (до 16 лет)',
]


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
            орг_наименование           = request.form.get('орг_наименование', '').strip(),
            орг_фио_руководителя       = request.form.get('орг_фио_руководителя', '').strip(),
            орг_должность_руководителя = request.form.get('орг_должность_руководителя', '').strip(),
            орг_телефон                = request.form.get('орг_телефон', '').strip(),
            орг_email                  = request.form.get('орг_email', '').strip(),
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
    return render_template('заявка/форма.html', сор=сор, з=None,
                           возрастные_группы=ВОЗРАСТНЫЕ_ГРУППЫ)


@bp_заявка.route('/<int:sor_id>/<int:z_id>/редактировать', methods=['GET', 'POST'])
def редактировать(sor_id, z_id):
    сор = Соревнование.query.get_or_404(sor_id)
    з = ПредварительнаяЗаявка.query.get_or_404(z_id)
    if request.method == 'POST':
        з.субъект_рф           = request.form.get('субъект_рф', '').strip()
        з.наименование_команды = request.form.get('наименование_команды', '').strip()
        з.орг_наименование           = request.form.get('орг_наименование', '').strip()
        з.орг_фио_руководителя       = request.form.get('орг_фио_руководителя', '').strip()
        з.орг_должность_руководителя = request.form.get('орг_должность_руководителя', '').strip()
        з.орг_телефон                = request.form.get('орг_телефон', '').strip()
        з.орг_email                  = request.form.get('орг_email', '').strip()
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
    return render_template('заявка/форма.html', сор=сор, з=з,
                           возрастные_группы=ВОЗРАСТНЫЕ_ГРУППЫ)


@bp_заявка.route('/<int:sor_id>/<int:z_id>/удалить', methods=['POST'])
def удалить(sor_id, z_id):
    з = ПредварительнаяЗаявка.query.get_or_404(z_id)
    db.session.delete(з)
    db.session.commit()
    flash('Заявка удалена.', 'warning')
    return redirect(url_for('заявка.список', sor_id=sor_id))


@bp_заявка.route('/<int:sor_id>/<int:z_id>/печать')
def печать(sor_id, z_id):
    сор = Соревнование.query.get_or_404(sor_id)
    з = ПредварительнаяЗаявка.query.get_or_404(z_id)
    return render_template('заявка/печать.html', сор=сор, з=з,
                           сегодня=date.today())


def _сохранить_спортсменов(з, form):
    фио_list  = form.getlist('фио[]')
    полы      = form.getlist('пол[]')
    даты      = form.getlist('дата_рождения[]')
    квалиф    = form.getlist('квалификация[]')
    приказы   = form.getlist('номер_приказа[]')
    страховки = form.getlist('номер_страховки[]')
    группы    = form.getlist('возрастная_группа[]')
    for i, фио in enumerate(фио_list):
        фио = фио.strip()
        if not фио:
            continue
        части = фио.split(None, 2)
        фам  = части[0] if len(части) > 0 else ''
        имя  = части[1] if len(части) > 1 else ''
        отч  = части[2] if len(части) > 2 else ''
        др = None
        try:
            from datetime import datetime as dt
            др = dt.strptime(даты[i], '%Y-%m-%d').date() if даты[i] else None
        except (ValueError, IndexError):
            pass
        сп = СпортсменЗаявки(
            заявка_id         = з.id,
            фамилия           = фам,
            имя               = имя,
            отчество          = отч,
            пол               = полы[i]              if i < len(полы)      else '',
            дата_рождения     = др,
            квалификация      = квалиф[i].strip()    if i < len(квалиф)    else '',
            номер_приказа     = приказы[i].strip()   if i < len(приказы)   else '',
            номер_страховки   = страховки[i].strip() if i < len(страховки) else '',
            возрастная_группа = группы[i].strip()    if i < len(группы)    else '',
        )
        db.session.add(сп)
