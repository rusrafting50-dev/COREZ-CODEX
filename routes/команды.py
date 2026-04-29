# routes/команды.py — CRUD экипажей (команд) + комиссия по допуску
from datetime import date
from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify
from models import db, Команда, Субъект, Соревнование, Спортсмен, ПредварительнаяЗаявка

bp_команды = Blueprint('команды', __name__)

ОРГИ = ['ДЮСШ', 'СДЮСШОР', 'СШОР', 'СШ', 'ЦСП', 'УОР', 'Другая']
ОПЛАТА_СПИСОК = ['Ожидает оплаты', 'Оплачено']
РАЗРЯДЫ_ДОПУСК = ['МСМК', 'МС', 'КМС', 'I', 'II', 'III']
ВОЗРАСТНЫЕ_ГРУППЫ = [
    'Мужчины', 'Женщины',
    'Юниоры (до 26 лет)', 'Юниорки (до 26 лет)',
    'Юниоры (до 24 лет)', 'Юниорки (до 24 лет)',
    'Юниоры (до 20 лет)', 'Юниорки (до 20 лет)',
    'Юноши (до 16 лет)', 'Девушки (до 16 лет)',
]


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
    return redirect(url_for('команды.новая_заявка', sor_id=sor_id))


@bp_команды.route('/<int:sor_id>/новая', methods=['GET'])
def новая_заявка(sor_id):
    Соревнование.query.get_or_404(sor_id)
    команда = Команда(
        соревнование_id=sor_id,
        нагрудный_номер=_следующий_нагрудный(sor_id),
    )
    db.session.add(команда)
    db.session.commit()
    return redirect(url_for('команды.допуск_команда', sor_id=sor_id, id=команда.id))


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


# ── Комиссия по допуску ───────────────────────────────────────────────────────

@bp_команды.route('/<int:sor_id>/допуск')
def допуск(sor_id):
    сор     = Соревнование.query.get_or_404(sor_id)
    команды = (Команда.query
               .filter_by(соревнование_id=sor_id)
               .order_by(Команда.нагрудный_номер)
               .all())
    return render_template('команды/допуск.html', сор=сор, команды=команды)


@bp_команды.route('/<int:sor_id>/<int:id>/допуск', methods=['GET', 'POST'])
def допуск_команда(sor_id, id):
    сор      = Соревнование.query.get_or_404(sor_id)
    команда  = Команда.query.get_or_404(id)
    субъекты = (Субъект.query
                .filter_by(соревнование_id=sor_id)
                .order_by(Субъект.номер_делегации)
                .all())

    if request.method == 'POST':
        action = request.form.get('action', '')

        if action == 'save_team':
            команда.субъект_id           = _int_or_none(request.form.get('субъект_id'))
            команда.название             = request.form.get('название', '').strip()
            команда.принадлежность_к_орг = request.form.get('принадлежность_к_орг', '').strip()
            команда.оплата_взноса        = request.form.get('оплата_взноса', 'Ожидает оплаты')
            команда.класс_судна          = request.form.get('класс_судна', 'R4')
            пол_возраст = request.form.get('пол_возраст', '')
            команда.возрастная_группа    = пол_возраст
            ж_keywords = ('Женщины', 'Юниорки', 'Девушки')
            команда.пол = 'ж' if any(пол_возраст.startswith(k) for k in ж_keywords) else 'м'
            тренер = request.form.get('тренер_представитель', '').strip()
            if команда.субъект and тренер:
                команда.субъект.представитель = тренер
            db.session.commit()
            flash('Данные команды сохранены.', 'success')
            return redirect(url_for('команды.допуск_команда', sor_id=sor_id, id=id))

        elif action == 'save_athletes':
            # Батчевое сохранение: удаляем всех и создаём заново
            for с in list(команда.спортсмены):
                db.session.delete(с)
            db.session.flush()
            фио_список    = request.form.getlist('фио[]')
            пол_список    = request.form.getlist('пол[]')
            дата_список   = request.form.getlist('дата_рождения[]')
            разряд_список = request.form.getlist('разряд[]')
            группа_список = request.form.getlist('возрастная_группа[]')
            приказ_список = request.form.getlist('номер_приказа[]')
            полис_список  = request.form.getlist('номер_страховки[]')
            дсп_список    = request.form.getlist('дист_спринт[]')
            дпс_список    = request.form.getlist('дист_пс[]')
            дсл_список    = request.form.getlist('дист_слалом[]')
            ддг_список    = request.form.getlist('дист_дг[]')
            n = len(фио_список)
            def _гет(lst, i): return lst[i].strip() if i < len(lst) else ''
            for i in range(n):
                фио = фио_список[i].strip()
                if not фио:
                    continue
                с = Спортсмен(
                    команда_id           = id,
                    субъект_id           = команда.субъект_id,
                    номер_в_экипаже      = i + 1,
                    фамилия_имя_отчество = фио,
                    пол                  = _гет(пол_список, i) or 'м',
                    дата_рождения        = _parse_date(_гет(дата_список, i)),
                    разряд               = _гет(разряд_список, i),
                    возрастная_группа    = _гет(группа_список, i),
                    номер_приказа        = _гет(приказ_список, i),
                    номер_страховки      = _гет(полис_список, i),
                    принадлежность_к_орг = команда.принадлежность_к_орг or '',
                    тренер_представитель = (команда.субъект.представитель if команда.субъект else ''),
                    участие_спринт       = _гет(дсп_список, i) == '1',
                    участие_пс           = _гет(дпс_список, i) == '1',
                    участие_слалом       = _гет(дсл_список, i) == '1',
                    участие_дг           = _гет(ддг_список, i) == '1',
                )
                db.session.add(с)
            db.session.commit()
            flash('Состав спортсменов сохранён.', 'success')
            return redirect(url_for('команды.допуск_команда', sor_id=sor_id, id=id))

        elif action == 'del_athlete':
            sp_id = _int_or_none(request.form.get('спортсмен_id'))
            if sp_id:
                с = Спортсмен.query.get(sp_id)
                if с and с.команда_id == id:
                    db.session.delete(с)
                    db.session.commit()
                    flash('Спортсмен удалён.', 'warning')
            return redirect(url_for('команды.допуск_команда', sor_id=sor_id, id=id))

    спортсмены = sorted(команда.спортсмены,
                        key=lambda с: (с.номер_в_экипаже or 9999, с.id))
    return render_template('команды/допуск_команда.html',
                           сор=сор, команда=команда, субъекты=субъекты,
                           спортсмены=спортсмены,
                           орги=ОРГИ, оплата_список=ОПЛАТА_СПИСОК,
                           разряды=РАЗРЯДЫ_ДОПУСК,
                           возрастные_группы=ВОЗРАСТНЫЕ_ГРУППЫ)


@bp_команды.route('/<int:sor_id>/подсказки-спортсменов')
def подсказки_спортсменов(sor_id):
    """AJAX: спортсмены из предварительных заявок данного соревнования."""
    команда_id = _int_or_none(request.args.get('команда_id'))
    команда    = Команда.query.get(команда_id) if команда_id else None

    заявки = ПредварительнаяЗаявка.query.filter_by(соревнование_id=sor_id).all()

    if команда:
        субъект_назв  = (команда.субъект.название or '').lower() if команда.субъект else ''
        назв_команды  = (команда.название or '').lower()
        подходящие = [
            з for з in заявки
            if (субъект_назв and з.субъект_рф
                and субъект_назв in з.субъект_рф.lower())
            or (назв_команды and з.наименование_команды
                and назв_команды in з.наименование_команды.lower())
        ]
        заявки = подходящие if подходящие else заявки

    результат = []
    seen: set = set()
    for з in заявки:
        for сп in з.спортсмены:
            части = [сп.фамилия, сп.имя, сп.отчество]
            фио = ' '.join(p for p in части if p).strip()
            if фио and фио not in seen:
                seen.add(фио)
                результат.append({
                    'фио':              фио,
                    'пол':              сп.пол or '',
                    'дата_рождения':    сп.дата_рождения.isoformat() if сп.дата_рождения else '',
                    'квалификация':     сп.квалификация or '',
                    'возрастная_группа': сп.возрастная_группа or '',
                })

    return jsonify(sorted(результат, key=lambda x: x['фио']))


# ── Вспомогательные функции ───────────────────────────────────────────────────

def _следующий_нагрудный(sor_id):
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


def _parse_date(val):
    if not val or not str(val).strip():
        return None
    try:
        return date.fromisoformat(val.strip())
    except ValueError:
        return None
