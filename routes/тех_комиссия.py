# routes/тех_комиссия.py — Техническая комиссия
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Соревнование, ТехПроверка, ТехСнаряжение, ТехСоставКомиссии, Судья, СУБЪЕКТЫ_РФ

bp_тех = Blueprint('тех', __name__)

НАИМЕНОВАНИЯ  = ['', 'Рафт R4', 'Рафт R6', 'Спасжилет', 'Каска', 'Весло']
ДОПУСК_СПИСОК = ['Допущен', 'Не допущен']


@bp_тех.route('/<int:sor_id>')
def список(sor_id):
    сор      = Соревнование.query.get_or_404(sor_id)
    проверки = (ТехПроверка.query
                .filter_by(соревнование_id=sor_id)
                .order_by(ТехПроверка.id)
                .all())
    return render_template('тех_комиссия/список.html', сор=сор, проверки=проверки)


@bp_тех.route('/<int:sor_id>/новая', methods=['GET', 'POST'])
def новая(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    if request.method == 'POST':
        проверка = ТехПроверка(
            соревнование_id      = sor_id,
            субъект_рф           = request.form.get('субъект_рф', '').strip(),
            наименование_команды = request.form.get('наименование_команды', '').strip(),
            тренер_фио           = request.form.get('тренер_фио', '').strip(),
        )
        db.session.add(проверка)
        db.session.flush()
        _сохранить_снаряжение(проверка, request.form)
        db.session.commit()
        flash('Проверка добавлена.', 'success')
        return redirect(url_for('тех.список', sor_id=sor_id))
    return render_template('тех_комиссия/форма.html', сор=сор, проверка=None,
                           наименования=НАИМЕНОВАНИЯ, допуск_список=ДОПУСК_СПИСОК)


@bp_тех.route('/<int:sor_id>/<int:id>/редактировать', methods=['GET', 'POST'])
def редактировать(sor_id, id):
    сор      = Соревнование.query.get_or_404(sor_id)
    проверка = ТехПроверка.query.get_or_404(id)
    if request.method == 'POST':
        проверка.субъект_рф           = request.form.get('субъект_рф', '').strip()
        проверка.наименование_команды = request.form.get('наименование_команды', '').strip()
        проверка.тренер_фио           = request.form.get('тренер_фио', '').strip()
        for с in list(проверка.снаряжение):
            db.session.delete(с)
        db.session.flush()
        _сохранить_снаряжение(проверка, request.form)
        db.session.commit()
        flash('Данные сохранены.', 'success')
        return redirect(url_for('тех.список', sor_id=sor_id))
    return render_template('тех_комиссия/форма.html', сор=сор, проверка=проверка,
                           наименования=НАИМЕНОВАНИЯ, допуск_список=ДОПУСК_СПИСОК)


@bp_тех.route('/<int:sor_id>/<int:id>/удалить', methods=['POST'])
def удалить(sor_id, id):
    проверка = ТехПроверка.query.get_or_404(id)
    db.session.delete(проверка)
    db.session.commit()
    flash('Запись удалена.', 'warning')
    return redirect(url_for('тех.список', sor_id=sor_id))


@bp_тех.route('/<int:sor_id>/состав', methods=['GET', 'POST'])
def состав(sor_id):
    сор   = Соревнование.query.get_or_404(sor_id)
    ск    = ТехСоставКомиссии.query.filter_by(соревнование_id=sor_id).first()
    судьи = (Судья.query
             .filter_by(соревнование_id=sor_id)
             .order_by(Судья.фамилия_имя_отчество)
             .all())
    if request.method == 'POST':
        if not ск:
            ск = ТехСоставКомиссии(соревнование_id=sor_id)
            db.session.add(ск)
        ск.председатель_фио          = request.form.get('председатель_фио', '').strip()
        ск.председатель_должность    = request.form.get('председатель_должность', '').strip()
        ск.председатель_квалификация = request.form.get('председатель_квалификация', '').strip()
        ск.председатель_субъект_рф   = request.form.get('председатель_субъект_рф', '').strip()
        ск.член_1_фио                = request.form.get('член_1_фио', '').strip()
        ск.член_1_должность          = request.form.get('член_1_должность', '').strip()
        ск.член_1_квалификация       = request.form.get('член_1_квалификация', '').strip()
        ск.член_1_субъект_рф         = request.form.get('член_1_субъект_рф', '').strip()
        ск.член_2_фио                = request.form.get('член_2_фио', '').strip()
        ск.член_2_должность          = request.form.get('член_2_должность', '').strip()
        ск.член_2_квалификация       = request.form.get('член_2_квалификация', '').strip()
        ск.член_2_субъект_рф         = request.form.get('член_2_субъект_рф', '').strip()
        db.session.commit()
        flash('Состав комиссии сохранён.', 'success')
        return redirect(url_for('тех.состав', sor_id=sor_id))
    return render_template('тех_комиссия/состав.html', сор=сор, ск=ск, судьи=судьи,
                           субъекты_рф=СУБЪЕКТЫ_РФ)


@bp_тех.route('/<int:sor_id>/акт')
def акт_все(sor_id):
    сор      = Соревнование.query.get_or_404(sor_id)
    проверки = (ТехПроверка.query
                .filter_by(соревнование_id=sor_id)
                .order_by(ТехПроверка.id)
                .all())
    ск = ТехСоставКомиссии.query.filter_by(соревнование_id=sor_id).first()
    return render_template('тех_комиссия/акт.html', сор=сор, проверки=проверки,
                           ск=ск, класс='')


@bp_тех.route('/<int:sor_id>/акт/<klass>')
def акт(sor_id, klass):
    сор      = Соревнование.query.get_or_404(sor_id)
    проверки = (ТехПроверка.query
                .filter_by(соревнование_id=sor_id)
                .order_by(ТехПроверка.id)
                .all())
    ск = ТехСоставКомиссии.query.filter_by(соревнование_id=sor_id).first()
    return render_template('тех_комиссия/акт.html', сор=сор, проверки=проверки,
                           ск=ск, класс=klass.upper())


@bp_тех.route('/<int:sor_id>/акт-word')
def акт_word(sor_id):
    from docx_utils import create_акт_тх_docx
    сор      = Соревнование.query.get_or_404(sor_id)
    проверки = (ТехПроверка.query
                .filter_by(соревнование_id=sor_id)
                .order_by(ТехПроверка.id)
                .all())
    ск = ТехСоставКомиссии.query.filter_by(соревнование_id=sor_id).first()
    safe = (сор.название or str(sor_id))[:50].replace('/', '_')
    return create_акт_тх_docx(сор, проверки, ск, f'Акт ТК — {safe}.docx')


def _сохранить_снаряжение(проверка, form):
    наим_список  = form.getlist('наименование[]')
    марка_список = form.getlist('марка[]')
    серн_список  = form.getlist('серийный_номер[]')
    кол_список   = form.getlist('количество[]')
    доп_список   = form.getlist('допуск[]')

    def _гет(lst, i):
        return lst[i].strip() if i < len(lst) else ''

    for i, наим in enumerate(наим_список):
        if not наим.strip():
            continue
        кол_str = _гет(кол_список, i)
        db.session.add(ТехСнаряжение(
            проверка_id    = проверка.id,
            наименование   = наим.strip(),
            марка          = _гет(марка_список, i),
            серийный_номер = _гет(серн_список, i),
            количество     = int(кол_str) if кол_str.isdigit() else 1,
            допуск         = _гет(доп_список, i) or 'Допущен',
            порядок        = i,
        ))
