# routes/дистанция.py — параметры дистанций и акт сдачи-приёмки
from flask import Blueprint, render_template, request, redirect, url_for, flash
from models import db, Соревнование, ПараметрыДистанции

bp_дистанция = Blueprint('дистанция', __name__)

MAX_ДИСТАНЦИЙ = 4


@bp_дистанция.route('/<int:sor_id>/дистанция')
def страница(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    дистанции = сор.параметры_дистанций
    if not дистанции:
        дистанции = [ПараметрыДистанции(соревнование_id=sor_id, порядок=1)]
    return render_template('дистанция/дистанция.html',
                           сор=сор, дистанции=дистанции,
                           max_дистанций=MAX_ДИСТАНЦИЙ)


@bp_дистанция.route('/<int:sor_id>/дистанция/сохранить', methods=['POST'])
def сохранить(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    кол = min(int(request.form.get('кол_дистанций', 1)), MAX_ДИСТАНЦИЙ)

    # Удалить старые записи и создать заново
    for д in list(сор.параметры_дистанций):
        db.session.delete(д)
    db.session.flush()

    for i in range(кол):
        название = request.form.get(f'название_{i}', '').strip()
        д = ПараметрыДистанции(
            соревнование_id=sor_id,
            порядок=i + 1,
            название=название or None,
            протяженность=request.form.get(f'протяженность_{i}', '').strip() or None,
            категория=request.form.get(f'категория_{i}', '').strip() or None,
            среднее_время=request.form.get(f'среднее_время_{i}', '').strip() or None,
            лимит_времени=request.form.get(f'лимит_времени_{i}', '').strip() or None,
            вариант_старта=request.form.get(f'вариант_старта_{i}', '').strip() or None,
            вариант_финиша=request.form.get(f'вариант_финиша_{i}', '').strip() or None,
            страховка=request.form.get(f'страховка_{i}', '').strip() or None,
        )
        db.session.add(д)

    db.session.commit()
    flash('Параметры дистанций сохранены.', 'success')
    return redirect(url_for('дистанция.страница', sor_id=sor_id))


@bp_дистанция.route('/<int:sor_id>/дистанция/акт')
def акт(sor_id):
    сор = Соревнование.query.get_or_404(sor_id)
    дистанции = сор.параметры_дистанций
    from datetime import date
    return render_template('дистанция/акт_дистанции.html',
                           сор=сор, дистанции=дистанции, сегодня=date.today())
