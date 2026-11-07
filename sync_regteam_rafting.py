# sync_regteam_rafting.py — поиск спортсменов в базе REGTEAM_RAFTING
#
# Используется только для автодополнения на форме предварительной заявки
# на соревнование (routes/заявка.py, поле «Поиск спортсмена в базе»,
# шаблон заявка/форма.html вызывает /спортсмены/поиск) — помимо
# локального поиска по спортсменам, уже когда-либо участвовавшим в
# соревнованиях COREZ, подмешиваются живые результаты поиска по базе
# спортсменов региональной сборной REGTEAM_RAFTING (её GET
# /api/athletes/search). Формат ответа там специально сделан идентичным
# локальному поиску COREZ (ключи фио, дата_рождения, пол, разряд,
# номер_в_базе), поэтому результат можно подмешивать в тот же список без
# преобразований.
#
# Недоступность REGTEAM_RAFTING не должна ломать локальный поиск COREZ —
# при любой ошибке просто возвращается пустой список.
import os

import requests


def _адрес():
    return os.environ.get('REGTEAM_RAFTING_URL', 'http://127.0.0.1:5001').rstrip('/')


def _ключ():
    return os.environ.get('REGTEAM_RAFTING_API_KEY', '')


def найти_спортсменов(q):
    """Возвращает список спортсменов из REGTEAM_RAFTING, подходящих под
    запрос q (формат — как у локального /спортсмены/поиск). Пустой список
    при любой ошибке (ключ не задан, сайт недоступен, неожиданный ответ)."""
    ключ = _ключ()
    if not ключ:
        return []

    адрес = _адрес()
    try:
        ответ = requests.get(
            f'{адрес}/api/athletes/search',
            params={'q': q},
            headers={'X-API-Key': ключ},
            timeout=5,
        )
    except requests.RequestException:
        return []

    if ответ.status_code != 200:
        return []

    try:
        данные = ответ.json()
    except ValueError:
        return []

    return данные if isinstance(данные, list) else []


def отправить_заявку(сор, з):
    """Отправляет спортсменов из предварительной заявки в REGTEAM_RAFTING
    (POST /api/competition/import). Возвращает (успех: bool, сообщение: str)."""
    ключ = _ключ()
    if not ключ:
        return False, 'REGTEAM_RAFTING_API_KEY не настроен — отправка невозможна.'

    спортсмены = []
    пропущено_имена = []
    for сп in з.спортсмены:
        if not (сп.фамилия or сп.имя):
            пропущено_имена.append(
                ' '.join(filter(None, [сп.фамилия, сп.имя, сп.отчество])) or '(пусто)'
            )
            continue
        спортсмены.append({
            'last_name':   сп.фамилия  or '',
            'first_name':  сп.имя      or '',
            'middle_name': сп.отчество or '',
            'birth_date':  сп.дата_рождения.isoformat() if сп.дата_рождения else '',
            'gender':      сп.пол          or '',
            'rank':        сп.квалификация or '',
            'subject':     з.субъект_рф    or '',
        })

    if not спортсмены:
        return False, 'В заявке нет спортсменов для отправки.'

    payload = {
        'competition_name': сор.название or '',
        'competition_date': '',
        'athletes': спортсмены,
    }

    try:
        ответ = requests.post(
            f'{_адрес()}/api/competition/import',
            json=payload,
            headers={'X-API-Key': ключ},
            timeout=10,
        )
    except requests.RequestException as e:
        return False, f'Ошибка соединения с REGTEAM_RAFTING: {e}'

    if ответ.status_code != 200:
        return False, f'REGTEAM_RAFTING вернул HTTP {ответ.status_code}.'

    try:
        р = ответ.json()
    except ValueError:
        return False, 'Неожиданный ответ от REGTEAM_RAFTING.'

    совпало   = р.get('matched', 0)
    создано   = р.get('created', 0)
    добавлено = р.get('participations_created', 0)
    конфликты = р.get('conflicts', 0)
    сообщение = (
        f'REGTEAM_RAFTING: совпало {совпало} чел., '
        f'добавлено новых {создано} чел., '
        f'записей участия создано {добавлено}.'
    )
    if конфликты:
        сообщение += f' ⚠ {конфликты} чел. числятся в сборной другого региона — проверьте страницу «Из COREZ».'
    if пропущено_имена:
        сообщение += f' Пропущено (нет фамилии и имени): {len(пропущено_имена)} — {", ".join(пропущено_имена[:5])}.'
    return True, сообщение
