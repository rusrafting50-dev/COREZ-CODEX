# sync_sports_referee.py — поиск судей в базе SPORTS_REFEREE
#
# Используется только для автодополнения на форме «Добавить судью»
# (routes/судьи.py, маршрут /поиск) — помимо локального поиска по судьям,
# уже когда-либо добавленным в коллегии COREZ, подмешиваются живые
# результаты поиска по полной базе судей SPORTS_REFEREE (её GET
# /api/judges/search — см. SPORTS_REFEREE routes/api.py). Формат ответа
# там специально сделан идентичным локальному поиску COREZ (ключи фио,
# квалификация, субъект_рф, пол, дата_рождения, дата_присвоения,
# дата_подтверждения, конец_аттестации, номер_в_базе), поэтому результат
# можно подмешивать в тот же список без преобразований.
#
# Недоступность SPORTS_REFEREE не должна ломать локальный поиск COREZ —
# при любой ошибке просто возвращается пустой список.
import os

import requests


def _адрес():
    return os.environ.get('SPORTS_REFEREE_RAFTING_URL',
                          os.environ.get('SPORTS_REFEREE_URL', 'http://127.0.0.1:5005')).rstrip('/')


def _ключ():
    return os.environ.get('SPORTS_REFEREE_RAFTING_API_KEY',
                          os.environ.get('SPORTS_REFEREE_API_KEY', ''))


def найти_судей(q):
    """Возвращает список судей из SPORTS_REFEREE, подходящих под запрос q
    (формат — как у локального /судьи/поиск). Пустой список при любой
    ошибке (ключ не задан, сайт недоступен, неожиданный ответ)."""
    ключ = _ключ()
    if not ключ:
        return []

    адрес = _адрес()
    try:
        ответ = requests.get(
            f'{адрес}/api/judges/search',
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


def _полный_статус(сор):
    """«Чемпионат субъекта Российской Федерации» и т. п. — по коду из
    справочника СТАТУСЫ, иначе из статуса и уровня соревнования."""
    from routes.соревнование import СТАТУСЫ
    по_коду = СТАТУСЫ.get(сор.код_соревнований or '')
    if по_коду:
        return по_коду['полное']
    return ' '.join(x for x in (сор.статус, сор.уровень) if x)


def отправить_судей(сор, судьи):
    """Отправляет судей из судейской коллегии соревнования в SPORTS_REFEREE_RAFTING
    (POST /api/competition/import). Возвращает (успех: bool, сообщение: str)."""
    ключ = _ключ()
    if not ключ:
        return False, 'SPORTS_REFEREE_API_KEY не настроен — отправка невозможна.'

    список = []
    пропущено = 0
    for с in судьи:
        фио = (с.фамилия_имя_отчество or '').strip()
        if not фио:
            пропущено += 1
            continue
        список.append({
            'full_name':     фио,
            'qualification': с.квалификация or '',
            'birth_date':    с.дата_рождения.isoformat() if с.дата_рождения else '',
            'region':        с.субъект_рф or '',
            'position':      с.должность or '',
            'score':         с.оценка or 'Отлично',
        })

    if not список:
        return False, 'В судейской коллегии нет судей для отправки.'

    # Данные соревнования — для строки «Практика судейства» в карточке каждого судьи
    payload = {
        'competition_name':     сор.название or '',
        'competition_id':       сор.id,
        'competition_dates':    сор.даты_проведения or '',
        'competition_start':    сор.дата_начала.isoformat() if сор.дата_начала else '',
        'competition_location': сор.место_проведения or '',
        'competition_ekp':      сор.номер_екп or '',
        'competition_status':   _полный_статус(сор),
        'judges': список,
    }

    try:
        ответ = requests.post(
            f'{_адрес()}/api/competition/import',
            json=payload,
            headers={'X-API-Key': ключ},
            timeout=10,
        )
    except requests.RequestException as e:
        return False, f'Ошибка соединения с SPORTS_REFEREE_RAFTING: {e}'

    if ответ.status_code != 200:
        return False, f'SPORTS_REFEREE_RAFTING вернул HTTP {ответ.status_code}.'

    try:
        р = ответ.json()
    except ValueError:
        return False, 'Неожиданный ответ от SPORTS_REFEREE_RAFTING.'

    совпало = р.get('matched', 0)
    создано = р.get('created', 0)
    сообщение = (
        f'SPORTS_REFEREE_RAFTING: совпало {совпало} чел., добавлено новых {создано} чел.'
    )
    if 'practice_added' in р:
        сообщение += (
            f' Практика судейства: добавлено {р.get("practice_added", 0)}, '
            f'обновлено {р.get("practice_updated", 0)}.'
        )
    if пропущено:
        сообщение += f' Пропущено (нет ФИО): {пропущено}.'
    return True, сообщение
