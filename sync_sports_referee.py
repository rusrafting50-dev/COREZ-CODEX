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
    return os.environ.get('SPORTS_REFEREE_URL', 'http://127.0.0.1:5002').rstrip('/')


def _ключ():
    return os.environ.get('SPORTS_REFEREE_API_KEY', '')


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
