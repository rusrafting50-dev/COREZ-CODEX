# sync_rafting_cfo.py — отправка соревнования на сайт RAFTING_CFO
#
# Сайт (rafting_cfo, отдельный Flask-проект) принимает соревнования через
# POST /api/sync/competitions с заголовком X-API-Key (см. его routes/api.py).
# Адрес сайта и ключ берутся из .env — RAFTING_CFO_URL / RAFTING_CFO_API_KEY,
# ключ должен совпадать с API_KEY в .env самого сайта.
#
# У соревнования здесь нет отдельного поля "дата окончания" — только
# дата_начала (структурированная) и даты_проведения (свободный текст вида
# "16 — 22 сентября 2025 года"). Дату окончания пытаемся вытащить из этого
# текста; не получилось — считаем соревнование однодневным.
import base64
import os
import re
from datetime import date as _date

import requests

RAFTING_CFO_URL = os.environ.get('RAFTING_CFO_URL', 'http://127.0.0.1:5003').rstrip('/')
RAFTING_CFO_API_KEY = os.environ.get('RAFTING_CFO_API_KEY', '')

ЛОГОТИП_ПАПКА = os.path.join(os.path.dirname(__file__), 'static', 'uploads', 'логотипы')

RU_MONTHS = {
    'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4, 'мая': 5, 'июня': 6,
    'июля': 7, 'августа': 8, 'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12,
}
_ДАТА_RE = re.compile(r'(\d{1,2})\s+(' + '|'.join(RU_MONTHS) + r')(?:\s+(\d{4}))?')


def _распознать_дату_окончания(текст, год_по_умолчанию):
    """Последнее совпадение вида 'DD месяц [YYYY]' в свободном тексте
    'Сроки проведения' — это и есть дата окончания при формате
    '16 — 22 сентября 2025 года'. Не найдено — None."""
    if not текст:
        return None
    совпадения = list(_ДАТА_RE.finditer(текст))
    if not совпадения:
        return None
    последнее = совпадения[-1]
    день = int(последнее.group(1))
    месяц = RU_MONTHS[последнее.group(2)]
    год = int(последнее.group(3)) if последнее.group(3) else год_по_умолчанию
    try:
        return _date(год, месяц, день)
    except ValueError:
        return None


def _определить_уровень(уровень_текст):
    """Грубое соответствие уровня COREZ (произвольный текст вида 'России',
    'субъекта Российской Федерации') уровням меню сайта (цфо/регион)."""
    текст = (уровень_текст or '').lower()
    if 'субъект' in текст or 'муницип' in текст or 'москв' in текст:
        return 'регион'
    return 'цфо'


def _определить_статус(начало, окончание):
    сегодня = _date.today()
    конец = окончание or начало
    if сегодня < начало:
        return 'анонс'
    if сегодня > конец:
        return 'завершены'
    return 'идёт'


def _логотип_base64(имя_файла):
    """Читает файл логотипа с диска и кодирует в base64 — по API логотип
    передаётся только так (это JSON-запрос, не форма с файлом)."""
    if not имя_файла:
        return None, None
    путь = os.path.join(ЛОГОТИП_ПАПКА, имя_файла)
    if not os.path.isfile(путь):
        return None, None
    with open(путь, 'rb') as f:
        данные = base64.b64encode(f.read()).decode('ascii')
    return имя_файла, данные


def собрать_данные(сор):
    """Строит словарь одного элемента для POST /api/sync/competitions."""
    окончание = _распознать_дату_окончания(сор.даты_проведения, сор.дата_начала.year)
    logo_filename, logo_base64 = _логотип_base64(сор.логотип)
    return {
        'source_id': сор.id,
        'name': сор.название,
        'comp_type': ' '.join(filter(None, [сор.статус, сор.уровень])) or None,
        'level': _определить_уровень(сор.уровень),
        'start_date': сор.дата_начала.isoformat(),
        'end_date': окончание.isoformat() if окончание else None,
        'location': сор.место_проведения,
        'status': _определить_статус(сор.дата_начала, окончание),
        'age_group': сор.возрастная_группа,
        'boat_class': сор.класс_судна,
        'head_judge': сор.главный_судья,
        'registration_number': сор.номер_екп,
        'logo_filename': logo_filename,
        'logo_base64': logo_base64,
    }


def отправить_на_сайт(сор):
    """Отправляет одно соревнование на сайт. Возвращает (успех, сообщение)."""
    if not RAFTING_CFO_API_KEY:
        return False, 'RAFTING_CFO_API_KEY не задан в .env — см. .env.example'
    if not сор.дата_начала:
        return False, 'У соревнования не заполнена дата начала — заполните на странице редактирования'

    try:
        ответ = requests.post(
            f'{RAFTING_CFO_URL}/api/sync/competitions',
            json={'items': [собрать_данные(сор)]},
            headers={'X-API-Key': RAFTING_CFO_API_KEY},
            timeout=10,
        )
    except requests.RequestException as exc:
        return False, f'Не удалось соединиться с сайтом ({RAFTING_CFO_URL}): {exc}'

    if ответ.status_code != 200:
        return False, f'Сайт ответил ошибкой {ответ.status_code}: {ответ.text[:200]}'

    данные = ответ.json()
    if данные.get('errors'):
        return False, '; '.join(данные['errors'])
    return True, 'Отправлено на сайт.'
