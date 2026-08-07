"""
Статические нормативы рафтинга.
Источник: таблица разрядов ЕВСК (Нормативы рафтинг.xlsx).
"""

# (код_соревнований, возрастная_категория) → {'М': [7], 'С': [6], 'П': [7]}
# None в списке = нет норматива для этого места.
_ДАННЫЕ = {
    ('ЧР', 'мж'):    {'М': ['МС','МС','МС','КМС','КМС','КМС',None],   'С': ['КМС','КМС',None,None,None,None],              'П': ['КМС','КМС','КМС',None,None,None,None]},
    ('КР', 'мж'):    {'М': ['МС','МС','КМС','КМС','КМС',None,None],   'С': ['КМС',None,None,None,None,None],               'П': ['КМС','КМС',None,None,None,None,None]},
    ('ПР', '24'):    {'М': ['КМС','КМС','КМС','I','I',None,None],     'С': ['I','I','II','II','II',None],                  'П': ['КМС','I','I',None,None,None,None]},
    ('ПР', '20'):    {'М': ['КМС','КМС','I','I','I',None,None],       'С': ['I','I','II','II','II',None],                  'П': ['КМС','I','I',None,None,None,None]},
    ('ПР', '16'):    {'М': ['КМС','I','I',None,None,None,None],       'С': ['I','II','II','II',None,None],                 'П': ['I','I','II','II','II',None,None]},
    ('ВС', 'мж'):    {'М': ['МС','КМС','КМС',None,None,None,None],    'С': ['I','I','II','II',None,None],                  'П': ['КМС','I','I',None,None,None,None]},
    ('ВС', '26'):    {'М': ['КМС','КМС','I','I',None,None,None],      'С': ['I','II','II',None,None,None],                 'П': ['КМС','I','II','II',None,None,None]},
    ('ВС', '24'):    {'М': ['КМС','КМС','I','I',None,None,None],      'С': ['I','II','II',None,None,None],                 'П': ['I','I','II','II',None,None,None]},
    ('ВС', '20'):    {'М': ['КМС','КМС','I','I',None,None,None],      'С': ['I','II','II',None,None,None],                 'П': ['I','I','II','II',None,None,None]},
    ('ВС', '16'):    {'М': ['КМС','I','I',None,None,None,None],       'С': ['II','II','III','III','III',None],             'П': ['I','II','II',None,None,None,None]},
    ('ЧФО', 'мж'):   {'М': ['КМС','I','I',None,None,None,None],       'С': ['I','II','II','III','III','III'],               'П': ['I','I','II','II','III','III','III']},
    ('ПФО', '24'):   {'М': ['I','I','I','II','II','II',None],         'С': ['II','III','III',None,None,None],              'П': ['II','II','III','III',None,None,None]},
    ('ПФО', '20'):   {'М': ['I','I','II','II','II','III','III'],       'С': ['II','III','III',None,None,None],              'П': ['II','II','III','III',None,None,None]},
    ('ПФО', '16'):   {'М': ['II','II','III','III','Iю','Iю',None],    'С': ['III','III','Iю','Iю',None,None],              'П': ['II','III','III','Iю','Iю',None,None]},
    ('МСС', '24'):   {'М': ['I','II','III',None,None,None,None],      'С': ['II','II','III',None,None,None],               'П': ['II','II','III',None,None,None,None]},
    ('МСС', '20'):   {'М': ['I','II','III',None,None,None,None],      'С': ['II','III','III','Iю','Iю',None],              'П': ['II','III','III','Iю','Iю',None,None]},
    ('МСС', '16'):   {'М': ['II','III','III','Iю','Iю',None,None],    'С': ['III','III','Iю','Iю',None,None],              'П': ['III','III','Iю','Iю',None,None,None]},
    ('ЧС', 'мж'):    {'М': ['I','I','I','II','II','III','III'],        'С': ['II','III','III',None,None,None],              'П': ['II','II','III','III',None,None,None]},
    ('КС', 'мж'):    {'М': ['I','I','II','II','III','III',None],       'С': ['II','III','III',None,None,None],              'П': ['II','II','III','III',None,None,None]},
    ('ПС', '24'):    {'М': ['II','II','II','III','III',None,None],     'С': ['II','III','III','III',None,None],             'П': ['II','III','III','III',None,None,None]},
    ('ПС', '20'):    {'М': ['II','II','II','III','III',None,None],     'С': ['II','III','III','III',None,None],             'П': ['II','III','III','III',None,None,None]},
    ('ПС', '16'):    {'М': ['III','III','Iю','Iю','Iю','IIю','IIю'],  'С': ['III','Iю','Iю','Iю','IIю','IIю'],            'П': ['III','Iю','Iю','Iю','IIю','IIю',None]},
    ('СС', 'мж'):    {'М': ['II','II','II','III','III',None,None],     'С': ['II','II','III','III',None,None],              'П': ['II','II','III','III',None,None,None]},
    ('СС', '24'):    {'М': ['II','III','III',None,None,None,None],     'С': ['III','III','Iю','Iю',None,None],             'П': ['III','III','Iю','Iю',None,None,None]},
    ('СС', '20'):    {'М': ['II','III','III',None,None,None,None],     'С': ['III','III','Iю','Iю',None,None],             'П': ['III','III','Iю','Iю',None,None,None]},
    ('СС', '16'):    {'М': ['III','III','Iю','Iю','Iю','IIю','IIю'],  'С': ['III','Iю','Iю','IIю','IIю',None],            'П': ['III','Iю','Iю','IIю','IIю',None,None]},
    ('ЧМО', 'мж'):   {'М': ['II','II','III','III',None,None,None],     'С': ['II','III','III',None,None,None],              'П': ['II','III','III',None,None,None,None]},
    ('ПМО', '24'):   {'М': ['II','II','II','III','III',None,None],     'С': ['III','III','III','III',None,None],            'П': ['III','III','III','III',None,None,None]},
    ('ПМО', '20'):   {'М': ['II','II','III','III',None,None,None],     'С': ['III','III','III','III',None,None],            'П': ['III','III','III','III',None,None,None]},
    ('ПМО', '16'):   {'М': ['III','Iю','Iю','Iю','IIю','IIю','IIю'], 'С': ['Iю','Iю','IIю','IIю','IIIю','IIIю'],         'П': ['Iю','Iю','IIю','IIю','IIIю','IIIю',None]},
    ('ССМО', 'мж'):  {'М': ['II','II','II','III','III',None,None],     'С': ['III','III',None,None,None,None],              'П': ['III','III',None,None,None,None,None]},
    ('ССМО', '24'):  {'М': ['II','II','II','III','III',None,None],     'С': ['III','III','III','III',None,None],            'П': ['III','III','III','III',None,None,None]},
    ('ССМО', '20'):  {'М': ['II','II','III','III',None,None,None],     'С': ['III','III','III','III',None,None],            'П': ['III','III','III','III',None,None,None]},
    ('ССМО', '16'):  {'М': ['III','Iю','Iю','Iю','IIю','IIю',None],  'С': ['Iю','Iю','IIю','IIю','IIIю','IIIю'],         'П': ['Iю','Iю','IIю','IIю','IIIю','IIIю',None]},
}

_ВОЗРАСТ_КАРТА = {
    'Мужчины, женщины':              'мж',
    'юниоры, юниорки (до 26 лет)':   '26',
    'юниоры, юниорки (до 24 лет)':   '24',
    'юниоры, юниорки (до 20 лет)':   '20',
    'юноши, девушки (до 16 лет)':    '16',
    'Мужчины':                       'мж',
    'Женщины':                       'мж',
    'Юниоры (до 26 лет)':            '26',
    'Юниорки (до 26 лет)':           '26',
    'Юниоры (до 24 лет)':            '24',
    'Юниорки (до 24 лет)':           '24',
    'Юниоры (до 20 лет)':            '20',
    'Юниорки (до 20 лет)':           '20',
    'Юноши (до 16 лет)':             '16',
    'Девушки (до 16 лет)':           '16',
    'Мужчины, Женщины':              'мж',
    'Юниоры, Юниорки (до 26 лет)':   '26',
    'Юниоры, Юниорки (до 24 лет)':   '24',
    'Юниоры, Юниорки (до 20 лет)':   '20',
    'Юноши, Девушки (до 16 лет)':    '16',
}


def _возр_категория(возрастная_группа):
    if not возрастная_группа:
        return 'мж'
    for часть in возрастная_группа.split(' / '):
        кат = _ВОЗРАСТ_КАРТА.get(часть.strip())
        if кат:
            return кат
    return 'мж'


def получить_норматив(код, возрастная_группа, тип, место):
    """Вернуть спортивный норматив или '—'.

    тип: 'М' (Многоборье), 'С' (Спринт), 'П' (ПС / Слалом / ДГ)
    место: место с 1
    """
    if not код or место is None:
        return '—'
    кат = _возр_категория(возрастная_группа)
    запись = _ДАННЫЕ.get((код, кат))
    if запись is None:
        return '—'
    список = запись.get(тип, [])
    idx = место - 1
    if idx < 0 or idx >= len(список):
        return '—'
    норм = список[idx]
    return норм if норм is not None else '—'
