# models.py — модели базы данных SQLite для системы учёта соревнований по рафтингу
# Язык интерфейса: русский
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

# ---------------------------------------------------------------------------
# Глобальные настройки сервера (одна строка на всю установку)
# ---------------------------------------------------------------------------

class Настройки(db.Model):
    """Общие настройки сервера. Единственная строка, id=1."""
    __tablename__ = 'настройки'

    id            = db.Column(db.Integer, primary_key=True)
    офлайн_режим  = db.Column(db.Boolean, nullable=False, default=False)


def получить_настройки():
    """Вернуть единственную строку настроек, создав её при первом обращении."""
    настр = Настройки.query.get(1)
    if not настр:
        настр = Настройки(id=1, офлайн_режим=False)
        db.session.add(настр)
        db.session.commit()
    return настр


# ---------------------------------------------------------------------------
# Справочные таблицы (заполняются один раз из СЧИТАЛКА.xls)
# ---------------------------------------------------------------------------

class ТаблицаОчков(db.Model):
    """Очки за места по дистанциям (из листа «Очки»)."""
    __tablename__ = 'таблица_очков'

    id          = db.Column(db.Integer, primary_key=True)
    дистанция   = db.Column(db.String(50), nullable=False)   # СПРИНТ, ПС, СЛАЛОМ, ДГ, МНОГОБОРЬЕ
    место       = db.Column(db.Integer,    nullable=False)
    очки        = db.Column(db.Float,      nullable=False, default=0)

    __table_args__ = (
        db.UniqueConstraint('дистанция', 'место', name='uq_дистанция_место'),
    )

    def __repr__(self):
        return f'<ТаблицаОчков {self.дистанция} #{self.место} = {self.очки}>'


class КоэффициентыСубъектов(db.Model):
    """Коэффициенты очков субъектов в зависимости от числа субъектов (лист «Очки_субъектыР»)."""
    __tablename__ = 'коэффициенты_субъектов'

    id             = db.Column(db.Integer, primary_key=True)
    место          = db.Column(db.Integer, nullable=False)
    кол_субъектов  = db.Column(db.Integer, nullable=False)   # число участвующих субъектов (1-25)
    коэффициент    = db.Column(db.Float,   nullable=False, default=0)

    __table_args__ = (
        db.UniqueConstraint('место', 'кол_субъектов', name='uq_место_субъекты'),
    )

    def __repr__(self):
        return f'<КоэффициентыСубъектов место={self.место} субъектов={self.кол_субъектов} k={self.коэффициент}>'


class Штраф(db.Model):
    """Справочник штрафов по дистанциям (листы Штрафы_*)."""
    __tablename__ = 'штрафы'

    id          = db.Column(db.Integer, primary_key=True)
    дистанция   = db.Column(db.String(30), nullable=False)   # СТАРТ, ФИНИШ, ПС, СЛАЛОМ, ДГ
    номер       = db.Column(db.Integer,    nullable=False)
    нарушение   = db.Column(db.Text,       nullable=False)
    штраф_сек   = db.Column(db.Integer,    nullable=True)    # NULL = дисквалификация
    дисквалификация = db.Column(db.Boolean, nullable=False, default=False)
    примечание  = db.Column(db.Text,       nullable=True)

    def __repr__(self):
        s = 'ДК' if self.дисквалификация else f'{self.штраф_сек} сек'
        return f'<Штраф [{self.дистанция}] #{self.номер} {s}>'


class ТаблицаРазрядов(db.Model):
    """Разряды, присваиваемые по занятому месту и статусу соревнований (листы Разряды_*)."""
    __tablename__ = 'таблица_разрядов'

    id              = db.Column(db.Integer, primary_key=True)
    тип             = db.Column(db.String(5),  nullable=False)   # С, Д, М (спринт/длинная/многоборье)
    место           = db.Column(db.Integer,    nullable=False)
    статус_код      = db.Column(db.String(10), nullable=False)   # ЧР, КР, ВС, ЧФО …
    разряд          = db.Column(db.String(10), nullable=True)    # МС, КМС, I, II, III, Iю … или '-'

    def __repr__(self):
        return f'<ТаблицаРазрядов {self.тип} #{self.место} {self.статус_код} → {self.разряд}>'


# ---------------------------------------------------------------------------
# Основные модели соревнований
# ---------------------------------------------------------------------------

class Соревнование(db.Model):
    """Соревнование по рафтингу (один класс судов: R4 или R6)."""
    __tablename__ = 'соревнования'

    id                          = db.Column(db.Integer, primary_key=True)
    организаторы                = db.Column(db.Text,    nullable=True)
    название                    = db.Column(db.String(300), nullable=False)
    даты_проведения             = db.Column(db.String(100), nullable=True)
    дата_начала                 = db.Column(db.Date,    nullable=True)          # для сортировки
    место_проведения            = db.Column(db.String(300), nullable=True)
    код_субъекта_рф             = db.Column(db.Integer,     nullable=True)
    номер_екп                   = db.Column(db.String(100), nullable=True)
    логотип                     = db.Column(db.String(255), nullable=True)   # имя файла в static/uploads/логотипы

    # Спортивные характеристики
    класс_судна                 = db.Column(db.String(5),  nullable=False, default='R4')  # R4 / R6
    возрастная_группа           = db.Column(db.String(50), nullable=True)   # '15 лет и старше'
    статус                      = db.Column(db.String(50), nullable=True)   # Чемпионат, Кубок …
    уровень                     = db.Column(db.String(50), nullable=True)   # России, ФО, области …
    код_соревнований            = db.Column(db.String(20), nullable=True)   # ЧР, КР, ВС …

    # Статус готовности разделов (ручное переключение В работе / Готово)
    готов_организаторы  = db.Column(db.Boolean, nullable=False, default=False)
    готов_судьи         = db.Column(db.Boolean, nullable=False, default=False)
    готов_заявки        = db.Column(db.Boolean, nullable=False, default=False)
    готов_допуск        = db.Column(db.Boolean, nullable=False, default=False)
    готов_тех           = db.Column(db.Boolean, nullable=False, default=False)
    готов_наград_спринт = db.Column(db.Boolean, nullable=False, default=False)
    готов_наград_пс     = db.Column(db.Boolean, nullable=False, default=False)
    готов_наград_слалом = db.Column(db.Boolean, nullable=False, default=False)
    готов_наград_дг     = db.Column(db.Boolean, nullable=False, default=False)
    готов_наград_мб     = db.Column(db.Boolean, nullable=False, default=False)
    готов_наград_зачет  = db.Column(db.Boolean, nullable=False, default=False)
    готов_отчеты        = db.Column(db.Boolean, nullable=False, default=False)

    # Организаторы (4 строки)
    организатор_1               = db.Column(db.String(300), nullable=True)
    организатор_2               = db.Column(db.String(300), nullable=True)
    организатор_3               = db.Column(db.String(300), nullable=True)
    организатор_4               = db.Column(db.String(300), nullable=True)

    # Главная судейская коллегия
    главный_судья               = db.Column(db.String(200), nullable=True)
    зам_гл_судьи_судейство      = db.Column(db.String(200), nullable=True)
    начальник_дистанции         = db.Column(db.String(200), nullable=True)
    начальник_спасательной_службы = db.Column(db.String(200), nullable=True)
    главный_секретарь           = db.Column(db.String(200), nullable=True)
    зам_главного_секретаря      = db.Column(db.String(200), nullable=True)
    зам_гл_судьи_орг            = db.Column(db.String(200), nullable=True)
    председатель_комиссии_допуска = db.Column(db.String(200), nullable=True)
    председатель_мандатной_комиссии = db.Column(db.String(200), nullable=True)
    технический_делегат          = db.Column(db.String(200), nullable=True)

    # Настройки соревнований
    кол_мужских_экипажей        = db.Column(db.Integer, nullable=True)
    кол_женских_экипажей        = db.Column(db.Integer, nullable=True)
    кол_попыток_слалом          = db.Column(db.Integer, nullable=False, default=2)
    кол_ворот_слалом            = db.Column(db.Integer, nullable=False, default=10)
    кол_секций_дг               = db.Column(db.Integer, nullable=False, default=5)
    кол_спортсменов_в_экипаже   = db.Column(db.Integer, nullable=False, default=4)
    кол_субъектов               = db.Column(db.Integer, nullable=True)

    # Лимиты времени (секунды)
    лимит_время_муж             = db.Column(db.Integer, nullable=True)
    лимит_время_жен             = db.Column(db.Integer, nullable=True)

    # Номера кодов дисциплин во ВРВС
    код_спринт                  = db.Column(db.String(30), nullable=True)
    код_параллельный_спринт     = db.Column(db.String(30), nullable=True)
    код_слалом                  = db.Column(db.String(30), nullable=True)
    код_длинная_гонка           = db.Column(db.String(30), nullable=True)
    код_многоборье              = db.Column(db.String(30), nullable=True)

    # Субъект РФ (текстовое название)
    субъект_рф_название         = db.Column(db.String(200), nullable=True)

    # Главный судья — расширенные поля
    главный_судья_квалификация  = db.Column(db.String(50),  nullable=True)
    главный_судья_субъект_рф    = db.Column(db.String(200), nullable=True)
    главный_судья_телефон       = db.Column(db.String(50),  nullable=True)

    # Спортивная федерация по рафтингу
    федерация_название          = db.Column(db.String(300), nullable=True)
    федерация_руководитель      = db.Column(db.String(200), nullable=True)
    федерация_должность         = db.Column(db.String(200), nullable=True)
    федерация_телефон           = db.Column(db.String(50),  nullable=True)
    федерация_email             = db.Column(db.String(100), nullable=True)

    # Проводящая организация
    проводящая_орг_название     = db.Column(db.String(300), nullable=True)
    проводящая_орг_руководитель = db.Column(db.String(200), nullable=True)
    проводящая_орг_должность    = db.Column(db.String(200), nullable=True)
    проводящая_орг_телефон      = db.Column(db.String(50),  nullable=True)
    проводящая_орг_email        = db.Column(db.String(100), nullable=True)

    # Регламентирующие документы
    правила_вида_спорта         = db.Column(db.Text, nullable=True)
    единый_кп                   = db.Column(db.Text, nullable=True)
    положение_о_соревнованиях   = db.Column(db.Text, nullable=True)

    # Обязанности организаторов
    обязанности_федерации       = db.Column(db.Text, nullable=True)
    обязанности_проводящей      = db.Column(db.Text, nullable=True)

    # Программа соревнований — расписание мероприятий (JSON-список)
    программа_соревнований      = db.Column(db.Text, nullable=True)
    # Виды программы (JSON-список {дисциплина, пол_возраст})
    виды_список                 = db.Column(db.Text, nullable=True)
    # Последовательность проведения дисциплин (positions 2-4, comma-separated)
    последовательность_r4       = db.Column(db.String(300), nullable=True)
    последовательность_r6       = db.Column(db.String(300), nullable=True)

    # Параметры жеребьёвки и стартового протокола (Спринт)
    последовательность_видов    = db.Column(db.String(100), nullable=True)   # "R4-м,R4-ж,R6-м,R6-ж"
    вывод_старт_протокола       = db.Column(db.String(20),  nullable=True)   # "Попарно"/"Раздельно"
    старший_судья_старта        = db.Column(db.String(200), nullable=True)
    старший_судья_финиша        = db.Column(db.String(200), nullable=True)
    # Параметры стартового протокола Слалома
    последовательность_слалом   = db.Column(db.String(100), nullable=True)
    вывод_старт_слалом          = db.Column(db.String(20),  nullable=True)
    связки_ворот_слалом         = db.Column(db.Text,        nullable=True)  # JSON
    старший_судья_старта_сл     = db.Column(db.String(200), nullable=True)
    старший_судья_финиша_сл     = db.Column(db.String(200), nullable=True)

    # Параметры стартового протокола Длинной Гонки
    последовательность_дг       = db.Column(db.String(100), nullable=True)
    вывод_старт_дг              = db.Column(db.String(20),  nullable=True)
    количество_секций_дг_ход    = db.Column(db.Integer,     nullable=True)
    судьи_секций_дг             = db.Column(db.String(500), nullable=True)

    # Параметры стартового протокола Параллельного Спринта
    последовательность_пс       = db.Column(db.String(100), nullable=True)
    вывод_старт_пс              = db.Column(db.String(20),  nullable=True)
    старший_судья_старта_пс     = db.Column(db.String(200), nullable=True)
    старший_судья_финиша_пс     = db.Column(db.String(200), nullable=True)
    количество_секций_пс        = db.Column(db.Integer,     nullable=True)
    судьи_секций_пс             = db.Column(db.String(500), nullable=True)
    судья_буя_1_пс              = db.Column(db.String(200), nullable=True)
    судья_буя_2_пс              = db.Column(db.String(200), nullable=True)
    судья_буя_3_пс              = db.Column(db.String(200), nullable=True)
    судья_буя_4_пс              = db.Column(db.String(200), nullable=True)

    последовательность_допуска  = db.Column(db.String(100), nullable=True)
    вывод_допуска               = db.Column(db.String(20),  nullable=True)

    # Отчёты о соревнованиях
    отч_номер_сметы      = db.Column(db.String(100), nullable=True)
    отч_объект_спорта    = db.Column(db.Text, nullable=True)
    отч_работа_комиссии  = db.Column(db.Text, nullable=True)
    отч_тех_совещания    = db.Column(db.Text, nullable=True)
    отч_открытие         = db.Column(db.Text, nullable=True)
    отч_информационное   = db.Column(db.Text, nullable=True)
    отч_зрители          = db.Column(db.String(200), nullable=True)
    отч_оценка_качества  = db.Column(db.Text, nullable=True)
    отч_медицинское      = db.Column(db.Text, nullable=True)
    отч_безопасность     = db.Column(db.Text, nullable=True)
    отч_замечания        = db.Column(db.Text, nullable=True)
    отч_характеристика   = db.Column(db.Text, nullable=True)

    # Финансы
    заявочный_взнос_r4          = db.Column(db.String(300), nullable=True)
    заявочный_взнос             = db.Column(db.String(300), nullable=True)
    реквизиты_оплаты            = db.Column(db.Text, nullable=True)
    назначение_платежа          = db.Column(db.Text, nullable=True)

    # Логистика
    проживание_описание         = db.Column(db.Text, nullable=True)
    инфраструктура_описание     = db.Column(db.Text, nullable=True)
    проезд_общественный         = db.Column(db.Text, nullable=True)
    проезд_собственный          = db.Column(db.Text, nullable=True)
    награждение_описание        = db.Column(db.Text, nullable=True)

    создано                     = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    @property
    def организаторы_строки(self):
        """Непустые строки организаторов (1–4)."""
        return [s for s in [self.организатор_1, self.организатор_2,
                             self.организатор_3, self.организатор_4] if s]

    @property
    def организаторы_блок(self):
        """Многострочный текст организаторов; fallback на устаревшее поле."""
        строки = self.организаторы_строки
        if строки:
            return '\n'.join(строки)
        return self.организаторы or ''

    # Связи
    субъекты            = db.relationship('Субъект',  back_populates='соревнование', cascade='all, delete-orphan')
    команды             = db.relationship('Команда',  back_populates='соревнование', cascade='all, delete-orphan')
    судьи               = db.relationship('Судья',    back_populates='соревнование', cascade='all, delete-orphan')
    тех_проверки        = db.relationship('ТехПроверка', back_populates='соревнование', cascade='all, delete-orphan')
    тех_состав_комиссии = db.relationship('ТехСоставКомиссии', back_populates='соревнование', uselist=False)
    состав_комиссии_допуска = db.relationship('СоставКомиссииДопуска', back_populates='соревнование', uselist=False)
    параметры_дистанций = db.relationship('ПараметрыДистанции', back_populates='соревнование',
                                          cascade='all, delete-orphan',
                                          order_by='ПараметрыДистанции.порядок')
    параметры_награждения = db.relationship('НаграждениеПараметры', back_populates='соревнование',
                                             uselist=False, cascade='all, delete-orphan')

    def __repr__(self):
        return f'<Соревнование id={self.id} "{self.название}" {self.класс_судна}>'


class Субъект(db.Model):
    """Субъект РФ (делегация) на соревновании."""
    __tablename__ = 'субъекты'

    id                  = db.Column(db.Integer, primary_key=True)
    соревнование_id     = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)
    номер_делегации     = db.Column(db.Integer, nullable=True)
    название            = db.Column(db.String(150), nullable=False)   # территория / регион
    код_субъекта_рф     = db.Column(db.Integer, nullable=True)        # официальный код субъекта
    представитель       = db.Column(db.String(200), nullable=True)    # ФИО представителя
    тренер              = db.Column(db.String(200), nullable=True)    # ФИО тренера

    # Связи
    соревнование    = db.relationship('Соревнование', back_populates='субъекты')
    команды         = db.relationship('Команда',      back_populates='субъект', cascade='all, delete-orphan')
    спортсмены      = db.relationship('Спортсмен',    back_populates='субъект', cascade='all, delete-orphan')

    def __repr__(self):
        return f'<Субъект "{self.название}" (дел.{self.номер_делегации})>'


class Команда(db.Model):
    """Экипаж (команда) — участник соревнования."""
    __tablename__ = 'команды'

    id                  = db.Column(db.Integer, primary_key=True)
    соревнование_id     = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)
    субъект_id          = db.Column(db.Integer, db.ForeignKey('субъекты.id'),     nullable=True)

    нагрудный_номер     = db.Column(db.Integer, nullable=False)   # уникальный внутри вида программы (класс+пол)
    название            = db.Column(db.String(150), nullable=True)  # название клуба/команды
    класс_судна         = db.Column(db.String(5),   nullable=False, default='R4')  # R4 / R6
    пол                 = db.Column(db.String(5),   nullable=False, default='м')   # м / ж
    возрастная_группа   = db.Column(db.String(20),  nullable=True)   # М-Ж, до 19 …

    # Флаги участия в дистанциях
    участие_спринт          = db.Column(db.Boolean, nullable=False, default=True)
    участие_пс              = db.Column(db.Boolean, nullable=False, default=True)   # параллельный спринт
    участие_слалом          = db.Column(db.Boolean, nullable=False, default=True)
    участие_дг              = db.Column(db.Boolean, nullable=False, default=True)   # длинная гонка
    участие_многоборье      = db.Column(db.Boolean, nullable=False, default=True)

    # Принадлежность к спортивной организации и оплата взноса (комиссия по допуску)
    принадлежность_к_орг    = db.Column(db.String(100), nullable=True)
    оплата_взноса           = db.Column(db.String(50),  nullable=True, default='Ожидает оплаты')

    # Номер жеребьёвки для спринта (стартовый порядок) и плановое время старта
    стартовый_номер_спринт  = db.Column(db.Integer, nullable=True)
    время_старта_спринт_план = db.Column(db.String(20), nullable=True)

    # Связи
    соревнование    = db.relationship('Соревнование',  back_populates='команды')
    субъект         = db.relationship('Субъект',        back_populates='команды')
    спортсмены      = db.relationship('Спортсмен',      back_populates='команда', cascade='all, delete-orphan')

    результат_спринт        = db.relationship('РезультатСпринт',       back_populates='команда', uselist=False, cascade='all, delete-orphan')
    результат_пс            = db.relationship('РезультатПараллельногоСпринта', back_populates='команда', cascade='all, delete-orphan')
    результат_слалом        = db.relationship('РезультатСлалома',       back_populates='команда', cascade='all, delete-orphan')
    результат_дг            = db.relationship('РезультатДлиннойГонки',  back_populates='команда', uselist=False, cascade='all, delete-orphan')
    результат_многоборья    = db.relationship('РезультатМногоборья',    back_populates='команда', uselist=False, cascade='all, delete-orphan')

    __table_args__ = (
        db.UniqueConstraint('соревнование_id', 'класс_судна', 'пол', 'нагрудный_номер', name='uq_сор_номер'),
    )

    def __repr__(self):
        return f'<Команда №{self.нагрудный_номер} "{self.название}" {self.пол}>'


class Спортсмен(db.Model):
    """Спортсмен — член экипажа."""
    __tablename__ = 'спортсмены'

    id                      = db.Column(db.Integer, primary_key=True)
    номер_в_базе            = db.Column(db.Integer, nullable=True)   # уникальный ID в базе спортсменов
    команда_id              = db.Column(db.Integer, db.ForeignKey('команды.id'),   nullable=False)
    субъект_id              = db.Column(db.Integer, db.ForeignKey('субъекты.id'),  nullable=True)
    номер_в_экипаже         = db.Column(db.Integer, nullable=True)   # позиция в экипаже 1–7

    фамилия_имя_отчество    = db.Column(db.String(200), nullable=False)
    дата_рождения           = db.Column(db.Date,        nullable=True)
    пол                     = db.Column(db.String(5),   nullable=False, default='м')  # м / ж
    разряд                  = db.Column(db.String(20),  nullable=True)   # МС, КМС, I, II, III, Iю …
    возрастная_группа       = db.Column(db.String(20),  nullable=True)
    принадлежность_к_орг    = db.Column(db.String(100), nullable=True)   # ДЮСШ, СДЮСШОР …
    тренер_представитель    = db.Column(db.String(200), nullable=True)

    # Номер экипажа в зачёте многоборья (может отличаться от нагрудного номера)
    номер_экипажа_многоборье = db.Column(db.Integer, nullable=True)

    # Участие в каждой дистанции
    участие_спринт          = db.Column(db.Boolean, nullable=False, default=True)
    участие_пс              = db.Column(db.Boolean, nullable=False, default=True)
    участие_слалом          = db.Column(db.Boolean, nullable=False, default=True)
    участие_дг              = db.Column(db.Boolean, nullable=False, default=True)

    # Документы
    номер_приказа           = db.Column(db.String(150), nullable=True)
    номер_страховки         = db.Column(db.String(100), nullable=True)

    # Связи
    команда  = db.relationship('Команда',  back_populates='спортсмены')
    субъект  = db.relationship('Субъект',  back_populates='спортсмены')

    def __repr__(self):
        return f'<Спортсмен "{self.фамилия_имя_отчество}" {self.разряд}>'


class Судья(db.Model):
    """Спортивный судья — участник судейской коллегии."""
    __tablename__ = 'судьи'

    id                      = db.Column(db.Integer, primary_key=True)
    номер_в_базе            = db.Column(db.Integer, nullable=True)
    соревнование_id         = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)

    фамилия_имя_отчество    = db.Column(db.String(200), nullable=False)
    квалификация            = db.Column(db.String(20),  nullable=True)   # ССВК, СС1К, …
    субъект_рф              = db.Column(db.String(150), nullable=True)   # Московская область …
    код_субъекта_рф         = db.Column(db.Integer,     nullable=True)
    пол                     = db.Column(db.String(5),   nullable=True)   # м / ж
    дата_рождения           = db.Column(db.Date,        nullable=True)
    дата_присвоения         = db.Column(db.Date,        nullable=True)
    дата_подтверждения      = db.Column(db.Date,        nullable=True)
    конец_аттестации        = db.Column(db.Date,        nullable=True)
    действующая_категория   = db.Column(db.Boolean,     nullable=True)

    # Должность на данном соревновании
    должность               = db.Column(db.String(100), nullable=True)

    # Связи
    соревнование = db.relationship('Соревнование', back_populates='судьи')

    def __repr__(self):
        return f'<Судья "{self.фамилия_имя_отчество}" {self.квалификация}>'


# ---------------------------------------------------------------------------
# Техническая комиссия
# ---------------------------------------------------------------------------

class ТехПроверка(db.Model):
    """Техническая проверка одной команды на соревновании."""
    __tablename__ = 'тех_проверки'

    id                   = db.Column(db.Integer, primary_key=True)
    соревнование_id      = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)
    субъект_рф           = db.Column(db.String(200), nullable=True)
    наименование_команды = db.Column(db.String(200), nullable=True)
    тренер_фио           = db.Column(db.String(200), nullable=True)
    создано              = db.Column(db.DateTime,    default=datetime.utcnow)

    снаряжение   = db.relationship('ТехСнаряжение', back_populates='проверка',
                                   cascade='all, delete-orphan', order_by='ТехСнаряжение.порядок')
    соревнование = db.relationship('Соревнование', back_populates='тех_проверки')

    def __repr__(self):
        return f'<ТехПроверка id={self.id} "{self.наименование_команды}">'


class ТехСнаряжение(db.Model):
    """Единица снаряжения в технической проверке."""
    __tablename__ = 'тех_снаряжение'

    id              = db.Column(db.Integer, primary_key=True)
    проверка_id     = db.Column(db.Integer, db.ForeignKey('тех_проверки.id'), nullable=False)
    наименование    = db.Column(db.String(50),  nullable=False)   # Рафт R4 / Рафт R6 / Спасжилет / Каска / Весло
    марка           = db.Column(db.String(100), nullable=True)
    серийный_номер  = db.Column(db.String(100), nullable=True)
    количество      = db.Column(db.Integer,     nullable=True, default=1)
    допуск          = db.Column(db.String(20),  nullable=True, default='Допущен')
    порядок         = db.Column(db.Integer,     nullable=True, default=0)

    проверка = db.relationship('ТехПроверка', back_populates='снаряжение')


class ТехСоставКомиссии(db.Model):
    """Состав технической комиссии на соревновании (одна запись на соревнование)."""
    __tablename__ = 'тех_состав_комиссии'

    id                       = db.Column(db.Integer, primary_key=True)
    соревнование_id          = db.Column(db.Integer, db.ForeignKey('соревнования.id'), unique=True, nullable=False)
    председатель_фио           = db.Column(db.String(200), nullable=True)
    председатель_должность     = db.Column(db.String(200), nullable=True)
    председатель_квалификация  = db.Column(db.String(100), nullable=True)
    председатель_субъект_рф    = db.Column(db.String(200), nullable=True)
    член_1_фио                 = db.Column(db.String(200), nullable=True)
    член_1_должность           = db.Column(db.String(200), nullable=True)
    член_1_квалификация        = db.Column(db.String(100), nullable=True)
    член_1_субъект_рф          = db.Column(db.String(200), nullable=True)
    член_2_фио                 = db.Column(db.String(200), nullable=True)
    член_2_должность           = db.Column(db.String(200), nullable=True)
    член_2_квалификация        = db.Column(db.String(100), nullable=True)
    член_2_субъект_рф          = db.Column(db.String(200), nullable=True)

    соревнование = db.relationship('Соревнование', back_populates='тех_состав_комиссии')


class СоставКомиссииДопуска(db.Model):
    """Состав комиссии по допуску на соревновании (одна запись на соревнование)."""
    __tablename__ = 'состав_комиссии_допуска'

    id                       = db.Column(db.Integer, primary_key=True)
    соревнование_id          = db.Column(db.Integer, db.ForeignKey('соревнования.id'), unique=True, nullable=False)
    председатель_фио           = db.Column(db.String(200), nullable=True)
    председатель_должность     = db.Column(db.String(200), nullable=True)
    председатель_квалификация  = db.Column(db.String(100), nullable=True)
    председатель_субъект_рф    = db.Column(db.String(200), nullable=True)
    член_1_фио                 = db.Column(db.String(200), nullable=True)
    член_1_должность           = db.Column(db.String(200), nullable=True)
    член_1_квалификация        = db.Column(db.String(100), nullable=True)
    член_1_субъект_рф          = db.Column(db.String(200), nullable=True)
    член_2_фио                 = db.Column(db.String(200), nullable=True)
    член_2_должность           = db.Column(db.String(200), nullable=True)
    член_2_квалификация        = db.Column(db.String(100), nullable=True)
    член_2_субъект_рф          = db.Column(db.String(200), nullable=True)

    соревнование = db.relationship('Соревнование', back_populates='состав_комиссии_допуска')


class ПараметрыДистанции(db.Model):
    """Параметры одной дистанции в рамках соревнования (до 4 дистанций)."""
    __tablename__ = 'параметры_дистанций'

    id              = db.Column(db.Integer, primary_key=True)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)
    порядок         = db.Column(db.Integer, nullable=False, default=1)

    название        = db.Column(db.String(200), nullable=True)
    протяженность   = db.Column(db.String(50),  nullable=True)
    категория       = db.Column(db.String(50),  nullable=True)
    среднее_время   = db.Column(db.String(50),  nullable=True)
    лимит_времени   = db.Column(db.String(50),  nullable=True)
    вариант_старта    = db.Column(db.Text,        nullable=True)
    вариант_финиша    = db.Column(db.Text,        nullable=True)
    страховка         = db.Column(db.Text,        nullable=True)
    тип               = db.Column(db.String(50),  nullable=False, default='спринт')
    буи               = db.Column(db.Text,        nullable=True)
    кол_ворот         = db.Column(db.String(50),  nullable=True)
    кол_обратных      = db.Column(db.String(50),  nullable=True)
    конструкция_ворот = db.Column(db.Text,        nullable=True)

    соревнование = db.relationship('Соревнование', back_populates='параметры_дистанций')

    def __repr__(self):
        return f'<ПараметрыДистанции #{self.порядок} сор={self.соревнование_id}>'


class НаграждениеПараметры(db.Model):
    """Параметры дипломов для страницы Награждение."""
    __tablename__ = 'награждение_параметры'

    id                        = db.Column(db.Integer, primary_key=True)
    соревнование_id           = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False, unique=True)

    использует_награждается   = db.Column(db.Boolean, nullable=False, default=True)
    отступ_сверху             = db.Column(db.Integer, nullable=False, default=75)
    отступ_снизу              = db.Column(db.Integer, nullable=False, default=25)
    отступ_слева              = db.Column(db.Integer, nullable=False, default=15)
    отступ_справа             = db.Column(db.Integer, nullable=False, default=15)
    размер_шрифта_текста      = db.Column(db.Integer, nullable=False, default=14)
    наименование_соревнований = db.Column(db.Text,    nullable=True)
    должность_подписывающего  = db.Column(db.Text,    nullable=True)
    ио_фамилия_подписывающего = db.Column(db.String(200), nullable=True)

    соревнование = db.relationship('Соревнование', back_populates='параметры_награждения')

    def __repr__(self):
        return f'<НаграждениеПараметры сор={self.соревнование_id}>'


# ---------------------------------------------------------------------------
# Результаты по дистанциям
# ---------------------------------------------------------------------------

# Общие статусы результата
СТАТУС_ФИНИШ   = 'ФИН'   # финишировал
СТАТУС_ДНС     = 'НСТ'   # не стартовал
СТАТУС_ДНФ     = 'НФН'   # не финишировал
СТАТУС_ДК      = 'ДК'    # дисквалификация
СТАТУС_ПМ      = 'ПМ'    # последнее место (штраф без времени)
СТАТУС_ПЛВ     = 'ПЛВ'   # превышение лимита времени


class РезультатСпринт(db.Model):
    """Результат экипажа на дистанции СПРИНТ."""
    __tablename__ = 'результаты_спринт'

    id              = db.Column(db.Integer, primary_key=True)
    команда_id      = db.Column(db.Integer, db.ForeignKey('команды.id'), nullable=False, unique=True)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)

    стартовый_номер = db.Column(db.Integer, nullable=True)   # порядок старта по жребию

    # Время (секунды с дробной частью)
    время_старта    = db.Column(db.Float, nullable=True)
    время_финиша    = db.Column(db.Float, nullable=True)
    время_прохождения = db.Column(db.Float, nullable=True)   # финиш − старт

    # Штрафы старта: кол-во нарушений по типу
    штраф_с10         = db.Column(db.Integer, nullable=False, default=0)   # +10 сек
    штраф_с50         = db.Column(db.Integer, nullable=False, default=0)   # +50 сек
    штраф_ф50         = db.Column(db.Integer, nullable=False, default=0)   # финиш +50 сек

    # Агрегированные поля (для совместимости)
    кол_штраф_старт   = db.Column(db.Integer, nullable=False, default=0)
    кол_штраф_финиш   = db.Column(db.Integer, nullable=False, default=0)
    штрафное_время    = db.Column(db.Float,   nullable=False, default=0)   # суммарно в сек

    результат   = db.Column(db.Float,   nullable=True)   # время_прохождения + штрафное_время
    место       = db.Column(db.Integer, nullable=True)
    очки        = db.Column(db.Float,   nullable=False, default=0)
    статус      = db.Column(db.String(5), nullable=False, default=СТАТУС_ФИНИШ)

    # Связи
    команда = db.relationship('Команда', back_populates='результат_спринт')

    def __repr__(self):
        return f'<РезультатСпринт команда_id={self.команда_id} место={self.место}>'


class РезультатПараллельногоСпринта(db.Model):
    """Результат одного заезда параллельного спринта."""
    __tablename__ = 'результаты_пс'

    id              = db.Column(db.Integer, primary_key=True)
    команда_id      = db.Column(db.Integer, db.ForeignKey('команды.id'), nullable=False)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)

    # Стадия турнирной сетки
    стадия          = db.Column(db.String(20), nullable=False)

    номер_заезда    = db.Column(db.Integer, nullable=True)
    дорожка         = db.Column(db.String(10), nullable=True)   # 'левая' / 'правая'

    время_старта    = db.Column(db.Float, nullable=True)
    время_финиша    = db.Column(db.Float, nullable=True)
    время           = db.Column(db.Float, nullable=True)   # время_финиша − время_старта (или введено напрямую)
    штраф_с10       = db.Column(db.Integer, nullable=False, default=0)   # +10 сек на старте
    штраф_ст50      = db.Column(db.Integer, nullable=False, default=0)   # +50 сек на старте
    # Прохождение буёв (ставят судьи на буях с телефона): True/False/None (ещё не отмечено)
    буй1            = db.Column(db.Boolean, nullable=True)
    буй2            = db.Column(db.Boolean, nullable=True)
    буй3            = db.Column(db.Boolean, nullable=True)
    буй4            = db.Column(db.Boolean, nullable=True)
    штраф_буй       = db.Column(db.Integer, nullable=False, default=0)   # +50 сек на буях (0/1/2 — считается по буям)
    штраф_с50       = db.Column(db.Integer, nullable=False, default=0)   # +50 сек на финише
    штрафное_время  = db.Column(db.Float, nullable=False, default=0)
    результат       = db.Column(db.Float, nullable=True)   # время + штраф
    победитель      = db.Column(db.Boolean, nullable=True)  # выиграл заезд

    # Итоговый результат по дистанции (заполняется после всех стадий)
    итоговое_место  = db.Column(db.Integer, nullable=True)
    очки            = db.Column(db.Float,   nullable=False, default=0)
    статус          = db.Column(db.String(5), nullable=False, default=СТАТУС_ФИНИШ)

    # Связи
    команда = db.relationship('Команда', back_populates='результат_пс')

    def __repr__(self):
        return f'<РезультатПС команда_id={self.команда_id} стадия={self.стадия}>'


class РезультатСлалома(db.Model):
    """Результат одной попытки на дистанции СЛАЛОМ (до 2 попыток на команду)."""
    __tablename__ = 'результаты_слалом'

    id              = db.Column(db.Integer, primary_key=True)
    команда_id      = db.Column(db.Integer, db.ForeignKey('команды.id'), nullable=False)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)

    номер_попытки   = db.Column(db.Integer, nullable=False, default=1)   # 1 или 2

    время_старта        = db.Column(db.Float, nullable=True)
    время_финиша        = db.Column(db.Float, nullable=True)
    время_прохождения   = db.Column(db.Float, nullable=True)

    # Штрафы слалома (количество нарушений каждого типа)
    кол_штраф_5         = db.Column(db.Integer, nullable=False, default=0)   # касание вехи (5 сек)
    кол_штраф_50        = db.Column(db.Integer, nullable=False, default=0)   # неправильный заход (50 сек)
    кол_штраф_с10       = db.Column(db.Integer, nullable=False, default=0)   # +10с на старте
    кол_штраф_старт     = db.Column(db.Integer, nullable=False, default=0)   # +50с на старте
    кол_штраф_финиш     = db.Column(db.Integer, nullable=False, default=0)   # +50с на финише
    штрафное_время      = db.Column(db.Float,   nullable=False, default=0)   # суммарный штраф

    результат_попытки   = db.Column(db.Float, nullable=True)   # время + штраф данной попытки
    статус_попытки      = db.Column(db.String(5), nullable=False, default=СТАТУС_ФИНИШ)

    # Итог по дистанции (лучшая попытка — заполняется после обеих попыток)
    лучший_результат    = db.Column(db.Float,   nullable=True)
    место               = db.Column(db.Integer, nullable=True)
    очки                = db.Column(db.Float,   nullable=False, default=0)
    статус              = db.Column(db.String(5), nullable=False, default=СТАТУС_ФИНИШ)

    __table_args__ = (
        db.UniqueConstraint('команда_id', 'номер_попытки', name='uq_слалом_попытка'),
    )

    # Связи
    команда = db.relationship('Команда', back_populates='результат_слалом')

    def __repr__(self):
        return f'<РезультатСлалом команда_id={self.команда_id} попытка={self.номер_попытки}>'


class ШтрафВорот(db.Model):
    """Штраф на отдельных воротах слалома (5 сек или 50 сек)."""
    __tablename__ = 'штрафы_ворот'

    id              = db.Column(db.Integer, primary_key=True)
    команда_id      = db.Column(db.Integer, db.ForeignKey('команды.id'), nullable=False)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)
    номер_попытки   = db.Column(db.Integer, nullable=False)    # 1 или 2
    номер_ворот     = db.Column(db.Integer, nullable=False)    # 1–20
    штраф           = db.Column(db.Integer, nullable=False, default=0)   # 0 / 5 / 50

    __table_args__ = (
        db.UniqueConstraint('команда_id', 'номер_попытки', 'номер_ворот', name='uq_ворота_штраф'),
    )

    команда = db.relationship('Команда')

    def __repr__(self):
        return f'<ШтрафВорот команда_id={self.команда_id} попытка={self.номер_попытки} ворота={self.номер_ворот} = {self.штраф}>'


class ШтрафСекции(db.Model):
    """Количество нарушений на отдельной секции длинной гонки (каждое = 50 сек)."""
    __tablename__ = 'штрафы_секций'

    id              = db.Column(db.Integer, primary_key=True)
    команда_id      = db.Column(db.Integer, db.ForeignKey('команды.id'), nullable=False)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)
    номер_секции    = db.Column(db.Integer, nullable=False)    # 1–10
    кол_нарушений   = db.Column(db.Integer, nullable=False, default=0)

    __table_args__ = (
        db.UniqueConstraint('команда_id', 'соревнование_id', 'номер_секции', name='uq_секции_штраф'),
    )

    команда = db.relationship('Команда')

    def __repr__(self):
        return f'<ШтрафСекции команда_id={self.команда_id} секция={self.номер_секции} кол={self.кол_нарушений}>'


class РезультатДлиннойГонки(db.Model):
    """Результат экипажа на дистанции ДЛИННАЯ ГОНКА."""
    __tablename__ = 'результаты_дг'

    id              = db.Column(db.Integer, primary_key=True)
    команда_id      = db.Column(db.Integer, db.ForeignKey('команды.id'), nullable=False, unique=True)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)

    стартовый_номер     = db.Column(db.Integer, nullable=True)
    время_старта        = db.Column(db.Float, nullable=True)
    время_финиша        = db.Column(db.Float, nullable=True)
    время_прохождения   = db.Column(db.Float, nullable=True)

    кол_штраф_с10       = db.Column(db.Integer, nullable=False, default=0)
    кол_штраф_старт     = db.Column(db.Integer, nullable=False, default=0)
    кол_штраф_финиш     = db.Column(db.Integer, nullable=False, default=0)
    штрафное_время      = db.Column(db.Float,   nullable=False, default=0)

    результат   = db.Column(db.Float,   nullable=True)
    место       = db.Column(db.Integer, nullable=True)
    очки        = db.Column(db.Float,   nullable=False, default=0)
    статус      = db.Column(db.String(5), nullable=False, default=СТАТУС_ФИНИШ)

    # Связи
    команда = db.relationship('Команда', back_populates='результат_дг')

    def __repr__(self):
        return f'<РезультатДГ команда_id={self.команда_id} место={self.место}>'


class РезультатМногоборья(db.Model):
    """Суммарный результат многоборья (сумма очков по всем дистанциям)."""
    __tablename__ = 'результаты_многоборье'

    id              = db.Column(db.Integer, primary_key=True)
    команда_id      = db.Column(db.Integer, db.ForeignKey('команды.id'), nullable=False, unique=True)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)

    очки_спринт     = db.Column(db.Float, nullable=False, default=0)
    очки_пс         = db.Column(db.Float, nullable=False, default=0)
    очки_слалом     = db.Column(db.Float, nullable=False, default=0)
    очки_дг         = db.Column(db.Float, nullable=False, default=0)
    сумма_очков     = db.Column(db.Float, nullable=False, default=0)

    место           = db.Column(db.Integer, nullable=True)
    очки_многоборье = db.Column(db.Float,   nullable=False, default=0)   # очки за место в многоборье

    # Связи
    команда = db.relationship('Команда', back_populates='результат_многоборья')

    def __repr__(self):
        return f'<РезультатМногоборья команда_id={self.команда_id} сумма={self.сумма_очков}>'


class ЗачётСубъектов(db.Model):
    """Итоговый командный зачёт субъектов РФ по таблице коэффициентов."""
    __tablename__ = 'зачет_субъектов'

    id              = db.Column(db.Integer, primary_key=True)
    субъект_id      = db.Column(db.Integer, db.ForeignKey('субъекты.id'), nullable=False)
    соревнование_id = db.Column(db.Integer, db.ForeignKey('соревнования.id'), nullable=False)

    # Очки субъекта по каждой дистанции (сумма по всем экипажам субъекта)
    очки_спринт     = db.Column(db.Float, nullable=False, default=0)
    очки_пс         = db.Column(db.Float, nullable=False, default=0)
    очки_слалом     = db.Column(db.Float, nullable=False, default=0)
    очки_дг         = db.Column(db.Float, nullable=False, default=0)
    очки_многоборье = db.Column(db.Float, nullable=False, default=0)
    итого_очков     = db.Column(db.Float, nullable=False, default=0)

    место           = db.Column(db.Integer, nullable=True)

    __table_args__ = (
        db.UniqueConstraint('субъект_id', 'соревнование_id', name='uq_субъект_сор'),
    )

    # Связи
    субъект = db.relationship('Субъект')

    def __repr__(self):
        return f'<ЗачётСубъектов субъект_id={self.субъект_id} итого={self.итого_очков}>'


# ---------------------------------------------------------------------------
# Вспомогательные функции инициализации БД
# ---------------------------------------------------------------------------

# Данные таблицы очков из листа «Очки» СЧИТАЛКА.xls
ДАННЫЕ_ОЧКОВ = {
    'СПРИНТ':              {1: 100, 2: 95,  3: 90,  4: 85,  5: 80,  6: 75,  7: 70,  8: 65,
                            9: 60,  10: 55, 11: 50, 12: 45, 13: 40, 14: 35, 15: 30, 16: 25,
                            17: 20, 18: 15, 19: 10, 20: 5},
    'ПАРАЛЛЕЛЬНЫЙ_СПРИНТ': {1: 200, 2: 190, 3: 180, 4: 170, 5: 160, 6: 150, 7: 140, 8: 130,
                            9: 120, 10: 110, 11: 100, 12: 90, 13: 80, 14: 70, 15: 60, 16: 50,
                            17: 40, 18: 30, 19: 20, 20: 10},
    'СЛАЛОМ':              {1: 300, 2: 285, 3: 270, 4: 255, 5: 240, 6: 225, 7: 210, 8: 195,
                            9: 180, 10: 165, 11: 150, 12: 135, 13: 120, 14: 105, 15: 90, 16: 75,
                            17: 60, 18: 45, 19: 30, 20: 15},
    'ДЛИННАЯ_ГОНКА':       {1: 400, 2: 380, 3: 360, 4: 340, 5: 320, 6: 300, 7: 280, 8: 260,
                            9: 240, 10: 220, 11: 200, 12: 180, 13: 160, 14: 140, 15: 120, 16: 100,
                            17: 80, 18: 60, 19: 40, 20: 20},
    'МНОГОБОРЬЕ':          {1: 400, 2: 380, 3: 360, 4: 340, 5: 320, 6: 300, 7: 280, 8: 260,
                            9: 240, 10: 220, 11: 200, 12: 180, 13: 160, 14: 140, 15: 120, 16: 100,
                            17: 80, 18: 60, 19: 40, 20: 20},
}


def заполнить_таблицу_очков():
    """Заполнить ТаблицаОчков начальными данными, если пустая."""
    if ТаблицаОчков.query.first():
        return
    for дистанция, места in ДАННЫЕ_ОЧКОВ.items():
        for место, очки in места.items():
            db.session.add(ТаблицаОчков(дистанция=дистанция, место=место, очки=очки))
    db.session.commit()


def получить_очки(дистанция: str, место: int) -> float:
    """Вернуть количество очков за занятое место на дистанции."""
    запись = ТаблицаОчков.query.filter_by(дистанция=дистанция, место=место).first()
    return запись.очки if запись else 0.0


ДАННЫЕ_КОЭФФИЦИЕНТОВ_ЗС = {
    # кол_субъектов: (место_1, место_2, место_3, место_4+)
    3:  (3.5, 2.3, 1.2, 1.0),   # 3 и менее
    4:  (4.6, 3.0, 1.5, 1.0),
    5:  (5.8, 3.8, 1.9, 1.0),
    6:  (6.9, 4.6, 2.3, 1.0),
    7:  (8.1, 5.4, 2.7, 1.0),
    8:  (9.2, 6.2, 3.1, 1.0),
    9:  (10.4, 6.9, 3.5, 1.0),
    10: (11.5, 7.7, 3.8, 1.0),
    11: (12.7, 8.5, 4.2, 1.0),
    12: (13.8, 9.2, 4.6, 1.0),
    13: (15.0, 10.0, 5.0, 1.0),  # 13 и более
}


def заполнить_коэффициенты_субъектов():
    """Заполнить таблицу ЗС коэффициентами из официальной Таблицы ЗС."""
    if КоэффициентыСубъектов.query.first():
        return
    for кол, (к1, к2, к3, к4) in ДАННЫЕ_КОЭФФИЦИЕНТОВ_ЗС.items():
        for место, коэф in [(1, к1), (2, к2), (3, к3), (4, к4)]:
            db.session.add(КоэффициентыСубъектов(место=место, кол_субъектов=кол, коэффициент=коэф))
    db.session.commit()


def заполнить_штрафы():
    """Заполнить справочник штрафов стандартными нарушениями для спринта."""
    if Штраф.query.filter_by(дистанция='СТАРТ').first():
        return
    штрафы = [
        Штраф(дистанция='СТАРТ', номер=1, нарушение='Преждевременный старт',
              штраф_сек=10, дисквалификация=False),
        Штраф(дистанция='СТАРТ', номер=2, нарушение='Нарушение стартовой позиции',
              штраф_сек=50, дисквалификация=False),
        Штраф(дистанция='СТАРТ', номер=3, нарушение='Дисквалификация на старте',
              штраф_сек=None, дисквалификация=True),
        Штраф(дистанция='ФИНИШ', номер=1, нарушение='Нарушение финишного коридора',
              штраф_сек=50, дисквалификация=False),
        Штраф(дистанция='ФИНИШ', номер=2, нарушение='Финиш вне трассы (последнее место)',
              штраф_сек=None, дисквалификация=False, примечание='ПМ'),
        Штраф(дистанция='ФИНИШ', номер=3, нарушение='Дисквалификация на финише',
              штраф_сек=None, дисквалификация=True),
    ]
    for ш in штрафы:
        db.session.add(ш)
    db.session.commit()


# ---------------------------------------------------------------------------
# Предварительная заявка
# ---------------------------------------------------------------------------

class ПредварительнаяЗаявка(db.Model):
    __tablename__ = 'предварительные_заявки'

    id                    = db.Column(db.Integer, primary_key=True)
    соревнование_id       = db.Column(db.Integer, db.ForeignKey('соревнования.id', ondelete='CASCADE'), nullable=False)
    субъект_рф            = db.Column(db.String(200), nullable=True)
    наименование_команды  = db.Column(db.String(200), nullable=True)
    # Организация
    орг_наименование           = db.Column(db.String(300), nullable=True)
    орг_фио_руководителя       = db.Column(db.String(200), nullable=True)
    орг_должность_руководителя = db.Column(db.String(200), nullable=True)
    орг_телефон                = db.Column(db.String(50),  nullable=True)
    орг_email                  = db.Column(db.String(100), nullable=True)
    # Тренер-представитель
    тренер_фио            = db.Column(db.String(200), nullable=True)
    тренер_телефон        = db.Column(db.String(50),  nullable=True)
    тренер_email          = db.Column(db.String(100), nullable=True)
    создано               = db.Column(db.DateTime, default=datetime.utcnow)

    соревнование = db.relationship('Соревнование', backref='предварительные_заявки')
    спортсмены   = db.relationship('СпортсменЗаявки', backref='заявка',
                                   cascade='all, delete-orphan',
                                   order_by='[СпортсменЗаявки.фамилия, СпортсменЗаявки.имя, СпортсменЗаявки.отчество]')


class СпортсменЗаявки(db.Model):
    __tablename__ = 'спортсмены_заявки'

    id              = db.Column(db.Integer, primary_key=True)
    заявка_id       = db.Column(db.Integer, db.ForeignKey('предварительные_заявки.id', ondelete='CASCADE'), nullable=False)
    фамилия         = db.Column(db.String(100), nullable=True)
    имя             = db.Column(db.String(100), nullable=True)
    отчество        = db.Column(db.String(100), nullable=True)
    пол             = db.Column(db.String(5),   nullable=True)
    дата_рождения   = db.Column(db.Date,        nullable=True)
    квалификация    = db.Column(db.String(50),  nullable=True)
    номер_приказа   = db.Column(db.String(150), nullable=True)
    номер_страховки = db.Column(db.String(100), nullable=True)
    возрастная_группа = db.Column(db.String(50), nullable=True)


def применить_миграции():
    """Добавить недостающие колонки в существующую БД (идемпотентно)."""
    from sqlalchemy import inspect, text
    inspector = inspect(db.engine)
    with db.engine.connect() as conn:
        # субъекты: представитель, тренер
        cols = {c['name'] for c in inspector.get_columns('субъекты')}
        if 'представитель' not in cols:
            conn.execute(text(
                'ALTER TABLE субъекты ADD COLUMN представитель VARCHAR(200)'
            ))
            conn.commit()
        if 'тренер' not in cols:
            conn.execute(text(
                'ALTER TABLE субъекты ADD COLUMN тренер VARCHAR(200)'
            ))
            conn.commit()
        # команды: класс_судна
        cols = {c['name'] for c in inspector.get_columns('команды')}
        if 'класс_судна' not in cols:
            conn.execute(text(
                "ALTER TABLE команды ADD COLUMN класс_судна VARCHAR(5) NOT NULL DEFAULT 'R4'"
            ))
            conn.commit()
        # спортсмены: номер_в_экипаже
        cols = {c['name'] for c in inspector.get_columns('спортсмены')}
        if 'номер_в_экипаже' not in cols:
            conn.execute(text(
                'ALTER TABLE спортсмены ADD COLUMN номер_в_экипаже INTEGER'
            ))
            conn.commit()
        # соревнования: кол_ворот_слалом, кол_секций_дг
        cols = {c['name'] for c in inspector.get_columns('соревнования')}
        if 'кол_ворот_слалом' not in cols:
            conn.execute(text(
                'ALTER TABLE соревнования ADD COLUMN кол_ворот_слалом INTEGER NOT NULL DEFAULT 10'
            ))
            conn.commit()
        if 'кол_секций_дг' not in cols:
            conn.execute(text(
                'ALTER TABLE соревнования ADD COLUMN кол_секций_дг INTEGER NOT NULL DEFAULT 5'
            ))
            conn.commit()
        # соревнования: новые поля организаторов и судейской коллегии
        cols = {c['name'] for c in inspector.get_columns('соревнования')}
        for col, ddl in [
            ('организатор_1',          'VARCHAR(300)'),
            ('организатор_2',          'VARCHAR(300)'),
            ('организатор_3',          'VARCHAR(300)'),
            ('организатор_4',          'VARCHAR(300)'),
            ('зам_гл_судьи_судейство',       'VARCHAR(200)'),
            ('зам_гл_судьи_орг',             'VARCHAR(200)'),
            ('председатель_мандатной_комиссии', 'VARCHAR(200)'),
            ('технический_делегат',          'VARCHAR(200)'),
            ('номер_екп',                    'VARCHAR(100)'),
            ('дата_начала',                  'DATE'),
        ]:
            if col not in cols:
                conn.execute(text(f'ALTER TABLE соревнования ADD COLUMN {col} {ddl}'))
                conn.commit()
        # соревнования: коды дисциплин во ВРВС и расширенные поля формы
        cols = {c['name'] for c in inspector.get_columns('соревнования')}
        for col, ddl in [
            ('код_спринт',                  'VARCHAR(30)'),
            ('код_параллельный_спринт',     'VARCHAR(30)'),
            ('код_слалом',                  'VARCHAR(30)'),
            ('код_длинная_гонка',           'VARCHAR(30)'),
            ('код_многоборье',              'VARCHAR(30)'),
            ('субъект_рф_название',         'VARCHAR(200)'),
            ('главный_судья_квалификация',  'VARCHAR(50)'),
            ('главный_судья_субъект_рф',    'VARCHAR(200)'),
            ('главный_судья_телефон',       'VARCHAR(50)'),
            ('федерация_название',           'VARCHAR(300)'),
            ('федерация_руководитель',       'VARCHAR(200)'),
            ('федерация_должность',          'VARCHAR(200)'),
            ('федерация_телефон',            'VARCHAR(50)'),
            ('федерация_email',              'VARCHAR(100)'),
            ('проводящая_орг_название',     'VARCHAR(300)'),
            ('проводящая_орг_руководитель', 'VARCHAR(200)'),
            ('проводящая_орг_должность',    'VARCHAR(200)'),
            ('проводящая_орг_телефон',      'VARCHAR(50)'),
            ('проводящая_орг_email',        'VARCHAR(100)'),
            ('правила_вида_спорта',         'TEXT'),
            ('единый_кп',                   'TEXT'),
            ('положение_о_соревнованиях',   'TEXT'),
            ('обязанности_федерации',       'TEXT'),
            ('обязанности_проводящей',      'TEXT'),
            ('программа_соревнований',      'TEXT'),
            ('заявочный_взнос_r4',          'VARCHAR(300)'),
            ('заявочный_взнос',             'VARCHAR(300)'),
            ('реквизиты_оплаты',            'TEXT'),
            ('назначение_платежа',          'TEXT'),
            ('проживание_описание',         'TEXT'),
            ('инфраструктура_описание',     'TEXT'),
            ('проезд_общественный',         'TEXT'),
            ('проезд_собственный',          'TEXT'),
            ('награждение_описание',        'TEXT'),
            ('виды_список',                  'TEXT'),
            ('последовательность_r4',        'VARCHAR(300)'),
            ('последовательность_r6',        'VARCHAR(300)'),
        ]:
            if col not in cols:
                conn.execute(text(f'ALTER TABLE соревнования ADD COLUMN {col} {ddl}'))
                conn.commit()
        # спортсмены_заявки: возрастная_группа
        if inspector.has_table('спортсмены_заявки'):
            cols = {c['name'] for c in inspector.get_columns('спортсмены_заявки')}
            if 'возрастная_группа' not in cols:
                conn.execute(text('ALTER TABLE спортсмены_заявки ADD COLUMN возрастная_группа VARCHAR(50)'))
                conn.commit()
        # команды: время_старта_спринт_план
        cols = {c['name'] for c in inspector.get_columns('команды')}
        if 'время_старта_спринт_план' not in cols:
            conn.execute(text("ALTER TABLE команды ADD COLUMN время_старта_спринт_план VARCHAR(20)"))
            conn.commit()
        # команды: принадлежность_к_орг, оплата_взноса
        cols = {c['name'] for c in inspector.get_columns('команды')}
        if 'принадлежность_к_орг' not in cols:
            conn.execute(text("ALTER TABLE команды ADD COLUMN принадлежность_к_орг VARCHAR(100)"))
            conn.commit()
        if 'оплата_взноса' not in cols:
            conn.execute(text("ALTER TABLE команды ADD COLUMN оплата_взноса VARCHAR(50) DEFAULT 'Ожидает оплаты'"))
            conn.commit()
        # спортсмены: документы
        cols = {c['name'] for c in inspector.get_columns('спортсмены')}
        if 'номер_приказа' not in cols:
            conn.execute(text('ALTER TABLE спортсмены ADD COLUMN номер_приказа VARCHAR(150)'))
            conn.commit()
        if 'номер_страховки' not in cols:
            conn.execute(text('ALTER TABLE спортсмены ADD COLUMN номер_страховки VARCHAR(100)'))
            conn.commit()
        # соревнования: статус готовности разделов
        cols = {c['name'] for c in inspector.get_columns('соревнования')}
        for col in ('готов_организаторы', 'готов_судьи', 'готов_заявки',
                    'готов_допуск', 'готов_тех',
                    'готов_наград_спринт', 'готов_наград_пс', 'готов_наград_слалом',
                    'готов_наград_дг', 'готов_наград_мб', 'готов_наград_зачет'):
            if col not in cols:
                conn.execute(text(
                    f"ALTER TABLE соревнования ADD COLUMN {col} BOOLEAN NOT NULL DEFAULT 0"
                ))
                conn.commit()
        # тех_состав_комиссии: квалификация и субъект членов
        if inspector.has_table('тех_состав_комиссии'):
            cols = {c['name'] for c in inspector.get_columns('тех_состав_комиссии')}
            for col, ddl in [
                ('председатель_квалификация', 'VARCHAR(100)'),
                ('председатель_субъект_рф',   'VARCHAR(200)'),
                ('член_1_квалификация',        'VARCHAR(100)'),
                ('член_1_субъект_рф',          'VARCHAR(200)'),
                ('член_2_квалификация',        'VARCHAR(100)'),
                ('член_2_субъект_рф',          'VARCHAR(200)'),
            ]:
                if col not in cols:
                    conn.execute(text(f'ALTER TABLE тех_состав_комиссии ADD COLUMN {col} {ddl}'))
                    conn.commit()
        # результаты_спринт: новые колонки штрафов по типу
        if inspector.has_table('результаты_спринт'):
            cols = {c['name'] for c in inspector.get_columns('результаты_спринт')}
            for col in ('штраф_с10', 'штраф_с50', 'штраф_ф50'):
                if col not in cols:
                    conn.execute(text(
                        f'ALTER TABLE результаты_спринт ADD COLUMN {col} INTEGER DEFAULT 0'
                    ))
                    conn.commit()
        # параметры_дистанций: создаём таблицу если нет, добавляем новые столбцы
        if not inspector.has_table('параметры_дистанций'):
            db.create_all()
        else:
            cols_д = {c['name'] for c in inspector.get_columns('параметры_дистанций')}
            for col, ddl in [
                ('тип',               "VARCHAR(50) NOT NULL DEFAULT 'спринт'"),
                ('буи',               'TEXT'),
                ('кол_ворот',         'VARCHAR(50)'),
                ('кол_обратных',      'VARCHAR(50)'),
                ('конструкция_ворот', 'TEXT'),
            ]:
                if col not in cols_д:
                    conn.execute(text(f'ALTER TABLE параметры_дистанций ADD COLUMN {col} {ddl}'))
                    conn.commit()
        # награждение_параметры: создаём таблицу если нет
        if not inspector.has_table('награждение_параметры'):
            db.create_all()
        # параметры жеребьёвки: два новых столбца в соревнованиях
        cols_сор = {c['name'] for c in inspector.get_columns('соревнования')}
        if 'последовательность_видов' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN последовательность_видов VARCHAR(100)"
                ))
                conn.commit()
        if 'вывод_старт_протокола' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN вывод_старт_протокола VARCHAR(20)"
                ))
                conn.commit()
        if 'последовательность_слалом' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN последовательность_слалом VARCHAR(100)"
                ))
                conn.commit()
        if 'вывод_старт_слалом' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN вывод_старт_слалом VARCHAR(20)"
                ))
                conn.commit()
        if 'связки_ворот_слалом' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN связки_ворот_слалом TEXT"
                ))
                conn.commit()
        for col in ('старший_судья_старта_сл', 'старший_судья_финиша_сл'):
            if col not in cols_сор:
                with db.engine.connect() as conn:
                    conn.execute(db.text(f'ALTER TABLE соревнования ADD COLUMN {col} VARCHAR(200)'))
                    conn.commit()
        if 'последовательность_дг' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN последовательность_дг VARCHAR(100)"
                ))
                conn.commit()
        if 'вывод_старт_дг' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN вывод_старт_дг VARCHAR(20)"
                ))
                conn.commit()
        if 'количество_секций_дг_ход' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN количество_секций_дг_ход INTEGER"
                ))
                conn.commit()
        if 'судьи_секций_дг' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(db.text(
                    "ALTER TABLE соревнования ADD COLUMN судьи_секций_дг VARCHAR(500)"
                ))
                conn.commit()
        # результаты_дг: штраф +10с (новый столбец)
        if inspector.has_table('результаты_дг'):
            cols_дг = {c['name'] for c in inspector.get_columns('результаты_дг')}
            if 'кол_штраф_с10' not in cols_дг:
                conn.execute(text('ALTER TABLE результаты_дг ADD COLUMN кол_штраф_с10 INTEGER DEFAULT 0'))
                conn.commit()
        # результаты_пс: время старта заезда
        if inspector.has_table('результаты_пс'):
            cols_пс = {c['name'] for c in inspector.get_columns('результаты_пс')}
            if 'время_старта' not in cols_пс:
                with db.engine.connect() as conn:
                    conn.execute(db.text(
                        "ALTER TABLE результаты_пс ADD COLUMN время_старта FLOAT"
                    ))
                    conn.commit()
        # результаты_пс: время финиша + раздельные штрафы +10с/+50с (мобильный ввод)
        if inspector.has_table('результаты_пс'):
            cols_пс = {c['name'] for c in inspector.get_columns('результаты_пс')}
            if 'время_финиша' not in cols_пс:
                with db.engine.connect() as conn:
                    conn.execute(db.text(
                        "ALTER TABLE результаты_пс ADD COLUMN время_финиша FLOAT"
                    ))
                    conn.commit()
            if 'штраф_с10' not in cols_пс:
                with db.engine.connect() as conn:
                    conn.execute(db.text(
                        "ALTER TABLE результаты_пс ADD COLUMN штраф_с10 INTEGER DEFAULT 0"
                    ))
                    conn.commit()
            if 'штраф_с50' not in cols_пс:
                with db.engine.connect() as conn:
                    conn.execute(db.text(
                        "ALTER TABLE результаты_пс ADD COLUMN штраф_с50 INTEGER DEFAULT 0"
                    ))
                    conn.commit()
            if 'штраф_ст50' not in cols_пс:
                with db.engine.connect() as conn:
                    conn.execute(db.text(
                        "ALTER TABLE результаты_пс ADD COLUMN штраф_ст50 INTEGER DEFAULT 0"
                    ))
                    conn.commit()
            if 'штраф_буй' not in cols_пс:
                with db.engine.connect() as conn:
                    conn.execute(db.text(
                        "ALTER TABLE результаты_пс ADD COLUMN штраф_буй INTEGER DEFAULT 0"
                    ))
                    conn.commit()
            for _буй_col in ('буй1', 'буй2', 'буй3', 'буй4'):
                if _буй_col not in cols_пс:
                    with db.engine.connect() as conn:
                        conn.execute(db.text(
                            f"ALTER TABLE результаты_пс ADD COLUMN {_буй_col} BOOLEAN"
                        ))
                        conn.commit()
        if 'последовательность_пс' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN последовательность_пс VARCHAR(100)"
                ))
                conn.commit()
        if 'вывод_старт_пс' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN вывод_старт_пс VARCHAR(20)"
                ))
                conn.commit()
        if 'старший_судья_старта_пс' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN старший_судья_старта_пс VARCHAR(200)"
                ))
                conn.commit()
        if 'старший_судья_финиша_пс' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN старший_судья_финиша_пс VARCHAR(200)"
                ))
                conn.commit()
        if 'количество_секций_пс' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN количество_секций_пс INTEGER"
                ))
                conn.commit()
        if 'судьи_секций_пс' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN судьи_секций_пс VARCHAR(500)"
                ))
                conn.commit()
        for _col in ('судья_буя_1_пс', 'судья_буя_2_пс', 'судья_буя_3_пс', 'судья_буя_4_пс'):
            if _col not in cols_сор:
                with db.engine.connect() as conn:
                    conn.execute(text(f"ALTER TABLE соревнования ADD COLUMN {_col} VARCHAR(200)"))
                    conn.commit()
        if 'последовательность_допуска' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN последовательность_допуска VARCHAR(100)"
                ))
                conn.commit()
        if 'вывод_допуска' not in cols_сор:
            with db.engine.connect() as conn:
                conn.execute(text(
                    "ALTER TABLE соревнования ADD COLUMN вывод_допуска VARCHAR(20)"
                ))
                conn.commit()
        # предварительные_заявки: должность руководителя
        if inspector.has_table('предварительные_заявки'):
            cols_з = {c['name'] for c in inspector.get_columns('предварительные_заявки')}
            if 'орг_должность_руководителя' not in cols_з:
                conn.execute(text(
                    'ALTER TABLE предварительные_заявки ADD COLUMN орг_должность_руководителя VARCHAR(200)'
                ))
                conn.commit()
        # соревнования: готовность и текстовые поля отчётов
        cols_сор3 = {c['name'] for c in inspector.get_columns('соревнования')}
        if 'готов_отчеты' not in cols_сор3:
            conn.execute(text("ALTER TABLE соревнования ADD COLUMN готов_отчеты BOOLEAN NOT NULL DEFAULT 0"))
            conn.commit()
        for col, ddl in [
            ('отч_объект_спорта',   'TEXT'),
            ('отч_работа_комиссии', 'TEXT'),
            ('отч_тех_совещания',   'TEXT'),
            ('отч_открытие',        'TEXT'),
            ('отч_информационное',  'TEXT'),
            ('отч_зрители',         'VARCHAR(200)'),
            ('отч_оценка_качества', 'TEXT'),
            ('отч_медицинское',     'TEXT'),
            ('отч_безопасность',    'TEXT'),
            ('отч_замечания',       'TEXT'),
            ('отч_характеристика',  'TEXT'),
            ('отч_номер_сметы',     'VARCHAR(100)'),
        ]:
            if col not in cols_сор3:
                conn.execute(text(f'ALTER TABLE соревнования ADD COLUMN {col} {ddl}'))
                conn.commit()
        # результаты_слалом: штрафы старта и финиша
        if inspector.has_table('результаты_слалом'):
            cols_сл = {c['name'] for c in inspector.get_columns('результаты_слалом')}
            for col in ('кол_штраф_с10', 'кол_штраф_старт', 'кол_штраф_финиш'):
                if col not in cols_сл:
                    conn.execute(text(
                        f'ALTER TABLE результаты_слалом ADD COLUMN {col} INTEGER DEFAULT 0'
                    ))
                    conn.commit()
        # команды: заменить UNIQUE(соревнование_id, нагрудный_номер)
        #          на     UNIQUE(соревнование_id, класс_судна, пол, нагрудный_номер)
        uq_indexes = {idx['name'] for idx in inspector.get_unique_constraints('команды')}
        if 'uq_сор_номер' in uq_indexes:
            # Проверяем, что старый constraint ещё двухколоночный
            old_cols = next(
                (idx['column_names'] for idx in inspector.get_unique_constraints('команды')
                 if idx['name'] == 'uq_сор_номер'),
                []
            )
            if 'класс_судна' not in old_cols:
                with db.engine.connect() as conn2:
                    conn2.execute(text('PRAGMA foreign_keys = OFF'))
                    conn2.execute(text('''
                        CREATE TABLE команды_new (
                            id INTEGER NOT NULL PRIMARY KEY,
                            соревнование_id INTEGER NOT NULL REFERENCES соревнования(id),
                            субъект_id INTEGER REFERENCES субъекты(id),
                            нагрудный_номер INTEGER NOT NULL,
                            название VARCHAR(150),
                            класс_судна VARCHAR(5) NOT NULL DEFAULT "R4",
                            пол VARCHAR(5) NOT NULL DEFAULT "м",
                            возрастная_группа VARCHAR(20),
                            участие_спринт BOOLEAN NOT NULL DEFAULT 1,
                            участие_пс BOOLEAN NOT NULL DEFAULT 1,
                            участие_слалом BOOLEAN NOT NULL DEFAULT 1,
                            участие_дг BOOLEAN NOT NULL DEFAULT 1,
                            участие_многоборье BOOLEAN NOT NULL DEFAULT 1,
                            принадлежность_к_орг VARCHAR(100),
                            оплата_взноса VARCHAR(50) DEFAULT "Ожидает оплаты",
                            стартовый_номер_спринт INTEGER,
                            время_старта_спринт_план VARCHAR(20),
                            UNIQUE (соревнование_id, класс_судна, пол, нагрудный_номер)
                        )
                    '''))
                    conn2.execute(text('''
                        INSERT INTO команды_new SELECT
                            id, соревнование_id, субъект_id, нагрудный_номер, название,
                            класс_судна, пол, возрастная_группа,
                            участие_спринт, участие_пс, участие_слалом, участие_дг, участие_многоборье,
                            принадлежность_к_орг, оплата_взноса,
                            стартовый_номер_спринт, время_старта_спринт_план
                        FROM команды
                    '''))
                    conn2.execute(text('DROP TABLE команды'))
                    conn2.execute(text('ALTER TABLE команды_new RENAME TO команды'))
                    conn2.execute(text('PRAGMA foreign_keys = ON'))
                    conn2.commit()
        # награждение_параметры: размер_шрифта_текста
        if inspector.has_table('награждение_параметры'):
            cols_нп = {c['name'] for c in inspector.get_columns('награждение_параметры')}
            if 'размер_шрифта_текста' not in cols_нп:
                conn.execute(text(
                    'ALTER TABLE награждение_параметры ADD COLUMN размер_шрифта_текста INTEGER NOT NULL DEFAULT 14'
                ))
                conn.commit()
        cols_сор_fin = {c['name'] for c in inspector.get_columns('соревнования')}
        for col in ('старший_судья_старта', 'старший_судья_финиша'):
            if col not in cols_сор_fin:
                with db.engine.connect() as conn:
                    conn.execute(text(f'ALTER TABLE соревнования ADD COLUMN {col} VARCHAR(200)'))
                    conn.commit()
        if 'логотип' not in cols_сор_fin:
            with db.engine.connect() as conn:
                conn.execute(text('ALTER TABLE соревнования ADD COLUMN логотип VARCHAR(255)'))
                conn.commit()


# Справочник: субъекты РФ для автодополнения
СУБЪЕКТЫ_РФ = [
    "Республика Адыгея", "Республика Алтай", "Республика Башкортостан",
    "Республика Бурятия", "Республика Дагестан", "Республика Ингушетия",
    "Кабардино-Балкарская Республика", "Республика Калмыкия",
    "Карачаево-Черкесская Республика", "Республика Карелия",
    "Республика Коми", "Республика Крым", "Республика Марий Эл",
    "Республика Мордовия", "Республика Саха (Якутия)",
    "Республика Северная Осетия — Алания", "Республика Татарстан",
    "Республика Тыва", "Удмуртская Республика", "Республика Хакасия",
    "Чеченская Республика", "Чувашская Республика",
    "Алтайский край", "Забайкальский край", "Камчатский край",
    "Краснодарский край", "Красноярский край", "Пермский край",
    "Приморский край", "Ставропольский край", "Хабаровский край",
    "Амурская область", "Архангельская область", "Астраханская область",
    "Белгородская область", "Брянская область", "Владимирская область",
    "Волгоградская область", "Вологодская область", "Воронежская область",
    "Ивановская область", "Иркутская область", "Калининградская область",
    "Калужская область", "Кемеровская область — Кузбасс",
    "Кировская область", "Костромская область", "Курганская область",
    "Курская область", "Ленинградская область", "Липецкая область",
    "Магаданская область", "Московская область", "Мурманская область",
    "Нижегородская область", "Новгородская область", "Новосибирская область",
    "Омская область", "Оренбургская область", "Орловская область",
    "Пензенская область", "Псковская область", "Ростовская область",
    "Рязанская область", "Самарская область", "Саратовская область",
    "Сахалинская область", "Свердловская область", "Смоленская область",
    "Тамбовская область", "Тверская область", "Томская область",
    "Тульская область", "Тюменская область", "Ульяновская область",
    "Челябинская область", "Ярославская область",
    "город Москва", "город Санкт-Петербург", "город Севастополь",
    "Еврейская автономная область",
    "Ненецкий автономный округ", "Ханты-Мансийский автономный округ — Югра",
    "Чукотский автономный округ", "Ямало-Ненецкий автономный округ",
]