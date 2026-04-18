# models.py — модели базы данных SQLite для системы учёта соревнований по рафтингу
# Язык интерфейса: русский
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

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
    место_проведения            = db.Column(db.String(300), nullable=True)
    код_субъекта_рф             = db.Column(db.Integer,     nullable=True)   # код региона-организатора

    # Спортивные характеристики
    класс_судна                 = db.Column(db.String(5),  nullable=False, default='R4')  # R4 / R6
    возрастная_группа           = db.Column(db.String(50), nullable=True)   # '15 лет и старше'
    статус                      = db.Column(db.String(50), nullable=True)   # Чемпионат, Кубок …
    уровень                     = db.Column(db.String(50), nullable=True)   # России, ФО, области …
    код_соревнований            = db.Column(db.String(20), nullable=True)   # ЧР, КР, ВС …

    # Состав судейской коллегии
    главный_судья               = db.Column(db.String(200), nullable=True)
    главный_секретарь           = db.Column(db.String(200), nullable=True)
    председатель_комиссии_допуска = db.Column(db.String(200), nullable=True)
    начальник_дистанции         = db.Column(db.String(200), nullable=True)
    начальник_спасательной_службы = db.Column(db.String(200), nullable=True)
    зам_главного_секретаря      = db.Column(db.String(200), nullable=True)

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

    создано                     = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)

    # Связи
    субъекты    = db.relationship('Субъект',  back_populates='соревнование', cascade='all, delete-orphan')
    команды     = db.relationship('Команда',  back_populates='соревнование', cascade='all, delete-orphan')
    судьи       = db.relationship('Судья',    back_populates='соревнование', cascade='all, delete-orphan')

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
    представитель       = db.Column(db.String(200), nullable=True)    # ФИО тренера-представителя

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

    нагрудный_номер     = db.Column(db.Integer, nullable=False)   # уникальный в рамках соревнования
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

    # Номер жеребьёвки для спринта (стартовый порядок)
    стартовый_номер_спринт  = db.Column(db.Integer, nullable=True)

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
        db.UniqueConstraint('соревнование_id', 'нагрудный_номер', name='uq_сор_номер'),
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
    # например: «Главный судья», «Главный секретарь», «Судья на старте», …

    # Связи
    соревнование = db.relationship('Соревнование', back_populates='судьи')

    def __repr__(self):
        return f'<Судья "{self.фамилия_имя_отчество}" {self.квалификация}>'


# ---------------------------------------------------------------------------
# Результаты по дистанциям
# ---------------------------------------------------------------------------

# Общие статусы результата
СТАТУС_ФИНИШ   = 'ФИН'   # финишировал
СТАТУС_ДНС     = 'НСТ'   # не стартовал
СТАТУС_ДНФ     = 'НФН'   # не финишировал
СТАТУС_ДК      = 'ДК'    # дисквалификация
СТАТУС_ПМ      = 'ПМ'    # последнее место (штраф без времени)


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
    # '1_этап', '1/16', '1/8', '1/4', '1/2', 'Финал_А', 'Финал_Б'

    номер_заезда    = db.Column(db.Integer, nullable=True)
    дорожка         = db.Column(db.String(10), nullable=True)   # 'левая' / 'правая'

    время           = db.Column(db.Float, nullable=True)
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
    штрафное_время      = db.Column(db.Float,   nullable=False, default=0)   # 5*n5 + 50*n50

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


def применить_миграции():
    """Добавить недостающие колонки в существующую БД (идемпотентно)."""
    from sqlalchemy import inspect, text
    inspector = inspect(db.engine)
    with db.engine.connect() as conn:
        # субъекты: представитель
        cols = {c['name'] for c in inspector.get_columns('субъекты')}
        if 'представитель' not in cols:
            conn.execute(text(
                'ALTER TABLE субъекты ADD COLUMN представитель VARCHAR(200)'
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
        # результаты_спринт: новые колонки штрафов по типу
        if inspector.has_table('результаты_спринт'):
            cols = {c['name'] for c in inspector.get_columns('результаты_спринт')}
            for col in ('штраф_с10', 'штраф_с50', 'штраф_ф50'):
                if col not in cols:
                    conn.execute(text(
                        f'ALTER TABLE результаты_спринт ADD COLUMN {col} INTEGER DEFAULT 0'
                    ))
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
