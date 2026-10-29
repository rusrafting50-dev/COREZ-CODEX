# app.py — главный Flask-сервер системы учёта соревнований по рафтингу
import os
from datetime import timedelta
from dotenv import load_dotenv
from flask import Flask, redirect, url_for, session, request
from models import (db, заполнить_таблицу_очков, применить_миграции, заполнить_штрафы,
                    заполнить_коэффициенты_субъектов, получить_настройки)

load_dotenv()

BASE_DIR = os.path.abspath(os.path.dirname(__file__))


def create_app() -> Flask:
    app = Flask(__name__)

    # ------------------------------------------------------------------
    # Конфигурация
    # ------------------------------------------------------------------
    app.config['SQLALCHEMY_DATABASE_URI'] = (
        'sqlite:///' + os.path.join(BASE_DIR, 'local.db')
    )
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-prod')
    app.config['JSON_AS_ASCII'] = False   # корректная отдача кириллицы в JSON
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=7)   # режим киоска на телефоне судьи

    # ------------------------------------------------------------------
    # База данных
    # ------------------------------------------------------------------
    db.init_app(app)

    with app.app_context():
        db.create_all()
        применить_миграции()
        заполнить_таблицу_очков()
        заполнить_штрафы()
        заполнить_коэффициенты_субъектов()

    # ------------------------------------------------------------------
    # Маршруты (blueprint-ы)
    # ------------------------------------------------------------------
    from routes.main        import bp_main
    from routes.соревнование import bp_соревнование
    from routes.субъекты    import bp_субъекты
    from routes.команды     import bp_команды
    from routes.спортсмены  import bp_спортсмены
    from routes.судьи       import bp_судьи
    from routes.документы   import bp_документы
    from routes.спринт      import bp_спринт
    from routes.пс          import bp_пс
    from routes.слалом      import bp_слалом
    from routes.дг          import bp_дг
    from routes.многоборье  import bp_многоборье
    from routes.зачет       import bp_зачет
    from routes.отчеты      import bp_отчеты
    from routes.заявка           import bp_заявка
    from routes.тех_комиссия    import bp_тех
    from routes.дистанция       import bp_дистанция
    from routes.награждение     import bp_награждение
    from routes.api             import bp_api

    app.register_blueprint(bp_main)
    app.register_blueprint(bp_соревнование, url_prefix='/соревнование')
    app.register_blueprint(bp_субъекты,     url_prefix='/субъекты')
    app.register_blueprint(bp_команды,      url_prefix='/команды')
    app.register_blueprint(bp_спортсмены,   url_prefix='/спортсмены')
    app.register_blueprint(bp_судьи,        url_prefix='/судьи')
    app.register_blueprint(bp_документы,    url_prefix='/документы')
    app.register_blueprint(bp_спринт,       url_prefix='/спринт')
    app.register_blueprint(bp_пс,           url_prefix='/пс')
    app.register_blueprint(bp_слалом,       url_prefix='/слалом')
    app.register_blueprint(bp_дг,           url_prefix='/дг')
    app.register_blueprint(bp_многоборье,   url_prefix='/многоборье')
    app.register_blueprint(bp_зачет,        url_prefix='/зачет')
    app.register_blueprint(bp_отчеты,       url_prefix='/отчеты')
    app.register_blueprint(bp_заявка,       url_prefix='/заявка')
    app.register_blueprint(bp_тех,          url_prefix='/тех-комиссия')
    app.register_blueprint(bp_дистанция,    url_prefix='/дистанция')
    app.register_blueprint(bp_награждение,  url_prefix='/награждение')
    app.register_blueprint(bp_api,          url_prefix='/api')

    # ------------------------------------------------------------------
    # Офлайн/онлайн режим — доступен во всех шаблонах как `офлайн_режим`
    # ------------------------------------------------------------------
    @app.context_processor
    def _inject_офлайн_режим():
        return dict(офлайн_режим=получить_настройки().офлайн_режим)

    # ------------------------------------------------------------------
    # Режим киоска: телефон судьи, закреплённый за связкой ворот слалома,
    # не может открыть ничего, кроме своей страницы ввода штрафов.
    # ------------------------------------------------------------------
    _КИОСК_РАЗРЕШЁННЫЕ_ЭНДПОИНТЫ = {
        'static',
        'слалом.ворота_связка',
        'слалом.ворота_связка_штраф',
        'слалом.ворота_связка_выйти',
    }

    @app.before_request
    def _ограничить_киоск():
        замок = session.get('киоск_связка')
        if not замок:
            return None
        if request.endpoint not in _КИОСК_РАЗРЕШЁННЫЕ_ЭНДПОИНТЫ:
            return redirect(url_for('слалом.ворота_связка', sor_id=замок['sor_id'],
                                     attempt=1, group=замок['group']))
        if request.endpoint == 'слалом.ворота_связка':
            if (request.view_args.get('sor_id') != замок['sor_id'] or
                    request.view_args.get('group') != замок['group']):
                return redirect(url_for('слалом.ворота_связка', sor_id=замок['sor_id'],
                                         attempt=1, group=замок['group']))
        elif request.endpoint == 'слалом.ворота_связка_штраф':
            if request.view_args.get('sor_id') != замок['sor_id']:
                return redirect(url_for('слалом.ворота_связка', sor_id=замок['sor_id'],
                                         attempt=1, group=замок['group']))
        return None

    # ------------------------------------------------------------------
    # Заголовки безопасности
    # ------------------------------------------------------------------
    @app.after_request
    def add_csp_header(response):
        response.headers['Content-Security-Policy'] = (
            "default-src 'self'; "
            "script-src 'self' 'unsafe-inline' 'unsafe-eval'; "
            "style-src 'self' 'unsafe-inline'; "
            "img-src 'self' data:; "
            "font-src 'self' data:; "
            "connect-src 'self';"
        )
        # Запрет кэширования скачиваемых файлов (Word/DOCX/PDF)
        if 'attachment' in response.headers.get('Content-Disposition', ''):
            response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'
        return response

    return app


app = create_app()

if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')
