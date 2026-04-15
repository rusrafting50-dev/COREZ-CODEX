# app.py — главный Flask-сервер системы учёта соревнований по рафтингу
import os
from flask import Flask, redirect, url_for
from models import db, заполнить_таблицу_очков, применить_миграции

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

    # ------------------------------------------------------------------
    # База данных
    # ------------------------------------------------------------------
    db.init_app(app)

    with app.app_context():
        db.create_all()
        применить_миграции()
        заполнить_таблицу_очков()

    # ------------------------------------------------------------------
    # Маршруты (blueprint-ы)
    # ------------------------------------------------------------------
    from routes.main        import bp_main
    from routes.соревнование import bp_соревнование
    from routes.субъекты    import bp_субъекты
    from routes.команды     import bp_команды
    from routes.спортсмены  import bp_спортсмены
    from routes.судьи       import bp_судьи

    app.register_blueprint(bp_main)
    app.register_blueprint(bp_соревнование, url_prefix='/соревнование')
    app.register_blueprint(bp_субъекты,     url_prefix='/субъекты')
    app.register_blueprint(bp_команды,      url_prefix='/команды')
    app.register_blueprint(bp_спортсмены,   url_prefix='/спортсмены')
    app.register_blueprint(bp_судьи,        url_prefix='/судьи')

    return app


app = create_app()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
