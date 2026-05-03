import os
from flask import Flask
from config import Config
from app.extensions import db, login_manager, csrf


def create_app():
    app = Flask(__name__, template_folder='templates', static_folder='static')
    app.config.from_object(Config)

    os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)

    db.init_app(app)
    login_manager.init_app(app)
    csrf.init_app(app)

    from app.routes.auth import auth_bp
    from app.routes.dashboard import dashboard_bp
    from app.routes.presentations import presentations_bp
    from app.routes.settings import settings_bp
    from app.routes.plans import plans_bp
    from app.routes.logs import logs_bp

    app.register_blueprint(auth_bp, url_prefix='/auth')
    app.register_blueprint(dashboard_bp, url_prefix='/')
    app.register_blueprint(presentations_bp, url_prefix='/presentations')
    app.register_blueprint(settings_bp, url_prefix='/settings')
    app.register_blueprint(plans_bp, url_prefix='/plans')
    app.register_blueprint(logs_bp, url_prefix='/logs')

    return app
