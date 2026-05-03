import os
from datetime import timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'ai-ppt-saas-2024-secret-xK9mP2vQ')
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        'DATABASE_URL',
        'mysql+pymysql://root:Passw0rd@localhost/ai_ppt_maker'
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_POOL_RECYCLE = 3600
    SQLALCHEMY_POOL_TIMEOUT = 20
    PERMANENT_SESSION_LIFETIME = timedelta(days=7)
    WTF_CSRF_ENABLED = True

    GENERATED_FOLDER = os.path.join(BASE_DIR, 'app', 'static', 'generated')
    MAX_CONTENT_LENGTH = 32 * 1024 * 1024

    PLAN_LIMITS = {
        'free':       {'ppts_per_month': 3,   'max_slides': 5},
        'pro':        {'ppts_per_month': 20,  'max_slides': 15},
        'enterprise': {'ppts_per_month': -1,  'max_slides': 30},
    }
