import os
from datetime import timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Detect Vercel serverless environment
_ON_VERCEL = bool(os.environ.get('VERCEL'))


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'ai-ppt-saas-2024-secret-xK9mP2vQ')

    # On Vercel use DATABASE_URL env var (set in Vercel dashboard).
    # Locally falls back to the dev MySQL instance.
    SQLALCHEMY_DATABASE_URI = os.environ.get(
        'DATABASE_URL',
        'mysql+pymysql://root:Passw0rd@localhost/ai_ppt_maker'
    )
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # Serverless: disable connection pool (each invocation is a fresh process).
    # Locally: keep the normal pool for performance.
    if _ON_VERCEL:
        SQLALCHEMY_ENGINE_OPTIONS = {
            'pool_pre_ping': True,
            'poolclass': __import__('sqlalchemy.pool', fromlist=['NullPool']).NullPool,
        }
    else:
        SQLALCHEMY_ENGINE_OPTIONS = {
            'pool_pre_ping': True,
            'pool_recycle': 3600,
            'pool_timeout': 20,
        }

    PERMANENT_SESSION_LIFETIME = timedelta(days=7)
    WTF_CSRF_ENABLED = True

    # On Vercel, /tmp is the only writable directory (ephemeral per invocation).
    # The download route regenerates the PPTX from slides_json if the file is gone.
    GENERATED_FOLDER = (
        '/tmp/generated_ppts'
        if _ON_VERCEL
        else os.environ.get(
            'GENERATED_FOLDER',
            os.path.join(BASE_DIR, 'app', 'static', 'generated')
        )
    )

    MAX_CONTENT_LENGTH = 32 * 1024 * 1024

    PLAN_LIMITS = {
        'free':       {'ppts_per_month': 3,   'max_slides': 5},
        'pro':        {'ppts_per_month': 20,  'max_slides': 15},
        'enterprise': {'ppts_per_month': -1,  'max_slides': 30},
    }
