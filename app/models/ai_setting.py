from datetime import datetime
from app.extensions import db

PROVIDERS = ['openai', 'anthropic', 'gemini']

PROVIDER_MODELS = {
    'openai': [
        'gpt-4.1',          # Latest flagship — best quality
        'gpt-4.1-mini',     # Fast + affordable
        'gpt-4.1-nano',     # Lightest/cheapest
        'gpt-4o',           # Previous flagship (stable)
        'gpt-4o-mini',      # Previous mini
        'o4-mini',          # Reasoning model (mini)
        'o3',               # Reasoning model (full)
    ],
    'anthropic': [
        'claude-opus-4-7',           # Most capable (Opus 4.7)
        'claude-sonnet-4-6',         # Balanced speed+quality (Sonnet 4.6)
        'claude-haiku-4-5-20251001', # Fastest / cheapest (Haiku 4.5)
    ],
    'gemini': [
        'gemini-2.5-pro',    # Most capable (latest)
        'gemini-2.5-flash',  # Fast + capable
        'gemini-2.0-flash',  # Stable, widely available
    ],
}


class AISetting(db.Model):
    __tablename__ = 'ai_settings'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    provider = db.Column(db.String(50), nullable=False)
    api_key = db.Column(db.String(512), nullable=False)
    model = db.Column(db.String(100), nullable=False)
    is_active = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    user = db.relationship('User', back_populates='ai_settings')

    def __repr__(self):
        return f'<AISetting {self.provider} user={self.user_id}>'
