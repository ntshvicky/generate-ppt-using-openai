from datetime import datetime
from app.extensions import db


class Plan(db.Model):
    __tablename__ = 'plans'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False, unique=True)
    slug = db.Column(db.String(50), nullable=False, unique=True)
    price = db.Column(db.Numeric(10, 2), nullable=False, default=0.00)
    ppts_per_month = db.Column(db.Integer, nullable=False, default=3)
    max_slides = db.Column(db.Integer, nullable=False, default=5)
    features = db.Column(db.JSON, nullable=True)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    users = db.relationship('User', back_populates='plan', lazy='dynamic')

    def __repr__(self):
        return f'<Plan {self.name}>'
