from datetime import datetime
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from app.extensions import db, login_manager


class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(255), nullable=False, unique=True, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    plan_id = db.Column(db.Integer, db.ForeignKey('plans.id'), nullable=True)
    ppt_count_this_month = db.Column(db.Integer, default=0)
    ppt_count_reset_date = db.Column(db.DateTime, default=datetime.utcnow)
    is_active = db.Column(db.Boolean, default=True)
    is_admin = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime, nullable=True)

    plan = db.relationship('Plan', back_populates='users')
    presentations = db.relationship('Presentation', back_populates='user', lazy='dynamic', cascade='all, delete-orphan')
    ai_settings = db.relationship('AISetting', back_populates='user', lazy='dynamic', cascade='all, delete-orphan')
    activity_logs = db.relationship('ActivityLog', back_populates='user', lazy='dynamic', cascade='all, delete-orphan')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def get_active_ai_setting(self):
        return self.ai_settings.filter_by(is_active=True).first()

    def can_generate_ppt(self):
        from calendar import monthrange
        from datetime import date
        now = datetime.utcnow()
        reset = self.ppt_count_reset_date
        if reset is None or (now.year != reset.year or now.month != reset.month):
            self.ppt_count_this_month = 0
            self.ppt_count_reset_date = now
            db.session.commit()
        if self.plan and self.plan.ppts_per_month == -1:
            return True
        limit = self.plan.ppts_per_month if self.plan else 3
        return self.ppt_count_this_month < limit

    def get_plan_name(self):
        return self.plan.name if self.plan else 'Free'

    def get_ppts_remaining(self):
        if self.plan and self.plan.ppts_per_month == -1:
            return -1
        limit = self.plan.ppts_per_month if self.plan else 3
        return max(0, limit - self.ppt_count_this_month)

    def __repr__(self):
        return f'<User {self.email}>'


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))
