from datetime import datetime
from app.extensions import db


class ActivityLog(db.Model):
    __tablename__ = 'activity_logs'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    action = db.Column(db.String(100), nullable=False)
    details = db.Column(db.Text, nullable=True)
    ip_address = db.Column(db.String(45), nullable=True)
    status = db.Column(db.String(20), default='success')  # success, error, warning
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', back_populates='activity_logs')

    @classmethod
    def log(cls, user_id, action, details=None, ip_address=None, status='success'):
        entry = cls(
            user_id=user_id,
            action=action,
            details=details,
            ip_address=ip_address,
            status=status,
        )
        db.session.add(entry)
        db.session.commit()

    def __repr__(self):
        return f'<ActivityLog {self.action} user={self.user_id}>'
