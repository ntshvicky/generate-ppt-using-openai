from datetime import datetime
from app.extensions import db


class Presentation(db.Model):
    __tablename__ = 'presentations'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    title = db.Column(db.String(255), nullable=False)
    topic = db.Column(db.Text, nullable=True)
    template_id = db.Column(db.String(50), default='corporate_blue')
    num_slides = db.Column(db.Integer, default=8)
    ai_provider = db.Column(db.String(50), nullable=True)
    ai_model = db.Column(db.String(100), nullable=True)
    status = db.Column(db.String(30), default='draft')  # draft, generating, ready, error
    file_path = db.Column(db.String(512), nullable=True)
    video_path = db.Column(db.String(512), nullable=True)
    slides_json = db.Column(db.JSON, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    user = db.relationship('User', back_populates='presentations')
    slides = db.relationship('Slide', back_populates='presentation', lazy='dynamic',
                             cascade='all, delete-orphan', order_by='Slide.slide_number')

    def __repr__(self):
        return f'<Presentation {self.title}>'


class Slide(db.Model):
    __tablename__ = 'slides'

    id = db.Column(db.Integer, primary_key=True)
    presentation_id = db.Column(db.Integer, db.ForeignKey('presentations.id'), nullable=False)
    slide_number = db.Column(db.Integer, nullable=False)
    slide_type = db.Column(db.String(50), default='content')  # title, content, chart, two_col, divider, conclusion
    title = db.Column(db.String(512), nullable=True)
    subtitle = db.Column(db.String(512), nullable=True)
    bullet_points = db.Column(db.JSON, nullable=True)
    include_chart = db.Column(db.Boolean, default=False)
    chart_type = db.Column(db.String(30), nullable=True)
    chart_data = db.Column(db.JSON, nullable=True)
    speaker_notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    presentation = db.relationship('Presentation', back_populates='slides')

    def to_dict(self):
        return {
            'id': self.id,
            'slide_number': self.slide_number,
            'slide_type': self.slide_type,
            'title': self.title,
            'subtitle': self.subtitle,
            'bullet_points': self.bullet_points or [],
            'include_chart': self.include_chart,
            'chart_type': self.chart_type,
            'chart_data': self.chart_data,
            'speaker_notes': self.speaker_notes,
        }

    def __repr__(self):
        return f'<Slide #{self.slide_number} pres={self.presentation_id}>'
