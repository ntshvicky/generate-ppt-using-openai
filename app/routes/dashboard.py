from flask import Blueprint, render_template
from flask_login import login_required, current_user
from app.models.presentation import Presentation
from app.models.activity_log import ActivityLog

dashboard_bp = Blueprint('dashboard', __name__)


@dashboard_bp.route('/')
@login_required
def index():
    total_ppts = current_user.presentations.count()
    recent_ppts = (current_user.presentations
                   .order_by(Presentation.created_at.desc())
                   .limit(5).all())
    ready_ppts = current_user.presentations.filter_by(status='ready').count()
    recent_logs = (current_user.activity_logs
                   .order_by(ActivityLog.created_at.desc())
                   .limit(6).all())
    ppts_remaining = current_user.get_ppts_remaining()

    return render_template('dashboard/index.html',
                           total_ppts=total_ppts,
                           ready_ppts=ready_ppts,
                           recent_ppts=recent_ppts,
                           recent_logs=recent_logs,
                           ppts_remaining=ppts_remaining)
