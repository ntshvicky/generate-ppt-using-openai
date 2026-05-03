from flask import Blueprint, render_template, request
from flask_login import login_required, current_user
from app.models.activity_log import ActivityLog

logs_bp = Blueprint('logs', __name__)


@logs_bp.route('/')
@login_required
def index():
    page = request.args.get('page', 1, type=int)
    status_filter = request.args.get('status', '')
    action_filter = request.args.get('action', '')

    query = current_user.activity_logs.order_by(ActivityLog.created_at.desc())

    if status_filter:
        query = query.filter(ActivityLog.status == status_filter)
    if action_filter:
        query = query.filter(ActivityLog.action.like(f'%{action_filter}%'))

    logs = query.paginate(page=page, per_page=20, error_out=False)

    action_types = ['LOGIN', 'LOGOUT', 'REGISTER', 'PPT_GENERATED', 'PPT_EXPORTED',
                    'PPT_DELETED', 'AI_SETTINGS_UPDATE', 'PLAN_UPGRADE']

    return render_template('logs/index.html',
                           logs=logs,
                           action_types=action_types,
                           status_filter=status_filter,
                           action_filter=action_filter)
