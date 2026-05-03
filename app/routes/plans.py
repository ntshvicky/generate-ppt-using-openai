from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required, current_user
from app.extensions import db
from app.models.plan import Plan
from app.models.activity_log import ActivityLog

plans_bp = Blueprint('plans', __name__)


@plans_bp.route('/')
@login_required
def index():
    plans = Plan.query.filter_by(is_active=True).order_by(Plan.price).all()
    return render_template('plans/index.html', plans=plans)


@plans_bp.route('/upgrade/<int:plan_id>', methods=['POST'])
@login_required
def upgrade(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    if current_user.plan_id == plan.id:
        flash(f'You are already on the {plan.name} plan.', 'info')
        return redirect(url_for('plans.index'))

    old_plan = current_user.get_plan_name()
    current_user.plan_id = plan.id
    db.session.commit()
    ActivityLog.log(current_user.id, 'PLAN_UPGRADE',
                    f'Changed from {old_plan} to {plan.name}',
                    ip_address=request.remote_addr)
    flash(f'Successfully upgraded to the {plan.name} plan!', 'success')
    return redirect(url_for('dashboard.index'))
