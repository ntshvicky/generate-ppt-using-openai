from datetime import datetime
from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from app.extensions import db
from app.models.user import User
from app.models.plan import Plan
from app.models.activity_log import ActivityLog

auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard.index'))

    error = None
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        remember = bool(request.form.get('remember'))

        user = User.query.filter_by(email=email).first()
        if user and user.check_password(password) and user.is_active:
            login_user(user, remember=remember)
            user.last_login = datetime.utcnow()
            db.session.commit()
            ActivityLog.log(user.id, 'LOGIN', f'Login from {request.remote_addr}',
                            ip_address=request.remote_addr)
            next_page = request.args.get('next')
            return redirect(next_page or url_for('dashboard.index'))
        else:
            error = 'Invalid email or password.'
            if user:
                ActivityLog.log(user.id, 'LOGIN_FAILED', 'Wrong password',
                                ip_address=request.remote_addr, status='error')

    return render_template('auth/login.html', error=error)


@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard.index'))

    error = None
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        confirm = request.form.get('confirm_password', '')

        if not name or not email or not password:
            error = 'All fields are required.'
        elif len(password) < 8:
            error = 'Password must be at least 8 characters.'
        elif password != confirm:
            error = 'Passwords do not match.'
        elif User.query.filter_by(email=email).first():
            error = 'An account with this email already exists.'
        else:
            free_plan = Plan.query.filter_by(slug='free').first()
            user = User(name=name, email=email, plan=free_plan)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            ActivityLog.log(user.id, 'REGISTER', 'New account created',
                            ip_address=request.remote_addr)
            login_user(user)
            flash('Welcome! Your account has been created on the Free plan.', 'success')
            return redirect(url_for('dashboard.index'))

    return render_template('auth/register.html', error=error)


@auth_bp.route('/logout')
@login_required
def logout():
    ActivityLog.log(current_user.id, 'LOGOUT', '', ip_address=request.remote_addr)
    logout_user()
    flash('You have been logged out.', 'info')
    return redirect(url_for('auth.login'))
