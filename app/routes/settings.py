from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from app.extensions import db
from app.models.ai_setting import AISetting, PROVIDERS, PROVIDER_MODELS
from app.models.activity_log import ActivityLog

settings_bp = Blueprint('settings', __name__)


@settings_bp.route('/')
@login_required
def index():
    ai_settings = current_user.ai_settings.order_by(AISetting.created_at.desc()).all()
    return render_template('settings/index.html',
                           ai_settings=ai_settings,
                           providers=PROVIDERS,
                           provider_models=PROVIDER_MODELS)


@settings_bp.route('/ai/add', methods=['POST'])
@login_required
def add_ai_setting():
    provider = request.form.get('provider')
    api_key = request.form.get('api_key', '').strip()
    model = request.form.get('model', '').strip()

    if not provider or not api_key or not model:
        flash('All fields are required.', 'danger')
        return redirect(url_for('settings.index'))

    if provider not in PROVIDERS:
        flash('Invalid AI provider.', 'danger')
        return redirect(url_for('settings.index'))

    # Deactivate existing settings for same provider
    existing = current_user.ai_settings.filter_by(provider=provider).first()
    if existing:
        existing.api_key = api_key
        existing.model = model
        existing.is_active = True
        db.session.commit()
        flash(f'{provider.title()} settings updated and set as active.', 'success')
    else:
        # Deactivate others if this is the first
        has_active = current_user.ai_settings.filter_by(is_active=True).first()
        setting = AISetting(
            user_id=current_user.id,
            provider=provider,
            api_key=api_key,
            model=model,
            is_active=(not has_active),
        )
        db.session.add(setting)
        db.session.commit()
        flash(f'{provider.title()} API key saved successfully.', 'success')

    ActivityLog.log(current_user.id, 'AI_SETTINGS_UPDATE',
                    f'Updated {provider} settings', ip_address=request.remote_addr)
    return redirect(url_for('settings.index'))


@settings_bp.route('/ai/activate/<int:setting_id>', methods=['POST'])
@login_required
def activate_ai_setting(setting_id):
    setting = AISetting.query.filter_by(id=setting_id, user_id=current_user.id).first_or_404()
    # Deactivate all
    current_user.ai_settings.update({'is_active': False})
    setting.is_active = True
    db.session.commit()
    flash(f'{setting.provider.title()} set as the active AI provider.', 'success')
    ActivityLog.log(current_user.id, 'AI_PROVIDER_CHANGED',
                    f'Active provider: {setting.provider}', ip_address=request.remote_addr)
    return redirect(url_for('settings.index'))


@settings_bp.route('/ai/delete/<int:setting_id>', methods=['POST'])
@login_required
def delete_ai_setting(setting_id):
    setting = AISetting.query.filter_by(id=setting_id, user_id=current_user.id).first_or_404()
    provider = setting.provider
    db.session.delete(setting)
    db.session.commit()
    flash(f'{provider.title()} API key removed.', 'info')
    return redirect(url_for('settings.index'))


@settings_bp.route('/api/models/<provider>')
@login_required
def get_models(provider):
    models = PROVIDER_MODELS.get(provider, [])
    return jsonify({'models': models})
