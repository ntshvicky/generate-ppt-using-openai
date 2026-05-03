import os
import json
from flask import (Blueprint, render_template, redirect, url_for, flash,
                   request, jsonify, send_file, current_app)
from flask_login import login_required, current_user
from app.extensions import db
from app.models.presentation import Presentation, Slide
from app.models.activity_log import ActivityLog
from app.services.template_service import get_all_templates, get_template
from app.services.ai_service import generate_presentation_content
from app.services.ppt_service import build_pptx

presentations_bp = Blueprint('presentations', __name__)

CATEGORIES = [
    'Business', 'Technology', 'Finance', 'Marketing', 'Sales',
    'Healthcare', 'Education', 'Startup', 'Sustainability', 'Research',
    'HR & Recruitment', 'Legal & Compliance', 'Project Management', 'General'
]


@presentations_bp.route('/')
@login_required
def list():
    page = request.args.get('page', 1, type=int)
    search = request.args.get('q', '')
    query = current_user.presentations.order_by(Presentation.created_at.desc())
    if search:
        query = query.filter(Presentation.title.ilike(f'%{search}%'))
    ppts = query.paginate(page=page, per_page=9, error_out=False)
    return render_template('presentations/list.html', ppts=ppts, search=search)


@presentations_bp.route('/create', methods=['GET', 'POST'])
@login_required
def create():
    templates = get_all_templates()
    if request.method == 'POST':
        topic = request.form.get('topic', '').strip()
        template_id = request.form.get('template_id', 'corporate_blue')
        num_slides = int(request.form.get('num_slides', 8))
        category = request.form.get('category', 'Business')

        if not topic:
            flash('Please enter a topic for your presentation.', 'danger')
            return render_template('presentations/create.html',
                                   templates=templates, categories=CATEGORIES)

        if not current_user.can_generate_ppt():
            flash('You have reached your monthly PPT limit. Please upgrade your plan.', 'warning')
            return redirect(url_for('plans.index'))

        plan = current_user.plan
        max_slides = plan.max_slides if plan else 5
        num_slides = min(num_slides, max_slides)

        ai_setting = current_user.get_active_ai_setting()
        if not ai_setting:
            flash('Please configure an AI provider in Settings before generating.', 'warning')
            return redirect(url_for('settings.index'))

        ppt = Presentation(
            user_id=current_user.id,
            title=topic[:255],
            topic=topic,
            template_id=template_id,
            num_slides=num_slides,
            ai_provider=ai_setting.provider,
            ai_model=ai_setting.model,
            status='generating',
        )
        db.session.add(ppt)
        db.session.commit()

        try:
            ai_result = generate_presentation_content(ai_setting, topic, num_slides, category)
            slides_data = ai_result.get('slides', [])
            ppt.title = ai_result.get('presentation_title', topic)[:255]
            ppt.slides_json = slides_data

            for sd in slides_data:
                slide = Slide(
                    presentation_id=ppt.id,
                    slide_number=sd.get('slide_number', 1),
                    slide_type=sd.get('slide_type', 'content'),
                    title=sd.get('title', ''),
                    subtitle=sd.get('subtitle'),
                    bullet_points=sd.get('bullet_points', []),
                    include_chart=sd.get('include_chart', False),
                    chart_type=sd.get('chart_type'),
                    chart_data=sd.get('chart_data'),
                    speaker_notes=sd.get('speaker_notes', ''),
                )
                db.session.add(slide)

            filepath, filename = build_pptx(
                slides_data, template_id, current_app.config['GENERATED_FOLDER']
            )
            ppt.file_path = filename
            ppt.status = 'ready'

            current_user.ppt_count_this_month += 1
            db.session.commit()

            ActivityLog.log(current_user.id, 'PPT_GENERATED',
                            f'Generated "{ppt.title}" ({num_slides} slides, {ai_setting.provider})',
                            ip_address=request.remote_addr)

            flash(f'"{ppt.title}" generated successfully!', 'success')
            return redirect(url_for('presentations.editor', ppt_id=ppt.id))

        except Exception as e:
            ppt.status = 'error'
            db.session.commit()
            ActivityLog.log(current_user.id, 'PPT_GENERATION_ERROR',
                            str(e)[:500], ip_address=request.remote_addr, status='error')
            flash(f'AI generation failed: {str(e)[:200]}', 'danger')
            return redirect(url_for('presentations.list'))

    return render_template('presentations/create.html',
                           templates=templates, categories=CATEGORIES)


@presentations_bp.route('/<int:ppt_id>/editor')
@login_required
def editor(ppt_id):
    ppt = Presentation.query.filter_by(id=ppt_id, user_id=current_user.id).first_or_404()
    slides = ppt.slides.order_by(Slide.slide_number).all()
    template = get_template(ppt.template_id)
    slides_dicts = [s.to_dict() for s in slides]
    return render_template('presentations/editor.html',
                           ppt=ppt,
                           slides=slides_dicts,
                           template=template,
                           templates=get_all_templates())


@presentations_bp.route('/<int:ppt_id>/save', methods=['POST'])
@login_required
def save(ppt_id):
    ppt = Presentation.query.filter_by(id=ppt_id, user_id=current_user.id).first_or_404()
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No data'}), 400

    slides_data = data.get('slides', [])
    template_id = data.get('template_id', ppt.template_id)

    for sd in slides_data:
        slide = Slide.query.filter_by(
            presentation_id=ppt.id, slide_number=sd.get('slide_number')
        ).first()
        if slide:
            slide.title = sd.get('title', slide.title)
            slide.subtitle = sd.get('subtitle', slide.subtitle)
            slide.bullet_points = sd.get('bullet_points', slide.bullet_points)
            slide.include_chart = sd.get('include_chart', slide.include_chart)
            slide.chart_type = sd.get('chart_type', slide.chart_type)
            slide.chart_data = sd.get('chart_data', slide.chart_data)
            slide.speaker_notes = sd.get('speaker_notes', slide.speaker_notes)

    ppt.template_id = template_id
    ppt.slides_json = slides_data

    try:
        filepath, filename = build_pptx(
            slides_data, template_id, current_app.config['GENERATED_FOLDER']
        )
        if ppt.file_path:
            old_path = os.path.join(current_app.config['GENERATED_FOLDER'], ppt.file_path)
            if os.path.exists(old_path):
                os.remove(old_path)
        ppt.file_path = filename
        ppt.status = 'ready'
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    db.session.commit()
    ActivityLog.log(current_user.id, 'PPT_EDITED',
                    f'Edited "{ppt.title}"', ip_address=request.remote_addr)
    return jsonify({'success': True, 'message': 'Presentation saved and rebuilt.'})


@presentations_bp.route('/<int:ppt_id>/download')
@login_required
def download(ppt_id):
    ppt = Presentation.query.filter_by(id=ppt_id, user_id=current_user.id).first_or_404()
    if not ppt.file_path:
        flash('File not found. Please regenerate the presentation.', 'danger')
        return redirect(url_for('presentations.editor', ppt_id=ppt_id))

    filepath = os.path.join(current_app.config['GENERATED_FOLDER'], ppt.file_path)
    if not os.path.exists(filepath):
        flash('File not found on disk. Please save again.', 'danger')
        return redirect(url_for('presentations.editor', ppt_id=ppt_id))

    safe_name = ppt.title[:40].replace(' ', '_').replace('/', '-') + '.pptx'
    ActivityLog.log(current_user.id, 'PPT_EXPORTED',
                    f'Downloaded "{ppt.title}"', ip_address=request.remote_addr)
    return send_file(filepath, as_attachment=True, download_name=safe_name)


@presentations_bp.route('/<int:ppt_id>/export-video', methods=['POST'])
@login_required
def export_video(ppt_id):
    ppt = Presentation.query.filter_by(id=ppt_id, user_id=current_user.id).first_or_404()

    plan = current_user.plan
    if not plan or plan.slug == 'free':
        return jsonify({'error': 'Video export requires a Pro or Enterprise plan.'}), 403

    data = request.get_json() or {}
    slide_images = data.get('slide_images', [])

    if not slide_images:
        return jsonify({'error': 'No slide images provided.'}), 400

    try:
        import base64
        from PIL import Image
        import io
        import imageio

        frames = []
        for img_data in slide_images:
            if ',' in img_data:
                img_data = img_data.split(',')[1]
            img_bytes = base64.b64decode(img_data)
            img = Image.open(io.BytesIO(img_bytes)).convert('RGB')
            img = img.resize((1280, 720), Image.LANCZOS)
            frames.append(img)

        video_filename = f"{ppt.file_path.replace('.pptx', '')}_video.mp4"
        video_path = os.path.join(current_app.config['GENERATED_FOLDER'], video_filename)

        writer = imageio.get_writer(video_path, fps=0.5, codec='libx264',
                                    ffmpeg_params=['-pix_fmt', 'yuv420p'])
        for frame in frames:
            for _ in range(4):  # 4 frames = 2 seconds per slide at 2fps
                writer.append_data(frame)
        writer.close()

        ppt.video_path = video_filename
        db.session.commit()

        ActivityLog.log(current_user.id, 'PPT_EXPORTED',
                        f'Video exported "{ppt.title}"', ip_address=request.remote_addr)

        return jsonify({'success': True, 'video_url': url_for('presentations.download_video',
                                                               ppt_id=ppt.id)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@presentations_bp.route('/<int:ppt_id>/download-video')
@login_required
def download_video(ppt_id):
    ppt = Presentation.query.filter_by(id=ppt_id, user_id=current_user.id).first_or_404()
    if not ppt.video_path:
        flash('Video not found. Please export again.', 'danger')
        return redirect(url_for('presentations.editor', ppt_id=ppt_id))

    filepath = os.path.join(current_app.config['GENERATED_FOLDER'], ppt.video_path)
    if not os.path.exists(filepath):
        flash('Video file not found. Please re-export.', 'danger')
        return redirect(url_for('presentations.editor', ppt_id=ppt_id))

    safe_name = ppt.title[:40].replace(' ', '_') + '.mp4'
    return send_file(filepath, as_attachment=True, download_name=safe_name)


@presentations_bp.route('/<int:ppt_id>/regenerate', methods=['POST'])
@login_required
def regenerate(ppt_id):
    """Re-run AI generation on the same topic/template, replacing slides."""
    ppt = Presentation.query.filter_by(id=ppt_id, user_id=current_user.id).first_or_404()

    ai_setting = current_user.get_active_ai_setting()
    if not ai_setting:
        flash('No active AI provider configured. Add one in Settings.', 'warning')
        return redirect(url_for('settings.index'))

    if not current_user.can_generate_ppt():
        flash('Monthly PPT limit reached. Please upgrade your plan.', 'warning')
        return redirect(url_for('plans.index'))

    # Allow changing template/num_slides from the form if provided
    template_id  = request.form.get('template_id', ppt.template_id)
    num_slides   = int(request.form.get('num_slides', ppt.num_slides) or ppt.num_slides)
    category     = request.form.get('category', 'Business')
    topic        = request.form.get('topic', ppt.topic or ppt.title)

    plan = current_user.plan
    max_slides = plan.max_slides if plan else 5
    num_slides = min(num_slides, max_slides)

    ppt.status = 'generating'
    ppt.ai_provider = ai_setting.provider
    ppt.ai_model = ai_setting.model
    ppt.template_id = template_id
    db.session.commit()

    try:
        ai_result   = generate_presentation_content(ai_setting, topic, num_slides, category)
        slides_data = ai_result.get('slides', [])

        # Replace title if the AI returned a better one
        new_title = ai_result.get('presentation_title', '').strip()
        if new_title:
            ppt.title = new_title[:255]

        ppt.slides_json = slides_data

        # Delete old slides and write fresh ones
        Slide.query.filter_by(presentation_id=ppt.id).delete()
        for sd in slides_data:
            slide = Slide(
                presentation_id=ppt.id,
                slide_number=sd.get('slide_number', 1),
                slide_type=sd.get('slide_type', 'content'),
                title=sd.get('title', ''),
                subtitle=sd.get('subtitle'),
                bullet_points=sd.get('bullet_points', []),
                include_chart=sd.get('include_chart', False),
                chart_type=sd.get('chart_type'),
                chart_data=sd.get('chart_data'),
                speaker_notes=sd.get('speaker_notes', ''),
            )
            db.session.add(slide)

        # Remove old PPTX file
        if ppt.file_path:
            old_path = os.path.join(current_app.config['GENERATED_FOLDER'], ppt.file_path)
            if os.path.exists(old_path):
                os.remove(old_path)

        filepath, filename = build_pptx(slides_data, template_id, current_app.config['GENERATED_FOLDER'])
        ppt.file_path = filename
        ppt.status    = 'ready'
        current_user.ppt_count_this_month += 1
        db.session.commit()

        ActivityLog.log(current_user.id, 'PPT_REGENERATED',
                        f'Regenerated "{ppt.title}" ({num_slides} slides, {ai_setting.provider})',
                        ip_address=request.remote_addr)

        flash(f'"{ppt.title}" regenerated successfully!', 'success')
        return redirect(url_for('presentations.editor', ppt_id=ppt.id))

    except Exception as e:
        ppt.status = 'error'
        db.session.commit()
        ActivityLog.log(current_user.id, 'PPT_REGENERATION_ERROR',
                        str(e)[:500], ip_address=request.remote_addr, status='error')
        flash(f'Regeneration failed: {str(e)[:300]}', 'danger')
        return redirect(url_for('presentations.editor', ppt_id=ppt.id))


@presentations_bp.route('/<int:ppt_id>/delete', methods=['POST'])
@login_required
def delete(ppt_id):
    ppt = Presentation.query.filter_by(id=ppt_id, user_id=current_user.id).first_or_404()
    title = ppt.title
    if ppt.file_path:
        fp = os.path.join(current_app.config['GENERATED_FOLDER'], ppt.file_path)
        if os.path.exists(fp):
            os.remove(fp)
    db.session.delete(ppt)
    db.session.commit()
    ActivityLog.log(current_user.id, 'PPT_DELETED',
                    f'Deleted "{title}"', ip_address=request.remote_addr)
    flash(f'"{title}" deleted.', 'info')
    return redirect(url_for('presentations.list'))
