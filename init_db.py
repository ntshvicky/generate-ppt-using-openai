"""Run once to create all tables and seed plan data."""
from app import create_app
from app.extensions import db
from app.models.plan import Plan

PLANS = [
    {
        'name': 'Free',
        'slug': 'free',
        'price': 0.00,
        'ppts_per_month': 3,
        'max_slides': 5,
        'features': {
            'templates': 5,
            'video_export': False,
            'priority_ai': False,
            'custom_branding': False,
        },
    },
    {
        'name': 'Pro',
        'slug': 'pro',
        'price': 19.00,
        'ppts_per_month': 20,
        'max_slides': 15,
        'features': {
            'templates': 10,
            'video_export': True,
            'priority_ai': True,
            'custom_branding': False,
        },
    },
    {
        'name': 'Enterprise',
        'slug': 'enterprise',
        'price': 49.00,
        'ppts_per_month': -1,
        'max_slides': 30,
        'features': {
            'templates': 10,
            'video_export': True,
            'priority_ai': True,
            'custom_branding': True,
        },
    },
]


def seed():
    app = create_app()
    with app.app_context():
        db.create_all()
        print('✅ Tables created.')

        for plan_data in PLANS:
            existing = Plan.query.filter_by(slug=plan_data['slug']).first()
            if existing:
                for k, v in plan_data.items():
                    setattr(existing, k, v)
                print(f'   ↻  Updated plan: {plan_data["name"]}')
            else:
                plan = Plan(**plan_data)
                db.session.add(plan)
                print(f'   ✚  Created plan: {plan_data["name"]}')

        db.session.commit()
        print('✅ Plans seeded. Database ready.')


if __name__ == '__main__':
    seed()
