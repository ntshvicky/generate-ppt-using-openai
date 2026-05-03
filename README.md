# SlideAI — AI-Powered Presentation Maker

<p align="center">
  <img src="screenshots/01_dashboard.png" alt="SlideAI Dashboard" width="100%"/>
</p>

> **SlideAI** is a full-stack SaaS web application that generates professional, download-ready PowerPoint presentations from a single text prompt — powered by OpenAI, Anthropic Claude, or Google Gemini.

---

## ✨ Features

| Category | What you get |
|---|---|
| **AI Generation** | One-click PPT creation from any topic using OpenAI, Anthropic, or Gemini |
| **Multi-provider** | Bring your own API key for GPT-4.1, Claude Sonnet, or Gemini Pro — switch any time |
| **10 Pro Templates** | Corporate Blue, Dark Executive, Tech Purple, Navy Pro, Ocean Teal, and more |
| **Industry-grade PPTX** | Entrance animations, slide transitions, per-bar chart colors, speaker notes |
| **Visual Editor** | Edit titles, bullets, speaker notes, swap templates — all in-browser |
| **One-click Download** | Export `.pptx` ready for PowerPoint, Keynote, or Google Slides |
| **Video Export** | Convert slide deck to `.mp4` slideshow (Pro/Enterprise) |
| **Slide Types** | Title · Content · Chart · Two-column · Divider · Conclusion |
| **Subscription Plans** | Free / Pro ($19) / Enterprise ($49) with per-plan slide limits |
| **Activity Logs** | Full audit trail of every generation, edit, download, and login |
| **Search & Paginate** | Full-text search across all your presentations |

---

## 📸 Screenshots

### Dashboard
<img src="screenshots/01_dashboard.png" alt="Dashboard" width="100%"/>

### Create Presentation
<img src="screenshots/02_create.png" alt="Create Presentation" width="100%"/>

### My Presentations
<img src="screenshots/03_my_presentations.png" alt="My Presentations" width="100%"/>

### Slide Editor
<img src="screenshots/07_editor.png" alt="Slide Editor" width="100%"/>

### AI Provider Settings
<img src="screenshots/04_ai_settings.png" alt="AI Settings" width="100%"/>

### Plans & Billing
<img src="screenshots/05_plans.png" alt="Plans and Billing" width="100%"/>

### Activity Logs
<img src="screenshots/06_activity_logs.png" alt="Activity Logs" width="100%"/>

### Login
<img src="screenshots/00_login.png" alt="Login" width="100%"/>

---

## 🏗️ Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Python 3.11 · Flask 3.0 · Flask-Login · Flask-WTF |
| **Database** | MySQL 8 · SQLAlchemy ORM · Flask-Migrate (Alembic) |
| **AI Providers** | OpenAI SDK · Anthropic SDK · Google Generative AI |
| **PPT Engine** | python-pptx with custom XML animations & transitions |
| **Video Export** | imageio · Pillow · ffmpeg (libx264) |
| **Frontend** | Jinja2 · Tailwind CSS (CDN) · Vanilla JS |
| **Auth** | Session-based · Werkzeug scrypt password hashing · CSRF protection |

---

## 🚀 Quick Start

### 1. Clone & install

```bash
git clone https://github.com/ntshvicky/generate-ppt-using-openai.git
cd generate-ppt-using-openai
python -m venv env
source env/bin/activate          # Windows: env\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure environment

Create a `.env` file in the project root:

```env
SECRET_KEY=your-secret-key-here
DATABASE_URL=mysql+pymysql://root:yourpassword@localhost/ai_ppt_maker
GENERATED_FOLDER=generated_ppts
```

### 3. Set up the database

```bash
# Create the MySQL database
mysql -u root -p -e "CREATE DATABASE ai_ppt_maker CHARACTER SET utf8mb4;"

# Initialise tables and seed default plans
python init_db.py
```

### 4. Run the server

```bash
python run.py
```

Open **http://localhost:5001** in your browser, register an account, add your AI provider API key under **AI Settings**, and generate your first deck.

---

## 🎨 Slide Templates

| Template | Best for | Primary color |
|---|---|---|
| Corporate Blue | Fortune 500 / Business | `#1565C0` |
| Dark Executive | Boardroom / Executive | `#212121` + gold |
| Tech Purple | SaaS / AI pitches | `#4A148C` + cyan |
| Nature Fresh | Sustainability / Wellness | `#2E7D32` |
| Creative Orange | Startup / Product launch | `#E65100` |
| Minimal White | Design / Apple-style | `#212121` |
| Ocean Teal | Healthcare / Research | `#00695C` |
| Sunset Warm | Marketing / Storytelling | `#BF360C` |
| Bold Red | Sales / CTA | `#B71C1C` |
| Navy Professional | Finance / Legal / Consulting | `#1A237E` |

---

## 🤖 AI Providers

| Provider | Recommended model | Notes |
|---|---|---|
| **OpenAI** | `gpt-4.1` · `gpt-4o` · `o4-mini` | Best structured JSON output |
| **Anthropic** | `claude-opus-4-5` · `claude-sonnet-4-5` · `claude-haiku-4-5` | Great for analytical decks |
| **Google Gemini** | `gemini-2.0-flash` · `gemini-2.5-pro` | Fast, multilingual, creative |

Bring your own key — keys are stored encrypted per user and never shared.

---

## 📊 Subscription Plans

| | Free | Pro | Enterprise |
|---|---|---|---|
| Presentations / month | 3 | 20 | Unlimited |
| Max slides per deck | 5 | 15 | 30 |
| Templates | 5 | All 10 | All 10 |
| Video export (.mp4) | ❌ | ✅ | ✅ |
| Priority AI generation | ❌ | ✅ | ✅ |
| Custom branding support | ❌ | ❌ | ✅ |
| Price | $0 | $19/mo | $49/mo |

---

## 🧠 How the PPT Engine Works

Each generated deck is built entirely in Python with **python-pptx** — no external design tools required. Highlights:

- **Staggered entrance animations** — shapes fade in with per-element delays (XML `<p:timing>`)
- **Varied slide transitions** — Fade · Push · Wipe · Zoom · Split, matched to slide type
- **Per-bar chart coloring** — `<c:dPt>` point elements give each bar its own vivid color from an 8-color palette
- **Professional data labels** — bold white labels on every bar/column
- **Insight panels** — chart slides get a right-side "KEY INSIGHTS" card panel
- **Numbered bullet badges** — circular badge with sequential number for each bullet
- **Conclusion cards** — takeaway slides use alternating-color cards with accent bars

---

## 📁 Project Structure

```
ai-ppt-maker/
├── app/
│   ├── models/              # SQLAlchemy models (User, Presentation, Slide, Plan …)
│   ├── routes/              # Flask blueprints (auth, dashboard, presentations, settings …)
│   ├── services/
│   │   ├── ai_service.py    # OpenAI / Anthropic / Gemini generation
│   │   ├── ppt_service.py   # Industry-grade PPTX builder with animations
│   │   └── template_service.py  # 10 color themes
│   ├── templates/           # Jinja2 HTML templates (dark Tailwind UI)
│   └── extensions.py        # DB, login manager, CSRF
├── config.py                # App configuration
├── init_db.py               # DB init + plan seeding
├── run.py                   # Entry point
├── requirements.txt
└── screenshots/             # UI screenshots used in this README
```

---

## 🔒 Security

- CSRF protection on all POST forms (Flask-WTF)
- Passwords hashed with Werkzeug `scrypt` (32768 rounds)
- API keys masked in the UI (`sk-...••••••••`)
- Session-based auth with `login_required` on every protected route
- Input sanitisation and parameterised queries via SQLAlchemy ORM

---

## 📄 License

MIT — free to use, modify, and distribute.

---

<p align="center">Built with ❤️ using Flask · python-pptx · OpenAI · Anthropic · Gemini</p>
