import json
import re

# Models that support response_format=json_object (OpenAI only, not o-series reasoning models)
_OPENAI_JSON_MODE_MODELS = {
    'gpt-4o', 'gpt-4o-mini', 'gpt-4.1', 'gpt-4.1-mini', 'gpt-4.1-nano',
    'gpt-4-turbo', 'gpt-4', 'gpt-3.5-turbo',
}

SLIDE_GENERATION_PROMPT = """You are a McKinsey-level presentation strategist and data storyteller.
Your job is to produce a high-impact, board-ready PowerPoint deck from the given topic.

════════════════════════════════════════
CONTENT QUALITY STANDARDS
════════════════════════════════════════

BULLET POINTS — every bullet must be:
  ✓ Specific and insight-driven (not generic)
  ✓ Include a number, stat, or concrete example where possible
  ✓ Written as a complete thought, 10-20 words
  ✓ Begin with a strong verb or compelling fact
  ✗ NEVER write: "This is important", "Key factor", "Consider the following"

GOOD bullets:  "Global AI market will reach $1.8T by 2030, growing 38% CAGR"
               "72% of enterprises report productivity gains above 30% after AI adoption"
BAD bullets:   "AI is growing rapidly"   |   "Important trends to consider"

TITLES — must be:
  ✓ Insight headline (state the SO WHAT, not just the topic)
  ✓ Max 8 words, no trailing punctuation
  ✓ Action-oriented or claim-based

GOOD titles: "AI Cuts Development Costs by Half"   |   "Three Forces Reshaping Healthcare"
BAD titles:  "AI Overview"   |   "Current Situation"   |   "Introduction"

SPEAKER NOTES — must include:
  ✓ The key message to land with the audience
  ✓ One supporting anecdote, case study, or data point
  ✓ Transition sentence to the next slide
  ✓ 3-5 sentences total

CHART DATA — must be:
  ✓ Realistic numbers grounded in the topic domain
  ✓ At least 4-6 data points (labels + values)
  ✓ Varied enough to show meaningful contrast
  ✓ Properly unitised ($B, %, pts, ×, etc.)

════════════════════════════════════════
NARRATIVE STRUCTURE (follow this arc)
════════════════════════════════════════

Slide 1       → TITLE: Hook the audience with a bold claim or provocative question
Slide 2       → CONTEXT: Set the scene — why this topic matters right now
Slide 3-4     → PROBLEM / OPPORTUNITY: The core challenge or market gap with data
Middle slides → INSIGHT BODY: Evidence, analysis, comparisons, breakdowns
               Use "chart" for data-heavy points, "two_col" for compare/contrast,
               "divider" before major new sections (every 3-4 slides)
Second-to-last → RECOMMENDATIONS / ROADMAP: Clear, actionable steps
Last slide    → CONCLUSION: 3-5 memorable takeaways; call to action

════════════════════════════════════════
SLIDE TYPE RULES
════════════════════════════════════════

"title"      Slide 1 only. Has subtitle. No bullets.
"content"    Standard insight slide. 4-5 rich bullets with data.
"chart"      Use when showing trends, comparisons, rankings, or market size.
             MUST include realistic chart_data. Bullet points = key insights from chart.
"two_col"    Use for compare/contrast: Before vs After, Pros vs Cons, Option A vs B.
             Provide 3 bullets per column (6 bullets total).
"divider"    Section break. Has subtitle. No bullets. Use before major topic shifts.
"conclusion" Final slide. 4-5 strong, memorable takeaways. Clear call to action.

════════════════════════════════════════
STRICT REQUIREMENTS
════════════════════════════════════════

- Produce EXACTLY {num_slides} slides
- Slide 1 must be type "title" with a compelling subtitle
- Last slide must be type "conclusion"
- At least 1 "chart" slide if {num_slides} >= 5
- At least 1 "divider" if {num_slides} >= 7
- At least 1 "two_col" if {num_slides} >= 6
- Every "content" and "conclusion" slide: minimum 4 bullets
- Every "chart" slide: minimum 2 insight bullets + complete chart_data
- chart_type must be one of: "bar", "column", "line", "pie"
- Choose chart_type wisely: bar/column for comparisons, line for trends, pie for composition

════════════════════════════════════════
CATEGORY-SPECIFIC TONE
════════════════════════════════════════

Business / Finance / Consulting → formal, data-heavy, ROI-focused language
Technology / AI / SaaS          → forward-looking, innovation-driven, metric-rich
Healthcare / Research           → evidence-based, patient-outcome focused, cautious claims
Marketing / Sales               → persuasive, benefit-led, customer-centric language
Startup / Pitch                 → bold, vision-driven, market-size framing, investor lens
Sustainability / ESG            → impact-focused, long-term thinking, measurable goals
Education / HR                  → people-centred, skill-building, competency language

════════════════════════════════════════
OUTPUT — return ONLY valid raw JSON, no markdown, no preamble:
════════════════════════════════════════

{{
  "presentation_title": "Bold, Specific Presentation Title (6-10 words)",
  "slides": [
    {{
      "slide_number": 1,
      "slide_type": "title",
      "title": "Punchy 5-7 Word Hook Title",
      "subtitle": "One sentence that frames the narrative and stakes",
      "bullet_points": [],
      "include_chart": false,
      "chart_type": null,
      "chart_data": null,
      "speaker_notes": "Open with a striking fact or question. State what the audience will leave knowing. Transition: 'Let me start by setting the stage...'"
    }},
    {{
      "slide_number": 2,
      "slide_type": "content",
      "title": "Why This Moment Is Different",
      "subtitle": null,
      "bullet_points": [
        "Global market reached $847B in 2024, up 34% from prior year — fastest growth in a decade",
        "Regulatory tailwinds: 60+ countries now mandate adoption, creating a $200B compliance market",
        "Workforce readiness gap: only 12% of employees have skills needed for next-gen operations",
        "Early movers capture 3× more market share than late adopters (McKinsey, 2024)"
      ],
      "include_chart": false,
      "chart_type": null,
      "chart_data": null,
      "speaker_notes": "Set urgency by anchoring to the macro moment. The 34% growth figure is the hook — it signals this is no longer optional. Acknowledge the skills gap as the hidden constraint. Transition: 'So what does the competitive landscape actually look like?'"
    }},
    {{
      "slide_number": 3,
      "slide_type": "chart",
      "title": "Market Leaders Pulling Away Fast",
      "subtitle": null,
      "bullet_points": [
        "Top-quartile companies outperform peers by 4.2× on revenue growth over 3 years",
        "Gap between leaders and laggards widening at 18% per year — window is closing"
      ],
      "include_chart": true,
      "chart_type": "bar",
      "chart_data": {{
        "title": "Revenue Growth Index by Adoption Tier (2021–2024, indexed to 100)",
        "labels": ["Early Adopters", "Fast Followers", "Cautious Adopters", "Laggards"],
        "values": [420, 280, 160, 100],
        "series_name": "Growth Index",
        "unit": "idx"
      }},
      "speaker_notes": "This chart tells the whole story: inaction is not neutral — it is falling behind. The 4.2× gap is the number executives remember. Use this to transition into the specific drivers behind the gap. Transition: 'What are these leaders actually doing differently?'"
    }}
  ]
}}

Topic: {topic}
Number of slides: {num_slides}
Category: {category}
"""


def _extract_json(raw: str) -> dict:
    """Robustly extract a JSON object from an AI response."""
    text = raw.strip()

    # Strip markdown code fences if present
    fence = re.search(r'```(?:json)?\s*([\s\S]+?)\s*```', text)
    if fence:
        text = fence.group(1).strip()

    # Find outermost { ... }
    brace_start = text.find('{')
    brace_end   = text.rfind('}')
    if brace_start >= 0 and brace_end > brace_start:
        text = text[brace_start:brace_end + 1]

    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        # Last-resort: strip control chars and retry
        text_clean = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
        try:
            return json.loads(text_clean)
        except json.JSONDecodeError:
            raise ValueError(
                f"AI returned invalid JSON. Parse error: {e}. "
                f"Raw snippet: {raw[:300]!r}"
            )


def _validate_structure(data: dict, num_slides: int) -> dict:
    """Ensure the parsed dict has the expected shape; patch minor gaps."""
    if 'slides' not in data or not isinstance(data['slides'], list):
        raise ValueError("AI response missing 'slides' array.")

    slides = data['slides']
    if len(slides) == 0:
        raise ValueError("AI returned 0 slides — please try again.")

    valid_types = {'title', 'content', 'chart', 'two_col', 'divider', 'conclusion'}
    for i, s in enumerate(slides):
        s.setdefault('slide_number', i + 1)
        s.setdefault('slide_type', 'content')
        s.setdefault('title', f'Slide {i + 1}')
        s.setdefault('subtitle', None)
        s.setdefault('bullet_points', [])
        s.setdefault('include_chart', False)
        s.setdefault('chart_type', None)
        s.setdefault('chart_data', None)
        s.setdefault('speaker_notes', '')
        if s['slide_type'] not in valid_types:
            s['slide_type'] = 'content'

    data.setdefault('presentation_title', 'Untitled Presentation')
    return data


# ── Provider implementations ──────────────────────────────────────────

def generate_with_openai(api_key: str, model: str, topic: str, num_slides: int, category: str) -> dict:
    from openai import OpenAI, APIError, AuthenticationError, RateLimitError

    client = OpenAI(api_key=api_key)
    prompt = SLIDE_GENERATION_PROMPT.format(topic=topic, num_slides=num_slides, category=category)

    kwargs = dict(
        model=model,
        messages=[
            {'role': 'system',
             'content': (
                 'You are a McKinsey-level presentation strategist. '
                 'Respond with raw JSON only — no markdown fences, no explanation, no preamble. '
                 'Every bullet must contain a specific stat, number, or concrete insight.'
             )},
            {'role': 'user', 'content': prompt},
        ],
        temperature=0.75,
        max_tokens=8192,
    )

    # json_object mode only for chat models that support it (not o-series reasoning models)
    if model in _OPENAI_JSON_MODE_MODELS:
        kwargs['response_format'] = {'type': 'json_object'}

    try:
        response = client.chat.completions.create(**kwargs)
    except AuthenticationError:
        raise ValueError("OpenAI: Invalid API key. Check your key in Settings.")
    except RateLimitError:
        raise ValueError("OpenAI: Rate limit reached. Wait a moment and try again.")
    except APIError as e:
        raise ValueError(f"OpenAI API error: {e.message or str(e)}")

    raw = response.choices[0].message.content or ''
    data = _extract_json(raw)
    return _validate_structure(data, num_slides)


def generate_with_anthropic(api_key: str, model: str, topic: str, num_slides: int, category: str) -> dict:
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    prompt = SLIDE_GENERATION_PROMPT.format(topic=topic, num_slides=num_slides, category=category)

    try:
        message = client.messages.create(
            model=model,
            max_tokens=8192,
            system=(
                'You are a McKinsey-level presentation strategist. '
                'Respond with raw JSON only — no markdown fences, no explanation, no preamble. '
                'Every bullet must contain a specific stat, number, or concrete insight.'
            ),
            messages=[{'role': 'user', 'content': prompt}],
        )
    except anthropic.AuthenticationError:
        raise ValueError("Anthropic: Invalid API key. Check your key in Settings.")
    except anthropic.RateLimitError:
        raise ValueError("Anthropic: Rate limit reached. Wait a moment and try again.")
    except anthropic.BadRequestError as e:
        raise ValueError(f"Anthropic: Bad request — {e}")
    except anthropic.APIStatusError as e:
        raise ValueError(f"Anthropic API error ({e.status_code}): {e.message or str(e)}")

    raw = message.content[0].text if message.content else ''
    data = _extract_json(raw)
    return _validate_structure(data, num_slides)


def generate_with_gemini(api_key: str, model: str, topic: str, num_slides: int, category: str) -> dict:
    import google.generativeai as genai
    from google.api_core.exceptions import InvalidArgument, PermissionDenied, ResourceExhausted

    genai.configure(api_key=api_key)
    prompt = SLIDE_GENERATION_PROMPT.format(topic=topic, num_slides=num_slides, category=category)

    generation_config = {
        'temperature': 0.75,
        'max_output_tokens': 8192,
    }

    try:
        gen_model = genai.GenerativeModel(
            model_name=model,
            generation_config=generation_config,
            system_instruction=(
                'You are a McKinsey-level presentation strategist. '
                'Respond with raw JSON only — no markdown fences, no explanation, no preamble. '
                'Every bullet must contain a specific stat, number, or concrete insight.'
            ),
        )
        response = gen_model.generate_content(prompt)
    except PermissionDenied:
        raise ValueError("Gemini: Invalid API key or permission denied. Check your key in Settings.")
    except ResourceExhausted:
        raise ValueError("Gemini: Quota exceeded. Wait a moment and try again.")
    except InvalidArgument as e:
        # Often means wrong model name
        raise ValueError(
            f"Gemini: Invalid argument — likely an unsupported model name '{model}'. "
            f"Try 'gemini-2.0-flash'. Details: {e}"
        )
    except Exception as e:
        raise ValueError(f"Gemini error: {e}")

    raw = response.text if hasattr(response, 'text') else ''
    if not raw:
        raise ValueError("Gemini returned an empty response. The model may have blocked the request.")

    data = _extract_json(raw)
    return _validate_structure(data, num_slides)


# ── Public entry point ─────────────────────────────────────────────────

def generate_presentation_content(ai_setting, topic: str, num_slides: int, category: str = 'Business') -> dict:
    provider = ai_setting.provider
    api_key  = ai_setting.api_key
    model    = ai_setting.model

    if provider == 'openai':
        return generate_with_openai(api_key, model, topic, num_slides, category)
    elif provider == 'anthropic':
        return generate_with_anthropic(api_key, model, topic, num_slides, category)
    elif provider == 'gemini':
        return generate_with_gemini(api_key, model, topic, num_slides, category)
    else:
        raise ValueError(f"Unknown AI provider: '{provider}'. Expected one of: openai, anthropic, gemini.")
