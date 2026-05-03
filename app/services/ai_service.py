import json
import re

# Models that support response_format=json_object (OpenAI only, not o-series reasoning models)
_OPENAI_JSON_MODE_MODELS = {
    'gpt-4o', 'gpt-4o-mini', 'gpt-4.1', 'gpt-4.1-mini', 'gpt-4.1-nano',
    'gpt-4-turbo', 'gpt-4', 'gpt-3.5-turbo',
}

SLIDE_GENERATION_PROMPT = """You are an expert presentation designer and business analyst.
Generate a complete, professional PowerPoint presentation as a JSON object.

STRICT RULES:
- Produce exactly {num_slides} slides total
- Slide 1: type "title"  |  Last slide: type "conclusion"
- Allowed types: title, content, chart, two_col, divider, conclusion
- Add a "chart" slide when the topic involves data, trends, comparisons, or metrics
- Bullet points: 3-5 concise, punchy lines per slide (empty list [] for title/divider)
- Slide titles: max 8 words, no punctuation at end
- Speaker notes: 2-3 sentences of presenter talking points
- Chart data must contain realistic numbers relevant to the topic

OUTPUT FORMAT — return ONLY the raw JSON object, no markdown fences, no extra text:

{{
  "presentation_title": "Full Presentation Title",
  "slides": [
    {{
      "slide_number": 1,
      "slide_type": "title",
      "title": "Short Punchy Title",
      "subtitle": "Supporting subtitle line",
      "bullet_points": [],
      "include_chart": false,
      "chart_type": null,
      "chart_data": null,
      "speaker_notes": "Welcome the audience and set context."
    }},
    {{
      "slide_number": 2,
      "slide_type": "content",
      "title": "Key Insight or Section Name",
      "subtitle": null,
      "bullet_points": ["Point one", "Point two", "Point three"],
      "include_chart": false,
      "chart_type": null,
      "chart_data": null,
      "speaker_notes": "Walk the audience through each point."
    }},
    {{
      "slide_number": 3,
      "slide_type": "chart",
      "title": "Quantitative Slide Title",
      "subtitle": null,
      "bullet_points": ["Key insight from the data", "Second insight"],
      "include_chart": true,
      "chart_type": "bar",
      "chart_data": {{
        "title": "Descriptive Chart Title",
        "labels": ["Category A", "Category B", "Category C", "Category D"],
        "values": [42, 78, 55, 91],
        "series_name": "Metric ($M)",
        "unit": "$M"
      }},
      "speaker_notes": "Highlight the standout data point and its business impact."
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
             'content': 'You are a professional presentation designer. Respond with raw JSON only — no markdown, no explanation.'},
            {'role': 'user', 'content': prompt},
        ],
        temperature=0.7,
        max_tokens=4096,
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
            max_tokens=4096,
            system='You are a professional presentation designer. Respond with raw JSON only — no markdown fences, no explanation.',
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
        'temperature': 0.7,
        'max_output_tokens': 4096,
    }

    try:
        gen_model = genai.GenerativeModel(
            model_name=model,
            generation_config=generation_config,
            system_instruction='You are a professional presentation designer. Respond with raw JSON only — no markdown fences, no explanation.',
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
