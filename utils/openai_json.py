import json
import logging
import os
import re
import time

from openai import OpenAI


_FENCED_JSON_RE = re.compile(r"^```(?:json)?\s*|\s*```$", re.IGNORECASE)


def _strip_json_fences(content):
    if not content:
        return "{}"
    return _FENCED_JSON_RE.sub("", content.strip())


def chat_completion_json(
    system_prompt,
    user_prompt,
    *,
    model=None,
    temperature=0.0,
    max_tokens=2000,
    retries=3,
):
    """
    Call the OpenAI Chat Completions API and parse a JSON object response.
    """
    selected_model = model or os.getenv("OPENAI_CHAT_MODEL", "gpt-4o")
    client = OpenAI()
    last_error = None

    json_instruction = (
        "Return only valid JSON. Do not include markdown fences, commentary, or prose. "
        "If a value is unknown, use an empty string or an empty array."
    )

    for attempt in range(1, retries + 1):
        try:
            response = client.chat.completions.create(
                model=selected_model,
                temperature=temperature,
                max_tokens=max_tokens,
                messages=[
                    {
                        "role": "system",
                        "content": f"{system_prompt}\n\n{json_instruction}",
                    },
                    {"role": "user", "content": user_prompt},
                ],
            )
            content = response.choices[0].message.content or "{}"
            return json.loads(_strip_json_fences(content))
        except Exception as exc:
            last_error = exc
            logging.warning(
                "OpenAI JSON call failed on attempt %s/%s: %s",
                attempt,
                retries,
                exc,
            )
            if attempt < retries:
                time.sleep(2 ** (attempt - 1))

    raise last_error
