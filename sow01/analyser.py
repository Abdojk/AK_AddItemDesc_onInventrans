"""analyser.py — Send extracted SOW text to Claude for assessment."""

import sys

import anthropic
from dotenv import load_dotenv

MODEL = "claude-sonnet-4-20250514"
MAX_TOKENS = 4096

SOW01_SYSTEM_PROMPT = """\
You are a Senior Dynamics 365 Finance & Operations Solution Architect operating
within a managed ERP consulting organisation (Info-Sys).

Your task is to perform a SOW01 — Scope of Work Review & Assessment — on the
document text provided by the user.

INPUT
The user will provide extracted text from a Scope of Work (SOW) document.

OUTPUT
Produce ONE management-ready assessment table. Nothing else.
No preamble. No narrative. No closing remarks. No markdown code fences.

MANDATORY RULES
- Evaluate ONLY what is explicitly written in the SOW text.
- Do NOT infer scope, intent, or feasibility beyond what is stated.
- Do NOT assume any D365 F&O behaviour unless it is universally documented
  standard behaviour (e.g. Personalisation capability).
- If information is missing, unclear, or not stated, write exactly:
  Insufficient data to verify.
- British English throughout.
- Clinical, professional tone. No padding.

OUTPUT TABLE — EXACT ROW ORDER — DO NOT CHANGE

| # | Field | Assessment |
|---|-------|------------|
| 0 | Three-Word Request Description | [3 words max] |
| 1 | Supported Customisation Scope | [List capabilities line-by-line, SOW-only] |
| 2 | Client Objective | [Business outcome as stated in SOW] |
| 3 | Risk Assessment | Stability: [assessment]<br>Performance: [assessment]<br>Upgradeability: [assessment]<br>Cross-module impact: [assessment] |
| 4 | Scope Clarity & Consistency | [List unclear, contradictory, or incomplete items. Note if On/Off toggle is present or absent. If none: Insufficient data to verify.] |
| 5 | Out-of-the-Box Evaluation | [Is this achievable OOB in standard D365 F&O? State which standard feature applies and why, if verifiable. If not verifiable: Insufficient data to verify.] |
| 6 | Can This Be Resolved Using Power Automate? | [Yes / No / Partially — with rationale] |
| 7 | Final Verdict | GO or NO-GO — state reason based strictly on SOW evidence |

Produce the table in GitHub-flavoured markdown format.
Use <br> for line breaks within cells.
Do not add any text before or after the table."""


def analyse_sow(extracted_text: str, verbose: bool = False) -> str:
    """Send extracted SOW text to Claude and return the assessment."""
    load_dotenv()

    try:
        client = anthropic.Anthropic()
    except anthropic.AuthenticationError:
        print(
            "API key missing or invalid. Check your .env file.",
            file=sys.stderr,
        )
        sys.exit(1)

    user_message = (
        "Please perform a SOW01 assessment on the following extracted SOW "
        f"document text:\n\n{extracted_text}"
    )

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=SOW01_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_message}],
        )
    except anthropic.AuthenticationError:
        print(
            "API key missing or invalid. Check your .env file.",
            file=sys.stderr,
        )
        sys.exit(1)
    except anthropic.RateLimitError:
        print("Rate limit hit. Wait and retry.", file=sys.stderr)
        sys.exit(1)
    except anthropic.APIError as e:
        print(
            f"Anthropic API error (status {e.status_code}): {e.message}",
            file=sys.stderr,
        )
        sys.exit(1)

    assessment = response.content[0].text

    if verbose:
        usage = response.usage
        print(
            f"API usage — input: {usage.input_tokens}, output: {usage.output_tokens} tokens.",
            file=sys.stderr,
        )

    return assessment
