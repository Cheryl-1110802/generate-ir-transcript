"""
Fill in the {{ variable }} / {% if %} / {% for %} placeholders in
transcript_en_template.html (or the zh version, once converted) with a
context dict — the same shape of dict your existing
core/report_generator.py already builds for docxtpl.

Two renderers are provided:

  render_with_jinja2(html, context)
      Uses real Jinja2 (autoescape on, so you can drop the manual
      fix_ampersand_in_context '&' hack you needed for the docx/XML
      pipeline — Jinja2 escapes &, <, > for you automatically).
      Needs the `jinja2` package importable. Inside n8n's Python Code
      node (Pyodide), that means:
          import micropip
          await micropip.install("jinja2")
      which requires the n8n process to reach PyPI/the Pyodide package
      index over the network. If your n8n instance has no outbound
      internet access, this will fail — use render_simple() instead.

  render_simple(html, context)
      Zero-dependency, pure-stdlib fallback. It does NOT implement
      Jinja2 — it hand-covers exactly the constructs actually used in
      this template: {{ a.b }} dotted lookups, a single-level
      {% if var == "x" %} / {% if var != "x" %} / {% if var %} ... {% endif %},
      and a single-level {% for item in list %} ... {% endfor %}.
      No elif/else, no filters, no nested loops/ifs — this template
      doesn't use any, so it doesn't need to.

Usage inside an n8n Python Code node (render_simple, no deps):

    from render_html_template import render_simple

    template_html = _  # e.g. read from a Read/Set node upstream, or embed inline
    context = {
        "event_quarter": "3Q25",
        "chairman": "...",
        "financial_results": {"revenue_abbv": "1.2B", ...},
        "chairman_remarks": [{"page": 1, "title": "...", "content": "..."}],
        ...
    }
    rendered = render_simple(template_html, context)
    return [{"json": {"html": rendered}}]
"""
import html as htmlmod
import re


def _get_value(context, dotted_path):
    """Resolve 'a.b.c' against nested dicts. Raises KeyError if missing."""
    node = context
    for part in dotted_path.split('.'):
        if isinstance(node, dict):
            node = node[part]
        else:
            node = getattr(node, part)
    return node


def render_with_jinja2(template_html, context):
    from jinja2 import Environment, StrictUndefined
    env = Environment(autoescape=True, undefined=StrictUndefined)
    return env.from_string(template_html).render(**context)


# ---------------------------------------------------------------------------
# Dependency-free fallback covering exactly this template's constructs.
# ---------------------------------------------------------------------------

_VAR_RE = re.compile(r'\{\{\s*([\w\.]+)\s*\}\}')
_FOR_RE = re.compile(r'\{%\s*for\s+(\w+)\s+in\s+([\w\.]+)\s*%\}(.*?)\{%\s*endfor\s*%\}', re.DOTALL)
_IF_RE = re.compile(
    r'\{%\s*if\s+([\w\.]+)\s*(==|!=)?\s*"?([^%"]*?)"?\s*%\}(.*?)\{%\s*endif\s*%\}', re.DOTALL
)


def _substitute_vars(text, context, escape):
    def repl(m):
        try:
            value = _get_value(context, m.group(1))
        except (KeyError, AttributeError):
            raise KeyError(f"Missing template variable: {m.group(1)!r}")
        value = '' if value is None else str(value)
        return htmlmod.escape(value, quote=False) if escape else value
    return _VAR_RE.sub(repl, text)


def _eval_condition(var_path, op, literal, context):
    try:
        value = _get_value(context, var_path)
    except (KeyError, AttributeError):
        value = None
    if op is None:
        return bool(value)
    if op == '==':
        return str(value) == literal
    return str(value) != literal  # op == '!='


def render_simple(template_html, context, escape=True):
    def for_repl(m):
        loop_var, list_path, body = m.group(1), m.group(2), m.group(3)
        try:
            items = _get_value(context, list_path)
        except (KeyError, AttributeError):
            items = []
        out = []
        for item in items:
            inner_ctx = dict(context)
            inner_ctx[loop_var] = item
            out.append(_substitute_vars(body, inner_ctx, escape))
        return ''.join(out)

    def if_repl(m):
        var_path, op, literal, body = m.group(1), m.group(2), m.group(3), m.group(4)
        return body if _eval_condition(var_path, op, literal, context) else ''

    rendered = _FOR_RE.sub(for_repl, template_html)
    rendered = _IF_RE.sub(if_repl, rendered)
    rendered = _substitute_vars(rendered, context, escape)
    return rendered


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print("Usage: python render_html_template.py <template.html>")
        sys.exit(1)
    with open(sys.argv[1], encoding='utf-8') as f:
        tpl = f.read()
    demo_context = {
        "event_quarter": "3Q25", "event_date": "2025/11/10", "chairman": "Mr. Michael Ho",
        "this_quarter_en": "3Q25", "this_quarter_year": "3Q25", "this_year": "2025", "ytd": "9M25",
        "financial_officer": "Ms. Cammie Chan", "president": "Mr. Tony Han",
        "president_name": "Tony", "chairman_name": "Michael",
        "this_quarter": "Q3",
        "opening_remarks": {"content": "[opening remarks placeholder]"},
        "financial_results": {k: "N/A" for k in [
            "revenue_abbv", "revenue", "revenue_qoq", "revenue_yoy",
            "operating_expenses_abbv", "operating_expenses", "operating_expenses_qoq", "operating_expenses_yoy",
            "operating_income_abbv", "operating_income", "operating_income_qoq", "operating_income_yoy",
            "operating_margin", "operating_margin_qoq", "operating_margin_yoy",
            "net_income_abbv", "net_income", "net_income_qoq", "net_income_yoy", "eps",
        ]},
        "revenue_streams": {k: "N/A" for k in [
            "licensing", "licensing_qoq", "licensing_yoy", "licensing_qoq_us", "licensing_yoy_us",
            "royalty", "royalty_qoq", "royalty_yoy", "royalty_qoq_us", "royalty_yoy_us",
            "total_qoq_us", "total_yoy_us",
            "licensing_ytd", "licensing_yoy_ytd", "licensing_yoy_ytd_us",
            "royalty_ytd", "royalty_yoy_ytd", "royalty_yoy_ytd_us",
            "total_yoy_ytd", "total_yoy_ytd_us",
        ]},
        "tech": {k: "N/A" for k in [
            "neobit_total", "neobit_licensing_qoq", "neobit_licensing_yoy", "neobit_royalty_qoq", "neobit_royalty_yoy",
            "neofuse_total", "neofuse_licensing_qoq", "neofuse_licensing_yoy", "neofuse_royalty_qoq", "neofuse_royalty_yoy",
            "pufbased_total", "pufbased_licensing_qoq", "pufbased_licensing_yoy", "pufbased_royalty",
            "mtp_total", "mtp_licensing_qoq", "mtp_licensing_yoy", "mtp_royalty_qoq", "mtp_royalty_yoy",
            "neobit_licensing_yoy_ytd", "neobit_royalty_yoy_ytd", "neobit_ytd",
            "neofuse_licensing_yoy_ytd", "neofuse_royalty_yoy_ytd", "neofuse_ytd",
            "pufbased_licensing_yoy_ytd", "pufbased_ytd",
            "mtp_licensing_yoy_ytd", "mtp_royalty_yoy_ytd", "mtp_ytd",
        ]},
        "wafer_size": {k: "N/A" for k in [
            "eight_inch", "eight_inch_qoq", "eight_inch_yoy", "eight_inch_qoq_us", "eight_inch_yoy_us",
            "twelve_inch", "twelve_inch_qoq", "twelve_inch_yoy", "twelve_inch_qoq_us", "twelve_inch_yoy_us",
        ]},
        "new_tapeouts": {"total": "N/A"},
        "future_outlook": {k: "[placeholder]" for k in [
            "licensing_outlook", "royalty_outlook", "new_ip_tech", "biz_dev_platform",
        ]},
        "chairman_remarks": [
            {"page": 1, "title": "Demo Title", "content": "Demo content for the featured topic."},
        ],
    }
    result = render_simple(tpl, demo_context)
    out_path = sys.argv[1].replace('.html', '.rendered.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(result)
    print("Rendered demo output ->", out_path)
