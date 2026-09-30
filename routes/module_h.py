"""Module H: an opt-in workbench with area selection over the shared F engine.

The original document is composed at request time, not forked. Existing control
IDs, handlers, validation and API calls remain authoritative. H explicitly opts
into all-area head selection; F keeps ranked-K defaults. Saved drawings are shared.
"""
import re
from flask import Flask, render_template


def compose_document(base: str, shell: str) -> str:
    """Attach H-only presentation assets, failing loudly if the contract changes."""
    replacements = {
        "<title>Module F · 배관 네트워크 위상 복원</title>":
            "<title>Module H · 스프링클러 설계 워크벤치</title>",
        "</head>": '<link rel="stylesheet" href="/static/module_h.css?v=20260928-area-1">\n'
            '<link rel="stylesheet" href="/static/module_h_access.css?v=20260929-access-1">\n'
            '<link rel="stylesheet" href="/static/module_h_attributes.css?v=20260928-1">\n'
            '<script src="/static/module_h_activity.js?v=20260928-activity-1"></script>\n</head>',
        "<body>": '<body class="module-h">',
        '<span class="chip">MODULE F</span>': '<span class="chip">MODULE H</span>',
        "<h1>배관 네트워크 위상 복원</h1>": "<h1>스프링클러 설계 워크벤치</h1>",
        "</body>": shell + '\n<script src="/static/module_h_progress.js?v=20260928-diameter-1"></script>'
            '\n<script src="/static/module_h_windows.js?v=20260928-live-1"></script>'
            '\n<script src="/static/module_h.js?v=20260928-diameter-1"></script>'
            '\n<script src="/static/module_h_access.js?v=20260929-access-1"></script>'
            '\n<script src="/static/module_h_attributes.js?v=20260928-review-2"></script>'
            '\n<script src="/static/module_h_properties.js?v=20260928-review-2"></script>'
            '\n<script src="/static/module_h_hydraulics.js?v=20260930-1"></script>'
            '\n<script src="/static/module_h_property_source.js?v=20260929-dock-1"></script>'
            '\n<script src="/static/module_h_references.js?v=20260929-source-reference-1"></script>'
            '\n<script src="/static/module_h_evidence.js?v=20260928-review-2"></script>\n</body>',
    }
    for anchor, replacement in replacements.items():
        if base.count(anchor) != 1:
            raise RuntimeError(f"Module H layout contract changed: {anchor!r}")
        base = base.replace(anchor, replacement, 1)
    # H needs the read-only operation getter even if F's script URL was cached.
    for version in ('20260928-area-1','20260928-diameter-1','20260928-activity-1','20260928-live-1','20260928-review-2','20260928-1','20260929-dock-1','20260929-view-1','20260929-source-reference-1','20260929-access-1'):
        base = base.replace(version, '20260930-multiselect-1')
    return re.sub(r'(<script src="/static/module_f(?:_editor)?\.js[^"\n]*)"',
                  r'\1&h_observer=20260930-multiselect-1"', base)


def register(app: Flask) -> None:
    """Register the separate H page under the application's existing auth gate."""
    from routes.module_h_progress import install
    from routes.module_h_attributes import register as register_attributes
    register_attributes(app)
    from routes.module_h_access import register as register_access
    register_access(app)
    from routes.module_f.cancellation import install as install_cancellation
    install_cancellation(app,register_controls=False,endpoint_prefix='h_attribute_')
    install(app)
    @app.get("/api/module-h/selection-policy")
    def module_h_selection_policy() -> dict:
        """Prevent refreshed static assets from using an older K-only server."""
        return {"ok": True, "selection_mode": "area_all", "version": 1}

    @app.get("/module-h")
    def module_h_page() -> str:
        return compose_document(
            render_template("module_f.html"), render_template("module_h_shell.html")
        )
