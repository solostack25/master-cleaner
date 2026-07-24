THEME_CSS = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
    :root {
        --ink: #14181C;
        --panel: #1B2027;
        --panel-line: #2A3038;
        --paper: #EDEAE2;
        --muted: #939AA6;
        --brass: #C89B3C;
        --brass-dim: #8A6D2C;
        --sage: #6FA287;
        --brick: #C1584A;
    }

    * { box-sizing: border-box; }

    body {
        background: var(--ink);
        color: var(--paper);
        font-family: 'IBM Plex Sans', Arial, sans-serif;
        margin: 0;
        line-height: 1.55;
        -webkit-font-smoothing: antialiased;
    }

    .shell {
        max-width: 760px;
        margin: 0 auto;
        padding: 48px 24px 80px 24px;
    }

    .shell.wide { max-width: 1040px; }

    h1, h2, h3 {
        font-family: 'Fraunces', Georgia, serif;
        font-weight: 600;
        letter-spacing: 0.2px;
        margin: 0 0 6px 0;
        color: var(--paper);
    }

    h1 { font-size: 28px; }
    h2 { font-size: 20px; }

    .stitch {
        border: none;
        border-top: 2px dotted var(--panel-line);
        margin: 14px 0 28px 0;
    }

    .seal {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 34px;
        height: 34px;
        border-radius: 50%;
        border: 1.5px solid var(--brass);
        color: var(--brass);
        font-family: 'Fraunces', serif;
        font-weight: 600;
        font-size: 15px;
        flex-shrink: 0;
    }

    .topbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        margin-bottom: 4px;
    }

    .topbar-title {
        display: flex;
        align-items: center;
        gap: 14px;
    }

    .topbar-links a {
        color: var(--muted);
        text-decoration: none;
        font-size: 14px;
        margin-left: 18px;
        border-bottom: 1px solid transparent;
        padding-bottom: 2px;
        transition: color 0.15s ease, border-color 0.15s ease;
    }

    .topbar-links a:hover {
        color: var(--brass);
        border-color: var(--brass);
    }

    p { color: var(--paper); }
    p.lede { color: var(--muted); font-size: 15px; max-width: 58ch; }

    strong { color: var(--paper); font-weight: 600; }

    label {
        display: block;
        font-size: 13px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        color: var(--muted);
        margin-bottom: 6px;
    }

    input[type=text], input[type=password], select {
        width: 100%;
        background: var(--panel);
        border: 1px solid var(--panel-line);
        color: var(--paper);
        padding: 10px 12px;
        border-radius: 4px;
        font-family: 'IBM Plex Sans', sans-serif;
        font-size: 14px;
    }

    input[type=text]:focus, input[type=password]:focus, select:focus {
        outline: 2px solid var(--brass);
        outline-offset: 1px;
        border-color: var(--brass);
    }

    input[type=file] {
        color: var(--muted);
        font-size: 14px;
    }

    button, .btn {
        background: var(--brass);
        color: var(--ink);
        border: none;
        padding: 11px 22px;
        border-radius: 4px;
        font-family: 'IBM Plex Sans', sans-serif;
        font-weight: 600;
        font-size: 14px;
        cursor: pointer;
        transition: background 0.15s ease;
    }

    button:hover, .btn:hover { background: #DBAF52; }

    button:focus-visible, a:focus-visible, input:focus-visible, select:focus-visible {
        outline: 2px solid var(--brass);
        outline-offset: 2px;
    }

    button.save-btn {
        padding: 7px 14px;
        font-size: 13px;
    }

    .card {
        background: var(--panel);
        border: 1px solid var(--panel-line);
        border-radius: 6px;
        padding: 22px 24px;
        margin-top: 22px;
    }

    .instructions {
        background: var(--panel);
        border-left: 3px solid var(--brass);
        border-radius: 0 4px 4px 0;
        padding: 16px 20px;
        margin-top: 18px;
        font-size: 14px;
        color: var(--muted);
    }

    .instructions strong { color: var(--paper); }

    table {
        border-collapse: collapse;
        width: 100%;
        margin-top: 18px;
        font-size: 14px;
    }

    th {
        text-align: left;
        color: var(--muted);
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        font-weight: 500;
        padding: 8px 12px;
        border-bottom: 1px solid var(--panel-line);
    }

    td {
        padding: 8px 12px;
        border-bottom: 1px solid var(--panel-line);
        font-family: 'IBM Plex Mono', monospace;
        font-size: 13px;
    }

    tr:hover td { background: rgba(200, 155, 60, 0.04); }

    .flash {
        background: rgba(111, 162, 135, 0.12);
        border: 1px solid var(--sage);
        color: var(--sage);
        border-radius: 4px;
        padding: 10px 16px;
        margin-bottom: 18px;
        font-size: 14px;
        display: inline-block;
    }

    .error-text {
        background: rgba(193, 88, 74, 0.12);
        border: 1px solid var(--brick);
        color: #E39187;
        border-radius: 4px;
        padding: 10px 16px;
        font-size: 14px;
        margin: 16px 0;
    }

    .nav-rail {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin: 18px 0 28px 0;
    }

    .nav-rail a {
        color: var(--muted);
        text-decoration: none;
        font-size: 13px;
        padding: 7px 12px;
        border: 1px solid var(--panel-line);
        border-radius: 4px;
        transition: color 0.15s ease, border-color 0.15s ease;
    }

    .nav-rail a:hover { color: var(--brass); border-color: var(--brass); }
    .nav-rail a.back { margin-left: auto; }

    ul.data-notes { color: var(--muted); font-size: 14px; padding-left: 20px; }
    ul.data-notes li { margin-bottom: 6px; }
    ul.data-notes strong { color: var(--paper); }
</style>
"""


def page_shell(title, body_html, wide=False):
    width_class = " wide" if wide else ""
    return f"""
    <html>
        <head>
            <title>{title}</title>
            {THEME_CSS}
        </head>
        <body>
            <div class="shell{width_class}">
                {body_html}
            </div>
        </body>
    </html>
    """
