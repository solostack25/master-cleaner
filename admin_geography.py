from fastapi import APIRouter, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse

import supabase_client
import geography_maps
from theme import page_shell, THEME_CSS

router = APIRouter()


def _authed(request: Request) -> bool:
    return request.session.get("user") is not None


def _login_redirect():
    return RedirectResponse(url="/login", status_code=303)


def _nav():
    return """
    <div class="nav-rail">
        <a href="/admin/geography">Overview</a>
        <a href="/admin/geography/regions">State &rarr; Region</a>
        <a href="/admin/geography/rsn">Region &rarr; RSN</a>
        <a href="/admin/geography/offices">State &rarr; Office</a>
        <a href="/admin/geography/chapters">Office &rarr; Chapter</a>
        <a href="/admin/geography/overrides">Special Overrides</a>
        <a href="/admin/geography/cities">City/Zip &rarr; Office</a>
        <a href="/admin/geography/keywords">Keyword &rarr; Program</a>
        <a href="/" class="back">&larr; Back to Cleaner</a>
    </div>
    """


def _flash(message):
    if not message:
        return ""
    return f'<div class="flash">{message}</div>'


@router.get("/admin/geography", response_class=HTMLResponse)
def admin_geography_home(request: Request):
    if not _authed(request):
        return _login_redirect()

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>Geography &amp; Config Admin</h1>
        <hr class="stitch">
        {_nav()}
        <p>
            This dashboard edits the shared geography and keyword rules that every
            report cleaner uses (state &rarr; region, field offices, chapters, etc.).
            Changes take effect immediately for all future "Run Cleaner" runs &mdash;
            no redeploy needed.
        </p>
        <ul>
            <li><b>State &rarr; Region</b> &mdash; e.g. moving Kansas from South Central to Midwest 2.</li>
            <li><b>Region &rarr; RSN</b> &mdash; the number assigned to each region.</li>
            <li><b>State &rarr; Office</b> &mdash; default field office per state (most states).</li>
            <li><b>Office &rarr; Chapter</b> &mdash; which chapter a field office rolls up into.</li>
            <li><b>Special Overrides</b> &mdash; edge-case rules, like the Kansas City Office override.</li>
            <li><b>City/Zip &rarr; Office</b> &mdash; the CA/TX/FL/VA/MO lookup table (9,000+ rows, searchable).</li>
            <li><b>Keyword &rarr; Program</b> &mdash; auto-detects program from campaign/source names.</li>
        </ul>
    </div></body></html>
    """)


# ------------------------------------------------------------
# State -> Region
# ------------------------------------------------------------

@router.get("/admin/geography/regions", response_class=HTMLResponse)
def list_regions(request: Request, saved: str = ""):
    if not _authed(request):
        return _login_redirect()

    rows = sorted(supabase_client.fetch_all("geo_state_to_region"), key=lambda r: r["state"])
    all_regions = sorted(set(r["region"] for r in rows))

    rows_html = ""
    for r in rows:
        options = "".join(
            f'<option value="{region}" {"selected" if region == r["region"] else ""}>{region}</option>'
            for region in all_regions
        )
        rows_html += f"""
        <tr>
            <form method="post" action="/admin/geography/regions/update">
                <td>{r['state']}</td>
                <td>
                    <input type="hidden" name="state" value="{r['state']}">
                    <select name="region">{options}</select>
                </td>
                <td><button class="save-btn" type="submit">Save</button></td>
            </form>
        </tr>
        """

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>State &rarr; Region</h1>
        <hr class="stitch">
        {_nav()}
        {_flash(saved)}
        <table>
            <tr><th>State</th><th>Region</th><th></th></tr>
            {rows_html}
        </table>
    </div></body></html>
    """)


@router.post("/admin/geography/regions/update")
def update_region(request: Request, state: str = Form(...), region: str = Form(...)):
    if not _authed(request):
        return _login_redirect()

    supabase_client.update_row("geo_state_to_region", "state", state, {"region": region})
    geography_maps.refresh_geo_cache()
    return RedirectResponse(url=f"/admin/geography/regions?saved=Saved+{state}+%E2%86%92+{region}", status_code=303)


# ------------------------------------------------------------
# Region -> RSN
# ------------------------------------------------------------

@router.get("/admin/geography/rsn", response_class=HTMLResponse)
def list_rsn(request: Request, saved: str = ""):
    if not _authed(request):
        return _login_redirect()

    rows = sorted(supabase_client.fetch_all("geo_region_to_rsn"), key=lambda r: r["rsn"])

    rows_html = ""
    for r in rows:
        rows_html += f"""
        <tr>
            <form method="post" action="/admin/geography/rsn/update">
                <td>{r['region']}</td>
                <td>
                    <input type="hidden" name="region" value="{r['region']}">
                    <input type="text" name="rsn" value="{r['rsn']}">
                </td>
                <td><button class="save-btn" type="submit">Save</button></td>
            </form>
        </tr>
        """

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>Region &rarr; RSN</h1>
        <hr class="stitch">
        {_nav()}
        {_flash(saved)}
        <table>
            <tr><th>Region</th><th>RSN</th><th></th></tr>
            {rows_html}
        </table>
    </div></body></html>
    """)


@router.post("/admin/geography/rsn/update")
def update_rsn(request: Request, region: str = Form(...), rsn: int = Form(...)):
    if not _authed(request):
        return _login_redirect()

    supabase_client.update_row("geo_region_to_rsn", "region", region, {"rsn": rsn})
    geography_maps.refresh_geo_cache()
    return RedirectResponse(url="/admin/geography/rsn?saved=Saved", status_code=303)


# ------------------------------------------------------------
# State -> Default Field Office
# ------------------------------------------------------------

@router.get("/admin/geography/offices", response_class=HTMLResponse)
def list_offices(request: Request, saved: str = ""):
    if not _authed(request):
        return _login_redirect()

    rows = sorted(supabase_client.fetch_all("geo_state_to_field_office"), key=lambda r: r["state"])

    rows_html = ""
    for r in rows:
        rows_html += f"""
        <tr>
            <form method="post" action="/admin/geography/offices/update">
                <td>{r['state']}</td>
                <td>
                    <input type="hidden" name="state" value="{r['state']}">
                    <input type="text" name="field_office" value="{r['field_office']}">
                </td>
                <td><button class="save-btn" type="submit">Save</button></td>
            </form>
        </tr>
        """

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>State &rarr; Default Field Office</h1>
        <hr class="stitch">
        {_nav()}
        {_flash(saved)}
        <p style="color:#555;">CA, TX, FL, VA, MO are handled by the City/Zip table instead, so they don't appear here.</p>
        <table>
            <tr><th>State</th><th>Field Office</th><th></th></tr>
            {rows_html}
        </table>
    </div></body></html>
    """)


@router.post("/admin/geography/offices/update")
def update_office(request: Request, state: str = Form(...), field_office: str = Form(...)):
    if not _authed(request):
        return _login_redirect()

    supabase_client.update_row("geo_state_to_field_office", "state", state, {"field_office": field_office})
    geography_maps.refresh_geo_cache()
    return RedirectResponse(url="/admin/geography/offices?saved=Saved", status_code=303)


# ------------------------------------------------------------
# Field Office -> Chapter
# ------------------------------------------------------------

@router.get("/admin/geography/chapters", response_class=HTMLResponse)
def list_chapters(request: Request, saved: str = ""):
    if not _authed(request):
        return _login_redirect()

    rows = sorted(supabase_client.fetch_all("geo_chapter_map"), key=lambda r: (r["region"], r["field_office"]))

    rows_html = ""
    for r in rows:
        office_label = r['field_office'] if r['field_office'] != '__default__' else '(default for region)'
        rows_html += f"""
        <tr>
            <form method="post" action="/admin/geography/chapters/update">
                <td>{r['region']}</td>
                <td>{office_label}</td>
                <td>
                    <input type="hidden" name="id" value="{r['id']}">
                    <input type="text" name="chapter" value="{r['chapter']}">
                </td>
                <td><button class="save-btn" type="submit">Save</button></td>
            </form>
        </tr>
        """

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>Field Office &rarr; Chapter</h1>
        <hr class="stitch">
        {_nav()}
        {_flash(saved)}
        <p style="color:#555;">"(default for region)" rows apply to any office in that region not listed individually.</p>
        <table>
            <tr><th>Region</th><th>Field Office</th><th>Chapter</th><th></th></tr>
            {rows_html}
        </table>
    </div></body></html>
    """)


@router.post("/admin/geography/chapters/update")
def update_chapter(request: Request, id: str = Form(...), chapter: str = Form(...)):
    if not _authed(request):
        return _login_redirect()

    supabase_client.update_row("geo_chapter_map", "id", id, {"chapter": chapter})
    geography_maps.refresh_geo_cache()
    return RedirectResponse(url="/admin/geography/chapters?saved=Saved", status_code=303)


# ------------------------------------------------------------
# Special overrides (e.g. Kansas City Office edge case)
# ------------------------------------------------------------

@router.get("/admin/geography/overrides", response_class=HTMLResponse)
def list_overrides(request: Request, saved: str = ""):
    if not _authed(request):
        return _login_redirect()

    rows = supabase_client.fetch_all("geo_special_overrides")

    rows_html = ""
    for r in rows:
        checked = "checked" if r.get("active") else ""
        rows_html += f"""
        <tr>
            <form method="post" action="/admin/geography/overrides/update">
                <td>{r['description']}</td>
                <td>{r['when_field_office']}</td>
                <td>
                    <input type="hidden" name="id" value="{r['id']}">
                    <input type="text" name="set_region" value="{r.get('set_region') or ''}">
                </td>
                <td><input type="text" name="set_rsn" value="{r.get('set_rsn') or ''}"></td>
                <td><input type="text" name="set_state" value="{r.get('set_state') or ''}"></td>
                <td><input type="checkbox" name="active" {checked}></td>
                <td><button class="save-btn" type="submit">Save</button></td>
            </form>
        </tr>
        """

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>Special Override Rules</h1>
        <hr class="stitch">
        {_nav()}
        {_flash(saved)}
        <p style="color:#555;">
            Edge-case rules that override the normal region/RSN/state for a specific
            resolved field office &mdash; e.g. Kansas City Office forcing region back
            to South Central since it spans KS/MO.
        </p>
        <table>
            <tr><th>Description</th><th>When Field Office =</th><th>Set Region</th><th>Set RSN</th><th>Set State</th><th>Active</th><th></th></tr>
            {rows_html}
        </table>
    </div></body></html>
    """)


@router.post("/admin/geography/overrides/update")
def update_override(
    request: Request,
    id: str = Form(...),
    set_region: str = Form(""),
    set_rsn: str = Form(""),
    set_state: str = Form(""),
    active: bool = Form(False),
):
    if not _authed(request):
        return _login_redirect()

    supabase_client.update_row("geo_special_overrides", "id", id, {
        "set_region": set_region or None,
        "set_rsn": int(set_rsn) if set_rsn else None,
        "set_state": set_state or None,
        "active": active,
    })
    geography_maps.refresh_geo_cache()
    return RedirectResponse(url="/admin/geography/overrides?saved=Saved", status_code=303)


# ------------------------------------------------------------
# City/Zip -> Office (large table, search only)
# ------------------------------------------------------------

@router.get("/admin/geography/cities", response_class=HTMLResponse)
def search_cities(request: Request, q: str = "", saved: str = ""):
    if not _authed(request):
        return _login_redirect()

    results = supabase_client.search_rows("geo_state_splits_cities", "main_city", q, limit=50) if q else []

    rows_html = ""
    for r in results:
        rows_html += f"""
        <tr>
            <form method="post" action="/admin/geography/cities/update">
                <td>{r['main_city']}</td>
                <td>{r['state']}</td>
                <td>{r['zip_code']}</td>
                <td>
                    <input type="hidden" name="id" value="{r['id']}">
                    <input type="text" name="assigned_to" value="{r['assigned_to']}">
                </td>
                <td><button class="save-btn" type="submit">Save</button></td>
            </form>
        </tr>
        """

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>City/Zip &rarr; Office (CA, TX, FL, VA, MO)</h1>
        <hr class="stitch">
        {_nav()}
        {_flash(saved)}
        <form method="get" action="/admin/geography/cities">
            <input type="text" name="q" placeholder="Search by city name..." value="{q}">
            <button class="save-btn" type="submit">Search</button>
        </form>
        <p style="color:#555;">9,000+ rows &mdash; search by city name to find and edit assignments.</p>
        <table>
            <tr><th>City</th><th>State</th><th>Zip</th><th>Assigned Office</th><th></th></tr>
            {rows_html}
        </table>
    </div></body></html>
    """)


@router.post("/admin/geography/cities/update")
def update_city(request: Request, id: str = Form(...), assigned_to: str = Form(...)):
    if not _authed(request):
        return _login_redirect()

    supabase_client.update_row("geo_state_splits_cities", "id", id, {"assigned_to": assigned_to})
    geography_maps.refresh_geo_cache()
    return RedirectResponse(url="/admin/geography/cities?saved=Saved", status_code=303)


# ------------------------------------------------------------
# Keyword -> Program
# ------------------------------------------------------------

@router.get("/admin/geography/keywords", response_class=HTMLResponse)
def list_keywords(request: Request, q: str = "", saved: str = ""):
    if not _authed(request):
        return _login_redirect()

    if q:
        rows = supabase_client.search_rows("geo_keyword_to_program", "keyword", q, limit=200)
    else:
        rows = supabase_client.fetch_all("geo_keyword_to_program", order="keyword")

    rows_html = ""
    for r in rows:
        rows_html += f"""
        <tr>
            <form method="post" action="/admin/geography/keywords/update">
                <td>{r['keyword']}</td>
                <td>
                    <input type="hidden" name="keyword" value="{r['keyword']}">
                    <input type="text" name="program" value="{r['program']}">
                </td>
                <td><button class="save-btn" type="submit">Save</button></td>
            </form>
        </tr>
        """

    return HTMLResponse(f"""
    <html><head>{THEME_CSS}</head><body><div class="shell wide">
        <h1>Keyword &rarr; Program</h1>
        <hr class="stitch">
        {_nav()}
        {_flash(saved)}
        <form method="get" action="/admin/geography/keywords">
            <input type="text" name="q" placeholder="Search keywords..." value="{q}">
            <button class="save-btn" type="submit">Search</button>
        </form>
        <table>
            <tr><th>Keyword</th><th>Program</th><th></th></tr>
            {rows_html}
        </table>
    </div></body></html>
    """)


@router.post("/admin/geography/keywords/update")
def update_keyword(request: Request, keyword: str = Form(...), program: str = Form(...)):
    if not _authed(request):
        return _login_redirect()

    supabase_client.update_row("geo_keyword_to_program", "keyword", keyword, {"program": program})
    geography_maps.refresh_geo_cache()
    return RedirectResponse(url="/admin/geography/keywords?saved=Saved", status_code=303)
