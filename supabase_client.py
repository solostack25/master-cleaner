import os
import httpx

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")


class SupabaseConnectionError(Exception):
    """Raised when Supabase can't be reached or rejects the request, with
    a message that points at the actual likely cause instead of a raw
    stack trace."""
    pass


def _headers():
    return {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }


def _check_configured():
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        raise SupabaseConnectionError(
            "SUPABASE_URL and/or SUPABASE_SERVICE_ROLE_KEY are not set on Render. "
            "Check the Environment tab on the Render service."
        )


def _request(method, url, **kwargs):
    try:
        with httpx.Client(timeout=30) as client:
            response = client.request(method, url, **kwargs)
    except httpx.ConnectError as error:
        raise SupabaseConnectionError(
            f"Could not connect to Supabase at {SUPABASE_URL}. This usually means "
            "either SUPABASE_URL is wrong, or the Supabase project is paused "
            "(free-tier projects pause automatically after about a week with no "
            "activity -- check the Supabase dashboard and click Resume if so)."
        ) from error
    except httpx.TimeoutException as error:
        raise SupabaseConnectionError(
            "Supabase did not respond in time. The project may be paused or "
            "waking up from being paused -- try again in a minute."
        ) from error

    if response.status_code in (401, 403):
        raise SupabaseConnectionError(
            "Supabase rejected the request as unauthorized. SUPABASE_SERVICE_ROLE_KEY "
            "on Render is likely wrong, expired, or was rotated in Supabase -- check "
            "Project Settings -> API in Supabase and compare against the Render env var."
        )

    if response.status_code >= 400:
        raise SupabaseConnectionError(
            f"Supabase returned an error ({response.status_code}) for {url}: "
            f"{response.text[:300]}"
        )

    return response


def fetch_all(table, order=None, limit=10000):
    """Fetch all rows from a table (paginated internally to avoid PostgREST's
    default row cap)."""
    _check_configured()
    rows = []
    offset = 0
    page_size = 1000

    while True:
        params = {"select": "*", "limit": page_size, "offset": offset}
        if order:
            params["order"] = order

        response = _request(
            "GET", f"{SUPABASE_URL}/rest/v1/{table}", headers=_headers(), params=params
        )
        page = response.json()
        rows.extend(page)

        if len(page) < page_size or len(rows) >= limit:
            break
        offset += page_size

    return rows


def search_rows(table, column, query_text, limit=50):
    """Case-insensitive partial match search on a single column."""
    _check_configured()
    response = _request(
        "GET", f"{SUPABASE_URL}/rest/v1/{table}",
        headers=_headers(),
        params={"select": "*", column: f"ilike.*{query_text}*", "limit": limit},
    )
    return response.json()


def insert_row(table, data):
    _check_configured()
    response = _request(
        "POST", f"{SUPABASE_URL}/rest/v1/{table}", headers=_headers(), json=data
    )
    return response.json()


def update_row(table, match_column, match_value, data):
    _check_configured()
    response = _request(
        "PATCH", f"{SUPABASE_URL}/rest/v1/{table}",
        headers=_headers(),
        params={match_column: f"eq.{match_value}"},
        json=data,
    )
    return response.json()


def delete_row(table, match_column, match_value):
    _check_configured()
    _request(
        "DELETE", f"{SUPABASE_URL}/rest/v1/{table}",
        headers=_headers(),
        params={match_column: f"eq.{match_value}"},
    )
    return True
