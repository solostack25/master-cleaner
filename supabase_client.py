import os
import httpx

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")


def _headers():
    return {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }


def _check_configured():
    if not SUPABASE_URL or not SUPABASE_SERVICE_ROLE_KEY:
        raise RuntimeError(
            "SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY must be set as environment "
            "variables for the geography admin dashboard to work."
        )


def fetch_all(table, order=None, limit=10000):
    """Fetch all rows from a table (paginated internally to avoid PostgREST's
    default row cap)."""
    _check_configured()
    rows = []
    offset = 0
    page_size = 1000

    with httpx.Client(timeout=30) as client:
        while True:
            params = {"select": "*", "limit": page_size, "offset": offset}
            if order:
                params["order"] = order

            response = client.get(
                f"{SUPABASE_URL}/rest/v1/{table}",
                headers=_headers(),
                params=params,
            )
            response.raise_for_status()
            page = response.json()
            rows.extend(page)

            if len(page) < page_size or len(rows) >= limit:
                break
            offset += page_size

    return rows


def search_rows(table, column, query_text, limit=50):
    """Case-insensitive partial match search on a single column."""
    _check_configured()
    with httpx.Client(timeout=30) as client:
        response = client.get(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers=_headers(),
            params={"select": "*", column: f"ilike.*{query_text}*", "limit": limit},
        )
        response.raise_for_status()
        return response.json()


def insert_row(table, data):
    _check_configured()
    with httpx.Client(timeout=30) as client:
        response = client.post(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers=_headers(),
            json=data,
        )
        response.raise_for_status()
        return response.json()


def update_row(table, match_column, match_value, data):
    _check_configured()
    with httpx.Client(timeout=30) as client:
        response = client.patch(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers=_headers(),
            params={match_column: f"eq.{match_value}"},
            json=data,
        )
        response.raise_for_status()
        return response.json()


def delete_row(table, match_column, match_value):
    _check_configured()
    with httpx.Client(timeout=30) as client:
        response = client.delete(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers=_headers(),
            params={match_column: f"eq.{match_value}"},
        )
        response.raise_for_status()
        return True
