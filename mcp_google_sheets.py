#!/usr/bin/env python3
"""
MCP Server for Google Sheets (Veles Dashboard)
Reads data from the Google Sheets spreadsheet using a service account.

Required: service_account.json in the project directory
Spreadsheet ID: 13Xrmh-cfWFoR3-9TGFq2OoYZGmKxDhUhfqXWXC_SMKM
"""

import json
import os
import sys
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from mcp.server.fastmcp import FastMCP

SPREADSHEET_ID = os.environ.get(
    "SPREADSHEET_ID", "13Xrmh-cfWFoR3-9TGFq2OoYZGmKxDhUhfqXWXC_SMKM"
)
SA_FILE = os.environ.get(
    "GOOGLE_SERVICE_ACCOUNT_FILE",
    str(Path(__file__).parent / "service_account.json"),
)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

mcp = FastMCP("google-sheets-veles")


def _get_service():
    if not Path(SA_FILE).exists():
        raise FileNotFoundError(
            f"Service account file not found: {SA_FILE}\n"
            "Please place service_account.json in the project directory."
        )
    creds = service_account.Credentials.from_service_account_file(SA_FILE, scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


@mcp.tool()
def list_sheets() -> str:
    """List all sheet names in the Veles Dashboard spreadsheet."""
    service = _get_service()
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = [s["properties"]["title"] for s in meta["sheets"]]
    return json.dumps({"sheets": sheets, "spreadsheet_id": SPREADSHEET_ID})


@mcp.tool()
def read_sheet(sheet_name: str, range_a1: str = "A1:Z1000") -> str:
    """
    Read data from a specific sheet.

    Args:
        sheet_name: Sheet tab name, e.g. 'РЕЙСЫ', 'МАРШРУТЫ', 'API_TRIPS'
        range_a1: A1 notation range, default reads up to 1000 rows
    """
    service = _get_service()
    full_range = f"'{sheet_name}'!{range_a1}"
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SPREADSHEET_ID, range=full_range)
        .execute()
    )
    values = result.get("values", [])
    if not values:
        return json.dumps({"sheet": sheet_name, "rows": 0, "data": []})

    headers = values[0] if values else []
    rows = []
    for row in values[1:]:
        # Pad short rows with empty strings
        padded = row + [""] * (len(headers) - len(row))
        rows.append(dict(zip(headers, padded)))

    return json.dumps(
        {
            "sheet": sheet_name,
            "rows": len(rows),
            "headers": headers,
            "data": rows,
        },
        ensure_ascii=False,
        indent=2,
    )


@mcp.tool()
def read_trips() -> str:
    """Read all trips (РЕЙСЫ sheet) from the Veles Dashboard spreadsheet."""
    return read_sheet("РЕЙСЫ")


@mcp.tool()
def read_routes() -> str:
    """Read all routes (МАРШРУТЫ sheet) from the Veles Dashboard spreadsheet."""
    return read_sheet("МАРШРУТЫ")


@mcp.tool()
def read_api_trips() -> str:
    """Read API_TRIPS sheet (machine-readable trip data) from the Veles Dashboard spreadsheet."""
    return read_sheet("API_TRIPS")


@mcp.tool()
def get_trip(trip_id: str) -> str:
    """
    Find a specific trip by its ID (e.g. 'RSH-01-0315-01').

    Args:
        trip_id: Trip code like RSH-01-0315-01
    """
    data_json = json.loads(read_sheet("РЕЙСЫ"))
    rows = data_json.get("data", [])
    for row in rows:
        if row.get("Код рейса") == trip_id:
            return json.dumps(row, ensure_ascii=False, indent=2)
    return json.dumps({"error": f"Trip '{trip_id}' not found"})


if __name__ == "__main__":
    mcp.run()
