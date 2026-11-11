#!/usr/bin/env python3
"""Mist account transfer - invite a new email as Super Observer everywhere
your current accounts are Super User.

For every API key in mist_account_transfer.ini (one or more per Mist cloud):

  1. Finds every org the key can see (from /self - MSP access is expanded to
     every org in the MSP).
  2. Orgs where the key is Super User (role 'admin', org or MSP level):
       * checks the org's admin list for the new email
       * if it is already an admin, or already has a pending invite, nothing
         is sent and its current level is reported
       * otherwise an invite is sent for Super Observer (role 'read' with
         the 'super_observer' view)
  3. Orgs where no key is Super User are listed with the access level the
     key does have.

Everything is shown on screen and saved to an Excel workbook (created, or
updated if it already exists) with two sheets:

  * Invitations     - one row per org (and email): Invited / Already admin /
                      Already invited, plus the email's existing level
  * Not Super User  - org ID, org name and the access level the key has

Nothing is sent until you confirm the list (use --dry-run to only look, or
--yes to skip the question). If any key or org fails while checking, the
script stops and lists the failures before sending anything.
"""
import argparse
import configparser
import os
import re
import sys
from datetime import datetime
from urllib.parse import urlparse

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INI = "mist_account_transfer.ini"


def build_parser():
    parser = argparse.ArgumentParser(
        description="Invite an email as Super Observer to every org your API keys are Super User on")
    parser.add_argument("--dry-run", action="store_true",
                        help="show what would be done and write the Excel file, but send no invites")
    parser.add_argument("--yes", action="store_true",
                        help="send the invites without asking for confirmation")
    parser.add_argument("--ini", metavar="PATH",
                        help=f"settings file to use (default: {DEFAULT_INI} next to the script); "
                             f"a relative output_file is saved next to it")
    return parser


ARGS = build_parser().parse_args()  # handles --help before any settings are read


def ini_path_from_args():
    """--ini PATH picks another settings file (default: the one next to the
    script). Read before anything else, since the settings are loaded at start."""
    if ARGS.ini:
        return os.path.abspath(os.path.expanduser(ARGS.ini))
    return os.path.join(SCRIPT_DIR, DEFAULT_INI)


INI_PATH = ini_path_from_args()
INI_DIR = os.path.dirname(INI_PATH)
INI_NAME = os.path.basename(INI_PATH)
config = configparser.ConfigParser()
config.optionxform = str  # keep cloud names exactly as written
if not config.read(INI_PATH):
    if INI_PATH == os.path.join(SCRIPT_DIR, DEFAULT_INI):
        sys.exit(f"{DEFAULT_INI} not found next to the script - copy {DEFAULT_INI}.example "
                 f"and fill it in, or use --ini PATH.")
    sys.exit(f"Settings file not found: {INI_PATH}")

REQUEST_TIMEOUT = 60
PAGE_LIMIT = 1000
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
STATUS_FILLS = {
    "Would invite": PatternFill("solid", fgColor="E2EFDA"),
    "Invited": PatternFill("solid", fgColor="C6EFCE"),
    "Already admin": PatternFill("solid", fgColor="DDEBF7"),
    "Already invited": PatternFill("solid", fgColor="FFF2CC"),
}
RISK_FILL = PatternFill("solid", fgColor="FFC7CE")
MSP_CHECK_FILL = PatternFill("solid", fgColor="FFEB9C")
SOLE_RISK = "ONLY SUPER USER - removing your account(s) deletes the org"
# Every at-risk Delete Risk text starts with "ONLY".
MSP_OTHERS_RISK = ("ONLY LOCAL SUPER USER - org is in MSP '{msp}', which has {n} other Super User(s): "
                   "check with them before deleting your account(s)")
MSP_NONE_RISK = ("ONLY SUPER USER - org is in MSP '{msp}' but the MSP has no other Super Users: "
                 "removing your account(s) deletes the org")
MSP_UNKNOWN_RISK = ("ONLY LOCAL SUPER USER - org is in MSP '{msp}': check the MSP's admins "
                    "before deleting your account(s), the org may not be deleted")


def is_at_risk(row):
    return str(row.get("delete_risk", "")).startswith("ONLY")

# What we invite as: Super Observer = read role + super_observer UI view.
INVITE_PRIVILEGE = {"scope": "org", "role": "read", "views": ["super_observer"]}

ROLE_NAMES = {
    "admin": "Super User",
    "write": "Network Admin",
    "read": "Observer",
    "helpdesk": "Helpdesk",
    "installer": "Installer",
}
VIEW_NAMES = {
    "super_observer": "Super Observer",
    "reporting": "Reporting",
    "marketing": "Marketing",
    "location": "Location",
    "security": "Security",
    "switch_admin": "Switch Admin",
    "mxedge_admin": "Mist Edge Admin",
    "lobby_admin": "Lobby Admin",
}
ROLE_RANK = {"admin": 5, "write": 4, "helpdesk": 3, "read": 2, "installer": 1}

INVITE_COLUMNS = [
    ("Cloud", "cloud"),
    ("Org ID", "org_id"),
    ("Org Name", "org_name"),
    ("Email", "email"),
    ("Status", "status"),
    ("Level", "level"),
    ("Delete Risk", "delete_risk"),
    ("Other Super Users", "other_supers"),
    ("MSP", "msp"),
    ("MSP Super Users", "msp_supers"),
    ("Inventory", "inventory"),
    ("Licences", "licences"),
    ("Invited On", "invited_on"),
    ("Last Checked", "last_checked"),
    ("Checked With", "key_owner"),
]
NO_SUPER_COLUMNS = [
    ("Cloud", "cloud"),
    ("Org ID", "org_id"),
    ("Org Name", "org_name"),
    ("Access Level", "access"),
    ("Key Owner", "key_owner"),
    ("Last Checked", "last_checked"),
]
INVITES_SHEET = "Invitations"
NO_SUPER_SHEET = "Not Super User"
INVENTORY_SHEET = "Inventory to Release"
INVENTORY_COLUMNS = [
    ("Cloud", "cloud"),
    ("Org ID", "org_id"),
    ("Org Name", "org_name"),
    ("Type", "type"),
    ("Model", "model"),
    ("Serial", "serial"),
    ("MAC", "mac"),
    ("Name", "name"),
    ("Site", "site"),
    ("Connected", "connected"),
    ("VC MAC", "vc_mac"),
    ("Last Checked", "last_checked"),
]
LICENCE_SHEET = "Licences to Remove"
LICENCE_COLUMNS = [
    ("Cloud", "cloud"),
    ("Org ID", "org_id"),
    ("Org Name", "org_name"),
    ("Kind", "kind"),
    ("Status", "status"),
    ("Subscription ID", "subscription_id"),
    ("Type", "type"),
    ("Order ID", "order_id"),
    ("Quantity", "quantity"),
    ("Remaining", "remaining"),
    ("Start", "start"),
    ("End", "end"),
    ("Amendment ID", "amendment_id"),
    ("Last Checked", "last_checked"),
]
TYPE_NAMES = {"ap": "AP", "switch": "switch", "gateway": "gateway", "mxedge": "Mist Edge"}


class ApiError(Exception):
    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status


# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------
def setting(name, required=True, fallback=""):
    value = config.get("settings", name, fallback=fallback).strip()
    if required and (not value or value.startswith("YOUR_")):
        sys.exit(f"Set '{name}' under [settings] in {INI_NAME}.")
    return value


INVITE_EMAIL = setting("invite_email").lower()
if not EMAIL_RE.match(INVITE_EMAIL):
    sys.exit(f"invite_email '{INVITE_EMAIL}' in {INI_NAME} doesn't look like an email address.")
FIRST_NAME = setting("first_name")
LAST_NAME = setting("last_name")
try:
    INVITE_HOURS = int(setting("invite_hours", required=False, fallback="168") or 168)
except ValueError:
    sys.exit(f"invite_hours in {INI_NAME} must be a number of hours (1-168).")
if not 1 <= INVITE_HOURS <= 168:
    sys.exit(f"invite_hours in {INI_NAME} must be between 1 and 168 (Mist's limit is 7 days).")
# A relative output_file is relative to the ini file's folder.
OUTPUT_FILE = os.path.join(INI_DIR, os.path.expanduser(
    setting("output_file", required=False, fallback="mist_account_transfer.xlsx")
    or "mist_account_transfer.xlsx"))
EXCLUDE_ORGS = {x.strip().lower() for x in
                config.get("settings", "exclude_orgs", fallback="").split(",") if x.strip()}


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------
def build_session():
    """Shared HTTP session, wired up for a TLS-inspecting proxy (e.g. Zscaler)
    when the optional [network] section is present in the .ini. With no
    [network] section, behaviour is unchanged (certifi CA store, and the
    HTTP_PROXY / HTTPS_PROXY / NO_PROXY environment variables if set)."""
    session = requests.Session()
    if "network" not in config:
        return session

    section = config["network"]
    ca_bundle = section.get("ca_bundle", "").strip()
    if ca_bundle:
        ca_path = os.path.expanduser(ca_bundle)
        if not os.path.isfile(ca_path):
            sys.exit(f"[network] ca_bundle does not exist: {ca_path}\n"
                     f"This should be a PEM file containing your proxy's root CA "
                     f"certificate (e.g. exported from Zscaler).")
        session.verify = ca_path
    elif not section.getboolean("verify_ssl", fallback=True):
        session.verify = False
        print("WARNING: TLS certificate verification is DISABLED ([network] verify_ssl = false). "
              "Only use this as a last resort.")
        requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)

    for scheme in ("http", "https"):
        proxy = section.get(f"{scheme}_proxy", "").strip()
        if proxy:
            session.proxies[scheme] = proxy
    return session


SESSION = build_session()


def cloud_to_host(cloud):
    """Normalise an .ini cloud value to an API host name like api.eu.mist.com."""
    value = cloud.lower().strip()
    if "://" in value:
        value = urlparse(value).netloc
    value = value.strip("/")
    if value.endswith("mist.com"):
        if value.startswith("manage."):
            value = "api." + value[len("manage."):]
        return value
    if value in ("us", "global", "global01"):
        return "api.mist.com"
    return f"api.{value}.mist.com"


def api_call(method, host, token, path, params=None, body=None):
    """Call https://{host}/api/v1{path}. Returns the Response, or raises ApiError."""
    url = f"https://{host}/api/v1{path}"
    try:
        resp = SESSION.request(method, url, headers={"Authorization": f"Token {token}"},
                               params=params, json=body, timeout=REQUEST_TIMEOUT)
    except requests.exceptions.SSLError as e:
        raise ApiError(f"TLS error ({e}) - if you're behind Zscaler or a similar proxy, "
                       f"set [network] ca_bundle in {INI_NAME}")
    except requests.exceptions.ProxyError as e:
        raise ApiError(f"proxy error ({e}) - check [network] http_proxy / https_proxy in {INI_NAME}")
    except requests.exceptions.RequestException as e:
        raise ApiError(f"could not connect ({e.__class__.__name__})")

    if resp.status_code == 401:
        raise ApiError("HTTP 401 - token not valid on this cloud")
    if resp.status_code == 403:
        raise ApiError(f"HTTP 403 - token has no access to {path}", status=403)
    if resp.status_code == 429:
        raise ApiError("HTTP 429 - API rate limit reached (5000 calls/hour), try again later")
    if not resp.ok:
        raise ApiError(f"HTTP {resp.status_code} on {method} {path}: {resp.text[:200]}")
    return resp


def api_get(host, token, path, params=None):
    return api_call("GET", host, token, path, params=params)


def api_get_all(host, token, path, params=None):
    """Page through a list endpoint (limit/page) and return every item."""
    items, page = [], 1
    while True:
        p = dict(params or {}, limit=PAGE_LIMIT, page=page)
        resp = api_get(host, token, path, p)
        batch = resp.json()
        if not isinstance(batch, list):
            raise ApiError(f"unexpected response from {path} (not a list)")
        items.extend(batch)
        total = resp.headers.get("X-Page-Total")
        if len(batch) < PAGE_LIMIT or (total is not None and len(items) >= int(total)):
            return items
        page += 1


# --------------------------------------------------------------------------
# Keys, orgs and access levels
# --------------------------------------------------------------------------
def load_keys():
    """Return [(api_host, token)] from [clouds]. A cloud can hold several
    comma-separated tokens (e.g. your Juniper and your Mist account)."""
    if "clouds" not in config or not config["clouds"]:
        sys.exit(f"No clouds listed - add 'cloud = api_token' lines under [clouds] in {INI_NAME}.")
    keys = []
    for cloud, value in config["clouds"].items():
        tokens = [t.strip() for t in value.split(",") if t.strip()]
        if not tokens:
            continue  # cloud listed but left blank - skip it
        if any(t.startswith("YOUR_") for t in tokens):
            sys.exit(f"Placeholder token left for '{cloud}' in {INI_NAME} - fill it in or blank it.")
        host = cloud_to_host(cloud)
        for token in tokens:
            if (host, token) in keys:
                sys.exit(f"The same token is listed twice for {host} in {INI_NAME}.")
            keys.append((host, token))
    if not keys:
        sys.exit(f"No API tokens set under [clouds] in {INI_NAME}.")
    return keys


def key_label(host, token, owner=None):
    who = f" ({owner})" if owner else ""
    return f"{host} key ...{token[-4:]}{who}"


def role_name(priv):
    """Human name for one privilege, e.g. 'Super Observer' or 'Observer'."""
    role = priv.get("role", "")
    views = priv.get("views") or []
    if views:
        names = [VIEW_NAMES.get(v, v) for v in views]
        return " + ".join(names)
    return ROLE_NAMES.get(role, role or "unknown")


def describe_privileges(privs):
    """'Super User (MSP: Acme)', 'Observer (site: London)' ... joined by '; '."""
    parts = []
    for p in privs:
        scope = p.get("scope", "")
        name = role_name(p)
        if scope == "org":
            parts.append(f"{name} (org)")
        elif scope == "msp":
            parts.append(f"{name} (MSP{': ' + p['msp_name'] if p.get('msp_name') else ''})")
        elif scope == "site":
            parts.append(f"{name} (site{': ' + p['name'] if p.get('name') else ''})")
        elif scope == "sitegroup":
            parts.append(f"{name} (site group)")
        elif scope == "orggroup":
            parts.append(f"{name} (org group)")
        else:
            parts.append(f"{name} ({scope or 'unknown scope'})")
    return "; ".join(dict.fromkeys(parts)) or "none"


def discover_orgs(host, token, self_data):
    """Return {org_id: {"name", "privs": [privileges that grant access]}}
    for one key, from /self. MSP privileges are expanded to every org in
    the MSP (the MSP privilege is recorded against each org)."""
    orgs = {}

    def add(org_id, name, priv, msp_id=None, msp_name=None):
        entry = orgs.setdefault(org_id, {"name": "", "privs": [], "msp_id": "", "msp_name": ""})
        entry["name"] = entry["name"] or name or ""
        entry["msp_id"] = entry["msp_id"] or msp_id or ""
        entry["msp_name"] = entry["msp_name"] or msp_name or ""
        entry["privs"].append(priv)

    for priv in self_data.get("privileges", []):
        scope = priv.get("scope")
        if scope == "org" and priv.get("org_id"):
            add(priv["org_id"], priv.get("name") or priv.get("org_name"), priv,
                priv.get("msp_id"), priv.get("msp_name"))
        elif scope == "msp" and priv.get("msp_id"):
            msp_priv = dict(priv, msp_name=priv.get("name") or priv.get("msp_name"))
            for org in api_get_all(host, token, f"/msps/{priv['msp_id']}/orgs"):
                add(org["id"], org.get("name"), msp_priv, priv["msp_id"], msp_priv["msp_name"])
        elif scope in ("site", "sitegroup") and priv.get("org_id"):
            add(priv["org_id"], priv.get("org_name"), priv)
    return orgs


def is_super_user(privs):
    return any(p.get("role") == "admin" and p.get("scope") in ("org", "msp") for p in privs)


def msp_risk(host, org_id, entry, org_token, msp_keys, cache, own_emails):
    """Delete Risk / MSP / MSP Super Users for an org with no other local
    Super User. Reads the MSP's admins when a key is Super User on that MSP."""
    msp_id = entry["msp_id"]
    if not msp_id:  # not known from /self - ask the org itself
        msp_id = api_get(host, org_token, f"/orgs/{org_id}").json().get("msp_id") or ""
        entry["msp_id"] = msp_id
    if not msp_id:
        return {"delete_risk": SOLE_RISK, "msp": "", "msp_supers": ""}
    msp = entry["msp_name"] or msp_id
    key = msp_keys.get((host, msp_id))
    if not key:
        return {"delete_risk": MSP_UNKNOWN_RISK.format(msp=msp), "msp": msp,
                "msp_supers": "(can't read - no key is MSP Super User)"}
    if (host, msp_id) not in cache:
        try:
            cache[(host, msp_id)] = api_get_all(host, key[0], f"/msps/{msp_id}/admins")
        except ApiError as e:
            if e.status != 403:
                raise
            cache[(host, msp_id)] = None
    admins = cache[(host, msp_id)]
    if admins is None:
        return {"delete_risk": MSP_UNKNOWN_RISK.format(msp=msp), "msp": msp,
                "msp_supers": "(can't read - HTTP 403)"}
    # The MSP's admin list also holds org-level admins of its orgs - only
    # MSP-level Super Users keep an org alive across the MSP.
    others = other_super_users(admins, own_emails, scopes=("msp",))
    if others:
        return {"delete_risk": MSP_OTHERS_RISK.format(msp=msp, n=len(others)), "msp": msp,
                "msp_supers": ", ".join(others)}
    return {"delete_risk": MSP_NONE_RISK.format(msp=msp), "msp": msp, "msp_supers": "none"}


def find_email(admins, email):
    """Return the admin-list entry for email (admin or pending invite), or None."""
    for a in admins:
        if (a.get("email") or "").strip().lower() == email:
            return a
    return None


def is_pending_invite(entry):
    # Pending invites appear in the org's admin list with an expiry time.
    return bool(entry.get("expire_time"))


def org_inventory(host, token, org_id, org_name, now):
    """Every device claimed to the org (all types, Virtual Chassis members
    listed one by one), as Inventory to Release rows. These have to be
    released before Mist will delete the org."""
    devices = api_get_all(host, token, f"/orgs/{org_id}/inventory", {"vc": "true"})
    sites = {}
    if any(d.get("site_id") for d in devices):
        sites = {s["id"]: s.get("name", "") for s in api_get_all(host, token, f"/orgs/{org_id}/sites")}
    rows = []
    for d in devices:
        site_id = d.get("site_id")
        rows.append({
            "cloud": host, "org_id": org_id, "org_name": org_name,
            "type": TYPE_NAMES.get(d.get("type"), d.get("type") or ""),
            "model": d.get("model") or "", "serial": d.get("serial") or "",
            "mac": d.get("mac") or "", "name": d.get("name") or d.get("hostname") or "",
            "site": (sites.get(site_id) or site_id) if site_id else "(unassigned)",
            "connected": "yes" if d.get("connected") else "no",
            "vc_mac": d.get("vc_mac") or "", "last_checked": now,
        })
    rows.sort(key=lambda r: (r["type"], r["model"], r["serial"]))
    return rows


def fmt_date(epoch):
    return datetime.fromtimestamp(epoch).strftime("%Y-%m-%d") if epoch else ""


def licence_status(start, end, now_epoch):
    if end and end < now_epoch:
        return "expired"
    if start and start > now_epoch:
        return "not started"
    return "active"


def org_licences(host, token, org_id, org_name, now):
    """Every subscription on the org, plus licence moves (amendments) in or
    out of it, as Licences to Remove rows."""
    data = api_get(host, token, f"/orgs/{org_id}/licenses").json()
    now_epoch = datetime.now().timestamp()
    rows = []
    base = {"cloud": host, "org_id": org_id, "org_name": org_name, "last_checked": now}
    for lic in data.get("licenses") or []:
        start, end = lic.get("start_time") or 0, lic.get("end_time") or 0
        sub_id = lic.get("subscription_id") or ""
        rows.append(dict(base,
            kind="Evaluation" if "eval" in sub_id.lower() else "Subscription",
            status=licence_status(start, end, now_epoch),
            subscription_id=sub_id, type=lic.get("type") or "",
            order_id=lic.get("order_id") or "", quantity=lic.get("quantity"),
            remaining=lic.get("remaining_quantity"),
            start=fmt_date(start), end=fmt_date(end), amendment_id=""))
    for am in data.get("amendments") or []:
        start, end = am.get("start_time") or 0, am.get("end_time") or 0
        qty = am.get("quantity") or 0
        rows.append(dict(base,
            kind="Moved in" if qty > 0 else "Moved out",
            status=licence_status(start, end, now_epoch),
            subscription_id=am.get("subscription_id") or "", type=am.get("type") or "",
            order_id="", quantity=qty, remaining="",
            start=fmt_date(start), end=fmt_date(end), amendment_id=am.get("id") or ""))
    rows.sort(key=lambda r: (r["kind"], r["status"], r["type"], r["subscription_id"]))
    return rows


def licence_summary(rows):
    """'4 active (SUB-MAN x19, SUB-VNA x19, ...), 1 expired' or 'none'."""
    if not rows:
        return "none"
    active = [r for r in rows if r["status"] == "active"]
    parts = []
    if active:
        parts.append(f"{len(active)} active (" +
                     ", ".join(f"{r['type']} x{abs(r['quantity'] or 0)}" + (f" {r['kind'].lower()}"
                               if r["kind"].startswith("Moved") else "") for r in active) + ")")
    for status in ("not started", "expired"):
        n = sum(r["status"] == status for r in rows)
        if n:
            parts.append(f"{n} {status}")
    return ", ".join(parts)


def inventory_summary(rows):
    """'13 (2 AP, 6 switch, 4 gateway, 1 Mist Edge)' or 'none'."""
    if not rows:
        return "none"
    counts = {}
    for r in rows:
        counts[r["type"]] = counts.get(r["type"], 0) + 1
    return f"{len(rows)} (" + ", ".join(f"{n} {t}" for t, n in counts.items()) + ")"


def other_super_users(admins, own_emails, scopes=("org", "msp")):
    """Emails of accepted Super Users (role admin, org or MSP scope) on the
    org, other than the key owners. Pending invites don't count - they can
    still expire unused."""
    others = []
    for a in admins:
        email = (a.get("email") or "").strip().lower()
        if not email or email in own_emails or is_pending_invite(a):
            continue
        if any(p.get("role") == "admin" and p.get("scope") in scopes
               for p in a.get("privileges") or []):
            others.append(email)
    return sorted(set(others))


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def open_workbook():
    if os.path.exists(OUTPUT_FILE):
        try:
            return load_workbook(OUTPUT_FILE), True
        except Exception as e:
            sys.exit(f"Could not open {OUTPUT_FILE} ({e}) - is it open in Excel?")
    wb = Workbook()
    wb.remove(wb.active)
    return wb, False


def read_sheet(wb, title, columns):
    if title not in wb.sheetnames:
        return []
    ws = wb[title]
    header_to_key = dict(columns)
    headers = [c.value for c in ws[1]]
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not any(values):
            continue
        rows.append({header_to_key[h]: v for h, v in zip(headers, values) if h in header_to_key})
    return rows


def write_sheet(wb, title, columns, rows, status_key=None):
    if title in wb.sheetnames:
        index = wb.sheetnames.index(title)
        wb.remove(wb[title])
        ws = wb.create_sheet(title, index)
    else:
        ws = wb.create_sheet(title)
    headers = [h for h, _ in columns]
    ws.append(headers)
    for cell in ws[1]:
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
    keys = [k for _, k in columns]
    for r in rows:
        ws.append([r.get(k, "") for _, k in columns])
        if "delete_risk" in keys and is_at_risk(r):
            fill = RISK_FILL if str(r["delete_risk"]).startswith("ONLY SUPER USER") else MSP_CHECK_FILL
            ws.cell(row=ws.max_row, column=keys.index("delete_risk") + 1).fill = fill
        if status_key:
            fill = STATUS_FILLS.get(str(r.get(status_key, "")).split(" (")[0])
            if fill:
                ws.cell(row=ws.max_row, column=[k for _, k in columns].index(status_key) + 1).fill = fill
    for i, (h, k) in enumerate(columns, 1):
        width = max([len(h)] + [len(str(r.get(k) or "")) for r in rows])
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 70)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def merge_rows(old, new, key_fields, keep_fields=()):
    """Replace old rows that match a new row on key_fields; keep the rest.
    keep_fields are carried over from the old row when the new one is blank."""
    by_key = {tuple(str(r.get(k, "")) for k in key_fields): r for r in old}
    for r in new:
        k = tuple(str(r.get(f, "")) for f in key_fields)
        prev = by_key.get(k)
        if prev:
            for f in keep_fields:
                if not r.get(f) and prev.get(f):
                    r[f] = prev[f]
        by_key[k] = r
    return sorted(by_key.values(), key=lambda r: (str(r.get("cloud", "")),
                                                  str(r.get("org_name", "")).lower()))


def save_report(invite_rows, no_super_rows, inventory_rows, licence_rows, checked_orgs, interactive):
    wb, existed = open_workbook()
    invites = merge_rows(read_sheet(wb, INVITES_SHEET, INVITE_COLUMNS), invite_rows,
                         ("cloud", "org_id", "email"), keep_fields=("invited_on",))
    no_super = merge_rows(read_sheet(wb, NO_SUPER_SHEET, NO_SUPER_COLUMNS), no_super_rows,
                          ("cloud", "org_id"))
    # An org that is now Super User for some key no longer belongs on the
    # Not Super User sheet.
    super_orgs = {(r["cloud"], r["org_id"]) for r in invite_rows}
    no_super = [r for r in no_super if (r.get("cloud"), r.get("org_id")) not in super_orgs]

    # Inventory is replaced for every org checked this run, so released
    # devices drop off, and orgs no longer at risk lose their rows.
    inventory = [r for r in read_sheet(wb, INVENTORY_SHEET, INVENTORY_COLUMNS)
                 if (r.get("cloud"), r.get("org_id")) not in checked_orgs] + inventory_rows
    inventory.sort(key=lambda r: (str(r.get("cloud", "")), str(r.get("org_name", "")).lower(),
                                  str(r.get("type", "")), str(r.get("serial", ""))))

    write_sheet(wb, INVITES_SHEET, INVITE_COLUMNS, invites, status_key="status")
    write_sheet(wb, NO_SUPER_SHEET, NO_SUPER_COLUMNS, no_super)
    write_sheet(wb, INVENTORY_SHEET, INVENTORY_COLUMNS, inventory)
    licences = [r for r in read_sheet(wb, LICENCE_SHEET, LICENCE_COLUMNS)
                if (r.get("cloud"), r.get("org_id")) not in checked_orgs] + licence_rows
    licences.sort(key=lambda r: (str(r.get("cloud", "")), str(r.get("org_name", "")).lower(),
                                 str(r.get("kind", "")), str(r.get("type", ""))))
    write_sheet(wb, LICENCE_SHEET, LICENCE_COLUMNS, licences)
    while True:
        try:
            wb.save(OUTPUT_FILE)
            break
        except PermissionError:
            if not interactive:
                sys.exit(f"Could not save {OUTPUT_FILE} - close it in Excel and run again.")
            input(f"Could not save {OUTPUT_FILE} - close it in Excel, then press Enter to retry "
                  f"(Ctrl+C to give up): ")
    print(f"\n{'Updated' if existed else 'Created'} {OUTPUT_FILE} "
          f"({len(invites)} invitation row(s), {len(no_super)} not-super-user org(s), "
          f"{len(inventory)} device(s) to release, {len(licences)} licence row(s))")


# --------------------------------------------------------------------------
# Output helpers
# --------------------------------------------------------------------------
def print_table(title, columns, rows):
    print(f"\n{title}")
    if not rows:
        print("  (none)")
        return
    widths = [max(len(h), *(len(str(r.get(k) or "")) for r in rows)) for h, k in columns]
    print("  " + "  ".join(h.ljust(w) for (h, _), w in zip(columns, widths)))
    print("  " + "  ".join("-" * w for w in widths))
    for r in rows:
        print("  " + "  ".join(str(r.get(k) or "").ljust(w) for (_, k), w in zip(columns, widths)))


def fail(title, failures, note="nothing sent, nothing written"):
    print(f"\n{title} - stopping, {note}.\n")
    width = max(len(w) for w, _ in failures)
    for where, reason in failures:
        print(f"  {where.ljust(width)}  {reason}")
    sys.exit(1)


# --------------------------------------------------------------------------
def main():
    args = ARGS
    interactive = sys.stdin.isatty()

    keys = load_keys()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"Mist account transfer - {now}")
    print(f"  Settings: {INI_PATH}")
    print(f"  New account: {INVITE_EMAIL} ({FIRST_NAME} {LAST_NAME}) as Super Observer, "
          f"invite valid {INVITE_HOURS}h")
    print(f"  {len(keys)} API key(s){' - DRY RUN, nothing will be sent' if args.dry_run else ''}")

    # 1. Check every key first, so a bad key stops the run before any work.
    failures, checked = [], []
    for host, token in keys:
        try:
            self_data = api_get(host, token, "/self").json()
            checked.append((host, token, self_data))
        except ApiError as e:
            failures.append((key_label(host, token), str(e)))
    if failures:
        fail("Some API keys failed", failures)

    # 2. Orgs per key. The same org can be seen by two keys on one cloud
    #    (e.g. Juniper and Mist accounts); a key with Super User wins.
    orgs = {}  # (host, org_id) -> {"name", "super": [(token, owner)], "access": [(owner, privs)], "msp_id", "msp_name"}
    msp_keys = {}  # (host, msp_id) -> (token, owner) of a key that is MSP Super User
    msp_admin_cache = {}
    for host, token, self_data in checked:
        owner = self_data.get("email") or self_data.get("name") or f"...{token[-4:]}"
        label = key_label(host, token, owner)
        if owner.strip().lower() == INVITE_EMAIL:
            fail("Config problem", [(label, f"this key belongs to {INVITE_EMAIL} itself - "
                                            f"use the keys of the accounts you're moving from")])
        try:
            found = discover_orgs(host, token, self_data)
        except ApiError as e:
            failures.append((label, f"listing orgs: {e}"))
            continue
        n_super = sum(is_super_user(o["privs"]) for o in found.values())
        print(f"  {label}: {len(found)} org(s), Super User on {n_super}")
        for org_id, info in found.items():
            entry = orgs.setdefault((host, org_id), {"name": "", "super": [], "access": [],
                                                     "msp_id": "", "msp_name": ""})
            entry["name"] = entry["name"] or info["name"]
            entry["msp_id"] = entry["msp_id"] or info["msp_id"]
            entry["msp_name"] = entry["msp_name"] or info["msp_name"]
            for p in self_data.get("privileges", []):  # keys that can read MSP admins
                if p.get("scope") == "msp" and p.get("role") == "admin" and p.get("msp_id"):
                    msp_keys.setdefault((host, p["msp_id"]), (token, owner))
            entry["access"].append((owner, info["privs"]))
            if is_super_user(info["privs"]):
                entry["super"].append((token, owner))
    if failures:
        fail("Org discovery failed", failures)
    if not orgs:
        sys.exit("No orgs found for any API key.")

    # Fill in any missing org names (e.g. site-only access) and apply exclusions.
    for (host, org_id), entry in list(orgs.items()):
        if org_id.lower() in EXCLUDE_ORGS or entry["name"].lower() in EXCLUDE_ORGS:
            print(f"    excluded: {entry['name'] or org_id}")
            del orgs[(host, org_id)]
            continue
        if not entry["name"] and entry["super"]:
            try:
                entry["name"] = api_get(host, entry["super"][0][0], f"/orgs/{org_id}").json().get("name", "")
            except ApiError:
                pass

    # Every key owner's account is assumed to be going away, so none of them
    # count as "another" Super User when checking for sole-Super-User orgs.
    own_emails = {(sd.get("email") or "").strip().lower() for _, _, sd in checked} - {""}

    # 3. For Super User orgs, check the org's admins/invites for the email,
    #    and whether anyone besides the key owners is Super User.
    invite_rows, no_super_rows, to_invite, inventory_rows, licence_rows = [], [], [], [], []
    super_orgs = sorted(((h, o, e) for (h, o), e in orgs.items() if e["super"]),
                        key=lambda x: (x[0], x[2]["name"].lower()))
    print(f"\nChecking {len(super_orgs)} Super User org(s) for {INVITE_EMAIL} ...")
    for host, org_id, entry in super_orgs:
        # /self can say role 'admin' for an account that is really only an
        # org admin: Mist then refuses the admin list (HTTP 403). Such an org
        # is treated as Not Super User. Try each Super User key in turn.
        admins, refused, error = None, [], None
        for token, owner in entry["super"]:
            try:
                admins = api_get_all(host, token, f"/orgs/{org_id}/admins")
                break
            except ApiError as e:
                if e.status == 403:
                    refused.append((token, owner))
                    continue
                error = e
                break
        if error:
            failures.append((f"{host} / {entry['name'] or org_id}", str(error)))
            continue
        if admins is None:
            entry["super"] = []
            entry["admin_list_refused"] = True
            continue
        token, owner = next(k for k in entry["super"] if k not in refused)
        entry["super"] = [(token, owner)] + [k for k in entry["super"] if k not in refused and k != (token, owner)]
        row = {"cloud": host, "org_id": org_id, "org_name": entry["name"], "email": INVITE_EMAIL,
               "last_checked": now, "key_owner": owner}
        others = other_super_users(admins, own_emails)
        row["other_supers"] = ", ".join(others)
        if others:
            row["delete_risk"] = f"OK - {len(others)} other Super User(s)"
            row["inventory"] = row["licences"] = ""
            row["msp"] = entry["msp_name"] or entry["msp_id"]
            row["msp_supers"] = ""
        else:
            # The org would be deleted along with your account - and Mist only
            # deletes an org once its inventory has been released. If the org
            # is in an MSP, the MSP's own Super Users may keep it alive.
            try:
                row.update(msp_risk(host, org_id, entry, token, msp_keys, msp_admin_cache, own_emails))
            except ApiError as e:
                failures.append((f"{host} / {entry['name'] or org_id} (MSP)", str(e)))
                continue
            try:
                devices = org_inventory(host, token, org_id, entry["name"], now)
            except ApiError as e:
                failures.append((f"{host} / {entry['name'] or org_id} (inventory)", str(e)))
                continue
            inventory_rows.extend(devices)
            row["inventory"] = inventory_summary(devices)
            try:
                lic_rows = org_licences(host, token, org_id, entry["name"], now)
            except ApiError as e:
                failures.append((f"{host} / {entry['name'] or org_id} (licences)", str(e)))
                continue
            licence_rows.extend(lic_rows)
            row["licences"] = licence_summary(lic_rows)
        existing = find_email(admins, INVITE_EMAIL)
        if existing:
            row["status"] = "Already invited" if is_pending_invite(existing) else "Already admin"
            row["level"] = describe_privileges(existing.get("privileges") or [])
            if is_pending_invite(existing):
                expiry = datetime.fromtimestamp(existing["expire_time"]).strftime("%Y-%m-%d %H:%M")
                row["status"] += f" (expires {expiry})"
        else:
            row["status"] = "To invite"
            row["level"] = "Super Observer (org)"
            to_invite.append(row)
        invite_rows.append(row)
    if failures:
        fail("Some orgs could not be checked", failures)

    for (host, org_id), entry in orgs.items():
        if entry["super"]:
            continue
        access = "; ".join(f"{describe_privileges(p)}" + (f" [{o}]" if len(entry["access"]) > 1 else "")
                           for o, p in entry["access"])
        if entry.get("admin_list_refused"):
            access = ("Org Admin - listed as Super User but Mist refuses its admin list (HTTP 403); "
                      f"/self says: {access}")
        no_super_rows.append({
            "cloud": host, "org_id": org_id, "org_name": entry["name"],
            "access": access,
            "key_owner": ", ".join(dict.fromkeys(o for o, _ in entry["access"])),
            "last_checked": now,
        })
    no_super_rows.sort(key=lambda r: (r["cloud"], r["org_name"].lower()))

    table_cols = [("Cloud", "cloud"), ("Org Name", "org_name"), ("Org ID", "org_id"),
                  ("Status", "status"), ("Level", "level"), ("Delete Risk", "delete_risk")]
    print_table(f"Super User orgs ({len(invite_rows)})", table_cols, invite_rows)
    sole = [r for r in invite_rows if is_at_risk(r)]
    if sole:
        print(f"\n  WARNING: on {len(sole)} org(s) your account(s) are the only Super User in the org. Deleting")
        print(f"  them may delete the org. Add another Super User first (e.g. promote {INVITE_EMAIL}")
        print(f"  once the invite is accepted):")
        for r in sole:
            inv = (f"inventory to release: {r['inventory']}" if r["inventory"] != "none"
                   else "no inventory")
            lic = (f"licences: {r['licences']}" if r["licences"] != "none" else "no licences")
            print(f"    - {r['cloud']}  {r['org_name']}  ({r['org_id']})")
            if r.get("msp"):
                supers = r["msp_supers"]
                msp_line = (f"other MSP Super Users: {supers}" if supers and not supers.startswith("(")
                            and supers != "none" else f"MSP Super Users: {supers or 'unknown'}")
                print(f"        in MSP '{r['msp']}' - {msp_line}")
            print(f"        {inv}; {lic}")
        if inventory_rows or licence_rows:
            print(f"  Details are on the '{INVENTORY_SHEET}' and '{LICENCE_SHEET}' sheets "
                  f"of the workbook.")
    elif invite_rows:
        print(f"\n  Every Super User org has another Super User besides your key account(s).")
    print(f"  (Key accounts not counted: {', '.join(sorted(own_emails))})")
    print_table(f"Orgs without Super User ({len(no_super_rows)}) - no invite possible",
                [("Cloud", "cloud"), ("Org Name", "org_name"), ("Org ID", "org_id"),
                 ("Access Level", "access")], no_super_rows)

    checked_orgs = {(r["cloud"], r["org_id"]) for r in invite_rows}
    if args.dry_run:
        for r in to_invite:
            r["status"] = "Would invite (dry run)"
        save_report(invite_rows, no_super_rows, inventory_rows, licence_rows, checked_orgs, interactive)
        print(f"Dry run - {len(to_invite)} invite(s) would be sent. Nothing sent.")
        return

    # 4. Send the invites.
    if to_invite and not args.yes:
        if not interactive:
            sys.exit(f"\n{len(to_invite)} invite(s) to send - run with --yes to send without asking. "
                     f"Nothing sent, nothing written.")
        answer = input(f"\nSend {len(to_invite)} Super Observer invite(s) to {INVITE_EMAIL}? [y/N] ")
        if answer.strip().lower() not in ("y", "yes"):
            print("Not sent.")
            for r in to_invite:
                r["status"] = "Not sent (cancelled)"
            to_invite = []

    body = {"email": INVITE_EMAIL, "first_name": FIRST_NAME, "last_name": LAST_NAME,
            "hours": INVITE_HOURS, "privileges": [INVITE_PRIVILEGE]}
    for i, row in enumerate(to_invite, 1):
        token = orgs[(row["cloud"], row["org_id"])]["super"][0][0]
        print(f"  [{i}/{len(to_invite)}] inviting to {row['cloud']}  {row['org_name']} ... ", end="", flush=True)
        try:
            api_call("POST", row["cloud"], token, f"/orgs/{row['org_id']}/invites", body=body)
        except ApiError as e:
            print("FAILED")
            failures.append((f"{row['cloud']} / {row['org_name']}", str(e)))
            row["status"] = f"Invite failed: {e}"
            # Stop here - but record what was already sent.
            for r in to_invite[i:]:
                r["status"] = "Not sent (stopped after failure)"
            break
        print("ok")
        row["status"] = "Invited"
        row["invited_on"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Don't let unsent rows overwrite earlier history in the workbook.
    save_rows = [r for r in invite_rows if not r["status"].startswith("Not sent")]
    save_report(save_rows, no_super_rows, inventory_rows, licence_rows, checked_orgs, interactive)
    if failures:
        fail("An invite failed", failures, note="invites sent so far are recorded in the workbook")
    sent = sum(r["status"] == "Invited" for r in invite_rows)
    print(f"Done - {sent} invite(s) sent to {INVITE_EMAIL}.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nStopped.")
