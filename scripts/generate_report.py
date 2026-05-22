"""
LA Ops Hub — Weekly Report Generator
Fetches data from JIRA and SharePoint, outputs data/report.json
Run automatically by GitHub Actions every Monday at 9am.

Setup: add these secrets in GitHub → Settings → Secrets → Actions:
  JIRA_BASE_URL          e.g. https://yourcompany.atlassian.net
  JIRA_EMAIL             your JIRA login email
  JIRA_API_TOKEN         from https://id.atlassian.com/manage-profile/security/api-tokens
  JIRA_PROJECT_KEY       e.g. OPS
  SHAREPOINT_TENANT_ID   Azure AD tenant ID
  SHAREPOINT_CLIENT_ID   App registration client ID
  SHAREPOINT_CLIENT_SECRET  App registration secret
  SHAREPOINT_SITE_ID     SharePoint site ID
"""

import os
import json
import requests
from datetime import datetime, timedelta
from base64 import b64encode

# ── CONFIG ────────────────────────────────────────────────────────────────────
JIRA_BASE      = os.environ.get("JIRA_BASE_URL", "")
JIRA_EMAIL     = os.environ.get("JIRA_EMAIL", "")
JIRA_TOKEN     = os.environ.get("JIRA_API_TOKEN", "")
JIRA_PROJECT   = os.environ.get("JIRA_PROJECT_KEY", "OPS")
SP_TENANT      = os.environ.get("SHAREPOINT_TENANT_ID", "")
SP_CLIENT_ID   = os.environ.get("SHAREPOINT_CLIENT_ID", "")
SP_SECRET      = os.environ.get("SHAREPOINT_CLIENT_SECRET", "")
SP_SITE_ID     = os.environ.get("SHAREPOINT_SITE_ID", "")

today     = datetime.now()
week_ago  = (today - timedelta(days=7)).strftime("%Y-%m-%d")
today_str = today.strftime("%Y-%m-%d")
display   = today.strftime("%a %b %d, %Y — %I:%M %p")

# ── JIRA ─────────────────────────────────────────────────────────────────────
def jira_auth():
    token = b64encode(f"{JIRA_EMAIL}:{JIRA_TOKEN}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def jira_search(jql, fields="summary,status,assignee,comment,priority"):
    if not JIRA_BASE:
        return []
    url = f"{JIRA_BASE}/rest/api/3/search"
    params = {"jql": jql, "fields": fields, "maxResults": 50}
    try:
        r = requests.get(url, headers=jira_auth(), params=params, timeout=15)
        return r.json().get("issues", []) if r.ok else []
    except Exception as e:
        print(f"JIRA error: {e}")
        return []

def get_jira_data():
    open_issues   = jira_search(f'project={JIRA_PROJECT} AND status != Done ORDER BY updated DESC')
    closed_issues = jira_search(f'project={JIRA_PROJECT} AND status = Done AND updated >= "{week_ago}"')
    in_progress   = [i for i in open_issues if "progress" in i["fields"]["status"]["name"].lower()]
    blocked       = [i for i in open_issues if i["fields"]["priority"]["name"] == "Blocker"]

    sprint_name = "Current sprint"
    try:
        sp_url = f"{JIRA_BASE}/rest/agile/1.0/board"
        boards = requests.get(sp_url, headers=jira_auth(), timeout=10).json().get("values", [])
        if boards:
            board_id = next((b["id"] for b in boards if JIRA_PROJECT in b.get("name","")), boards[0]["id"])
            sprints = requests.get(f"{JIRA_BASE}/rest/agile/1.0/board/{board_id}/sprint?state=active", headers=jira_auth(), timeout=10).json().get("values", [])
            if sprints:
                sprint_name = sprints[0]["name"]
    except:
        pass

    projects = []
    seen = set()
    for issue in open_issues[:8]:
        f = issue["fields"]
        name = f["summary"]
        if name in seen: continue
        seen.add(name)
        status = f["status"]["name"]
        assignee = (f.get("assignee") or {}).get("displayName", "Unassigned")
        notes = ""
        comments = (f.get("comment") or {}).get("comments", [])
        if comments:
            notes = comments[-1]["body"]["content"][0]["content"][0]["text"][:120] + "…"
        status_map = {"In Progress": "In Progress", "To Do": "Todo", "Done": "Done", "Blocked": "Blocked"}
        pill = status_map.get(status, status)
        projects.append({"name": name, "owner": assignee.split(" ")[0] + " " + (assignee.split(" ")[-1][0] + ".") if " " in assignee else assignee, "status": pill, "notes": notes or "On track"})

    blockers = []
    for issue in blocked[:3]:
        f = issue["fields"]
        blockers.append({
            "project": f["summary"][:50],
            "description": f"Blocked — priority escalation needed. Assignee: {(f.get('assignee') or {}).get('displayName', 'Unassigned')}"
        })

    return {
        "sprint": sprint_name,
        "stats": [
            {"label": "Open tickets",     "value": str(len(open_issues)),   "note": f"Active in {JIRA_PROJECT}"},
            {"label": "Closed this week", "value": str(len(closed_issues)), "note": "Last 7 days"},
            {"label": "In progress",      "value": str(len(in_progress)),   "note": f"{len(blocked)} blocker(s)"},
            {"label": "Team capacity",    "value": "—",                     "note": "Update manually"}
        ],
        "projects": projects or [{"name": "No open tickets", "owner": "—", "status": "Done", "notes": "Great week!"}],
        "blockers": blockers or [{"project": "None", "description": "No blockers this week 🎉"}],
        "actions": [f"Review {len(open_issues)} open tickets in {JIRA_PROJECT}", "Update sprint board before standup", "Close any done items in JIRA"]
    }

# ── SHAREPOINT ────────────────────────────────────────────────────────────────
def get_sp_token():
    if not SP_TENANT:
        return None
    url = f"https://login.microsoftonline.com/{SP_TENANT}/oauth2/v2.0/token"
    data = {"grant_type": "client_credentials", "client_id": SP_CLIENT_ID,
            "client_secret": SP_SECRET, "scope": "https://graph.microsoft.com/.default"}
    try:
        r = requests.post(url, data=data, timeout=15)
        return r.json().get("access_token") if r.ok else None
    except Exception as e:
        print(f"SharePoint auth error: {e}")
        return None

def get_sharepoint_data():
    token = get_sp_token()
    if not token:
        return [
            {"label": "Documents", "value": "—", "note": "Configure SP secrets"},
            {"label": "Active sites", "value": "—", "note": "Configure SP secrets"},
            {"label": "Pending approvals", "value": "—", "note": "Configure SP secrets"},
            {"label": "Storage used", "value": "—", "note": "Configure SP secrets"}
        ]
    headers = {"Authorization": f"Bearer {token}"}
    stats = []
    try:
        drive_url = f"https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/drive"
        drive = requests.get(drive_url, headers=headers, timeout=10).json()
        quota = drive.get("quota", {})
        used = quota.get("used", 0)
        total = quota.get("total", 1)
        pct = round(used / total * 100) if total else 0
        stats.append({"label": "Storage used", "value": f"{pct}%", "note": f"Of {round(total/1e9)}GB quota"})
    except:
        stats.append({"label": "Storage used", "value": "—", "note": "Check SP config"})
    try:
        items_url = f"https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/drive/root/children?$filter=lastModifiedDateTime ge {week_ago}"
        items = requests.get(items_url, headers=headers, timeout=10).json().get("value", [])
        stats.insert(0, {"label": "Documents updated", "value": str(len(items)), "note": "This week"})
    except:
        stats.insert(0, {"label": "Documents updated", "value": "—", "note": "This week"})
    stats.append({"label": "Active sites", "value": "4", "note": "Ops, HR, Finance, IT"})
    stats.append({"label": "Pending approvals", "value": "—", "note": "Check SP manually"})
    return stats

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("Fetching JIRA data…")
    jira = get_jira_data()

    print("Fetching SharePoint data…")
    sp = get_sharepoint_data()

    report = {
        "generated": display,
        "sprint": jira["sprint"],
        "stats": jira["stats"],
        "projects": jira["projects"],
        "blockers": jira["blockers"],
        "actions": jira["actions"],
        "sharepoint": sp
    }

    with open("data/report.json", "w") as f:
        json.dump(report, f, indent=2)

    print(f"Report generated: {display}")
    print(f"  Projects: {len(report['projects'])}")
    print(f"  Blockers: {len(report['blockers'])}")

if __name__ == "__main__":
    main()
