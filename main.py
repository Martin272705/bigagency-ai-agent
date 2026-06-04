"""
BigAgency AI Agent
Automaticky spracuva emaily, vytvara tasky v ClickUp a sleduje deadliny.
"""

import os
import json
import re
import logging
import schedule
import time
import requests
from datetime import datetime, timedelta
from anthropic import Anthropic

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

anthropic = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

CLICKUP_API_KEY = os.environ.get("CLICKUP_API_KEY")
CLICKUP_TEAM_ID = "90152385133"
CLICKUP_LIST_ID = "901521595520"
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET")
MS_TENANT_ID = os.environ.get("MS_TENANT_ID")
MS_USER_EMAIL = "cano@bigagency.sk"

MICHAL_ID = "106588503"
PETER_ID = "106591392"

def strip_html(html_text):
    """Odstrani HTML tagy a vrati cisty text."""
    if not html_text:
        return ""
    text = html_text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
    text = re.sub(r'<style[^>]*>.*?</style>', ' ', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<script[^>]*>.*?</script>', ' ', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</div>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    return '\n'.join(lines)

def get_ms_token():
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {"grant_type": "client_credentials", "client_id": MS_CLIENT_ID,
            "client_secret": MS_CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"}
    return requests.post(url, data=data).json().get("access_token")

def get_unread_emails_from_folder(folder_name="INFO Requesty", limit=15):
    """Nacita IBA NEPRECITANE emaily - deduplication."""
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}"}
    folders_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders",
        headers=headers).json()
    folder_id = next((f["id"] for f in folders_resp.get("value", [])
                      if f["displayName"].lower() == folder_name.lower()), None)
    if not folder_id:
        logger.error(f"Priecinok '{folder_name}' nenajdeny!")
        return []
    # Nacita full body + from/sender polia + isRead filter
    url = (f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}"
           f"/mailFolders/{folder_id}/messages"
           f"?\$filter=isRead eq false&\$top={limit}"
           f"&\$orderby=receivedDateTime desc"
           f"&\$select=id,subject,bodyPreview,body,from,sender,replyTo,receivedDateTime")
    emails = requests.get(url, headers=headers).json().get("value", [])
    logger.info(f"Najdenych {len(emails)} neprecitanych emailov v '{folder_name}'")
    return emails

def mark_email_as_read(email_id):
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    requests.patch(
        f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/messages/{email_id}",
        headers=headers, json={"isRead": True})

def get_sender_email(email):
    """Ziska skutocnu emailovu adresu odosielatela - vzdy z From pola, nie z tela."""
    # Skus from -> emailAddress -> address
    from_field = email.get("from", {})
    if from_field:
        addr = from_field.get("emailAddress", {}).get("address", "")
        name = from_field.get("emailAddress", {}).get("name", "")
        if addr and "noreply" not in addr.lower():
            return addr, name
    # Skus replyTo
    reply_to = email.get("replyTo", [])
    if reply_to:
        addr = reply_to[0].get("emailAddress", {}).get("address", "")
        name = reply_to[0].get("emailAddress", {}).get("name", "")
        if addr and "noreply" not in addr.lower():
            return addr, name
    # Skus sender
    sender = email.get("sender", {})
    if sender:
        addr = sender.get("emailAddress", {}).get("address", "")
        name = sender.get("emailAddress", {}).get("name", "")
        if addr and "noreply" not in addr.lower():
            return addr, name
    return "", ""

def clickup_get(endpoint):
    return requests.get(f"https://api.clickup.com/api/v2/{endpoint}",
                        headers={"Authorization": CLICKUP_API_KEY}).json()

def clickup_post(endpoint, data):
    return requests.post(f"https://api.clickup.com/api/v2/{endpoint}",
                         headers={"Authorization": CLICKUP_API_KEY, "Content-Type": "application/json"},
                         json=data).json()

def get_team_workload():
    michal = len(clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={MICHAL_ID}&statuses[]=to do&statuses[]=in progress").get("tasks", []))
    peter = len(clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={PETER_ID}&statuses[]=to do&statuses[]=in progress").get("tasks", []))
    logger.info(f"Vytazenost: Michal={michal}, Peter={peter}")
    return {"michal": {"id": MICHAL_ID, "count": michal}, "peter": {"id": PETER_ID, "count": peter}}

def get_less_busy_assignee(workload):
    if workload["michal"]["count"] <= workload["peter"]["count"]:
        return MICHAL_ID, "Michal Macai"
    return PETER_ID, "Peter Gerbel"

def create_task(name, description, assignee_id, priority="high", due_date=None):
    data = {"name": name, "markdown_description": description,
            "assignees": [int(assignee_id)], "priority": 2 if priority == "high" else 3, "status": "to do"}
    if due_date:
        data["due_date"] = int(due_date.timestamp() * 1000)
    result = clickup_post(f"list/{CLICKUP_LIST_ID}/task", data)
    logger.info(f"Task vytvoreny: {name} (ID: {result.get('id')})")
    return result.get("id")

def add_comment_to_task(task_id, comment):
    clickup_post(f"task/{task_id}/comment", {"comment_text": comment})

def get_tasks_with_upcoming_deadlines(days=7):
    now = datetime.now()
    due_before = now + timedelta(days=days)
    result = clickup_get(
        f"team/{CLICKUP_TEAM_ID}/task?statuses[]=to do&statuses[]=in progress"
        f"&due_date_gt={int(now.timestamp()*1000)}&due_date_lt={int(due_before.timestamp()*1000)}")
    return result.get("tasks", [])

def analyze_email_with_claude(email_subject, email_body_clean, sender_email, sender_name):
    """Analyzuje email - sender_email uz je vytiahnuty z From pola."""
    prompt = f"""Si asistent eventovej agentury BigAgency. Analyzuj tuto spravu a urc ci je to realny dopyt od klienta.

Odosielatel: {sender_name} <{sender_email}>
Predmet: {email_subject}
Obsah spravy: {email_body_clean}

Realny dopyt je ak obsahuje popis eventu/podujatia a je po slovensky/cesky/anglicky.
IGNORUJ: rusticnu/ukrajinskunu spravu, newslettery, Profesia.sk, bankove vypisy, brigady.

Odpoved MUSI byt iba ciste JSON bez backticks:
{{"is_real_request": true/false, "client_name": "meno klienta", "event_description": "kratky popis co chcu", "event_date": "datum ak je zname alebo null", "task_name": "nazov tasku pre ClickUp"}}"""

    response = anthropic.messages.create(
        model="claude-sonnet-4-5-20250929", max_tokens=1024,
        messages=[{"role": "user", "content": prompt}])
    raw = response.content[0].text.strip().replace("```json","").replace("```","").strip()
    return json.loads(raw)

def generate_task_description(analysis, email_body_clean, sender_email, sender_name):
    today = datetime.now().strftime("%d.%m.%Y")
    assignee_name = analysis.get("assignee_name", "")
    client_name = analysis.get("client_name", sender_name or "N/A")

    return f"""## Kontakt na klienta
👤 **Meno:** {client_name}
✉️ **Email:** {sender_email}

---

## Popis dopytu
{analysis.get('event_description','N/A')}

📅 **Termin:** {analysis.get('event_date') or 'neuvedeny'}

---

## Kompletny text spravy
{email_body_clean[:3000]}

---
📧 Dopyt prijaty: {today}
@{assignee_name} prosim spracuj tuto ponuku a odpovedz klientovi na: {sender_email}"""

def process_info_emails():
    """Spracuje NEPRECITANE emaily z INFO Requesty."""
    logger.info("Spracuvam INFO Requesty emaily...")
    try:
        emails = get_unread_emails_from_folder("INFO Requesty", limit=15)
        if not emails:
            logger.info("Ziadne nove neprecitane emaily.")
            return
        workload = get_team_workload()
        processed = 0
        for email in emails:
            email_id = email.get("id", "")
            subject = email.get("subject", "")

            # Ziskaj skutocny email odosielatela z From pola
            sender_email, sender_name = get_sender_email(email)

            # Strip HTML z body
            html_body = email.get("body", {}).get("content", "")
            clean_body = strip_html(html_body) if html_body else email.get("bodyPreview", "")

            try:
                analysis = analyze_email_with_claude(subject, clean_body[:3000], sender_email, sender_name)
            except Exception as e:
                logger.error(f"Chyba analyzy: {e}")
                mark_email_as_read(email_id)
                continue

            if not analysis.get("is_real_request"):
                logger.info(f"Ignorovany: {subject[:50]}")
                mark_email_as_read(email_id)
                continue

            assignee_id, assignee_name = get_less_busy_assignee(workload)
            analysis["assignee_name"] = assignee_name
            if assignee_id == MICHAL_ID:
                workload["michal"]["count"] += 1
            else:
                workload["peter"]["count"] += 1

            description = generate_task_description(analysis, clean_body[:3000], sender_email, sender_name)
            task_id = create_task(
                name=analysis.get("task_name", subject[:100]),
                description=description, assignee_id=assignee_id, priority="high")

            if task_id:
                processed += 1
                logger.info(f"Task: {sender_email} -> {assignee_name}")
                mark_email_as_read(email_id)

        logger.info(f"Spracovanych {processed} novych requestov")
    except Exception as e:
        logger.error(f"Chyba: {e}")

def check_deadlines():
    logger.info("Kontrolujem deadliny...")
    try:
        for task in get_tasks_with_upcoming_deadlines(days=7):
            task_id = task.get("id")
            due_date_ms = task.get("due_date")
            if not due_date_ms:
                continue
            due_date = datetime.fromtimestamp(int(due_date_ms) / 1000)
            days_left = (due_date - datetime.now()).days
            urgency = "URGENTNE" if days_left <= 1 else ("BLIZI SA DEADLINE" if days_left <= 3 else "Pripomienka")
            add_comment_to_task(task_id, f"{urgency}\nDeadline: {due_date.strftime('%d.%m.%Y')} (zostava {days_left} dni)")
    except Exception as e:
        logger.error(f"Chyba deadlinov: {e}")

def weekly_report():
    logger.info("Generujem tyzdenný report...")
    try:
        all_tasks = clickup_get(f"team/{CLICKUP_TEAM_ID}/task?statuses[]=to do&statuses[]=in progress").get("tasks",[])
        now = datetime.now()
        week_end = now + timedelta(days=7)
        urgent = [t for t in all_tasks if t.get("due_date") and datetime.fromtimestamp(int(t["due_date"])/1000) < week_end]
        michal_cnt = len([t for t in all_tasks if any(a["id"]==int(MICHAL_ID) for a in t.get("assignees",[]))])
        peter_cnt = len([t for t in all_tasks if any(a["id"]==int(PETER_ID) for a in t.get("assignees",[]))])
        logger.info(f"Report {now.strftime('%d.%m.%Y')}: Michal={michal_cnt} Peter={peter_cnt} Urgent={len(urgent)}")
    except Exception as e:
        logger.error(f"Chyba reportu: {e}")

def setup_schedule():
    schedule.every().day.at("08:30").do(check_deadlines)
    schedule.every().day.at("09:00").do(process_info_emails)
    schedule.every().day.at("16:00").do(process_info_emails)
    schedule.every().monday.at("08:00").do(weekly_report)
    logger.info("Scheduler: 08:30 deadliny | 09:00+16:00 emaily | pon 08:00 report")

def main():
    logger.info("BigAgency AI Agent spusteny!")
    missing = [v for v in ["ANTHROPIC_API_KEY","CLICKUP_API_KEY","MS_CLIENT_ID","MS_CLIENT_SECRET","MS_TENANT_ID"]
               if not os.environ.get(v)]
    if missing:
        logger.error(f"Chybaju env vars: {missing}")
        return
    setup_schedule()
    logger.info("Cakam na naplanovane ulohy...")
    while True:
        schedule.run_pending()
        time.sleep(60)

if __name__ == "__main__":
    main()
