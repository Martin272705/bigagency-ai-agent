"""
BigAgency AI Agent
Automaticky spracúva emaily, vytvára tasky v ClickUp a sleduje deadliny.
"""

import os
import json
import logging
import schedule
import time
import requests
from datetime import datetime, timedelta
from anthropic import Anthropic

# Logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# API klienti
anthropic = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

# Konštanty
CLICKUP_API_KEY = os.environ.get("CLICKUP_API_KEY")
CLICKUP_TEAM_ID = "90152385133"
CLICKUP_LIST_ID = "901521595520"  # Cenové ponuky
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET")
MS_TENANT_ID = os.environ.get("MS_TENANT_ID")
MS_USER_EMAIL = "cano@bigagency.sk"

# Kolegovia
MICHAL_ID = "106588503"
PETER_ID = "106591392"

MICHAL_DM_CHANNEL = "2kyr0ekd-735"  # DM kanál (treba overiť)
PETER_DM_CHANNEL = "2kyr0ekd-55"


# ============================================================
# MICROSOFT 365 - EMAIL
# ============================================================

def get_ms_token():
    """Získa Microsoft access token."""
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data)
    return response.json().get("access_token")


def get_emails_from_folder(folder_name="INFO Requesty", limit=15):
    """Načíta emaily z daného priečinka."""
    token = get_ms_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Najprv nájdeme folder ID
    folders_url = f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders"
    folders_resp = requests.get(folders_url, headers=headers).json()

    folder_id = None
    for folder in folders_resp.get("value", []):
        if folder["displayName"].lower() == folder_name.lower():
            folder_id = folder["id"]
            break

    if not folder_id:
        logger.error(f"Priečinok '{folder_name}' nenájdený!")
        return []

    # Načítame emaily
    emails_url = f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders/{folder_id}/messages?$top={limit}&$orderby=receivedDateTime desc"
    emails_resp = requests.get(emails_url, headers=headers).json()
    return emails_resp.get("value", [])


# ============================================================
# CLICKUP
# ============================================================

def clickup_get(endpoint):
    """GET request na ClickUp API."""
    headers = {"Authorization": CLICKUP_API_KEY}
    url = f"https://api.clickup.com/api/v2/{endpoint}"
    response = requests.get(url, headers=headers)
    return response.json()


def clickup_post(endpoint, data):
    """POST request na ClickUp API."""
    headers = {
        "Authorization": CLICKUP_API_KEY,
        "Content-Type": "application/json"
    }
    url = f"https://api.clickup.com/api/v2/{endpoint}"
    response = requests.post(url, headers=headers, json=data)
    return response.json()


def get_team_workload():
    """Pozrie vyťaženosť Michala a Petra."""
    michal_tasks = clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={MICHAL_ID}&statuses[]=to do&statuses[]=in progress")
    peter_tasks = clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={PETER_ID}&statuses[]=to do&statuses[]=in progress")

    michal_count = len(michal_tasks.get("tasks", []))
    peter_count = len(peter_tasks.get("tasks", []))

    logger.info(f"Vyťaženosť: Michal={michal_count}, Peter={peter_count}")
    return {
        "michal": {"id": MICHAL_ID, "count": michal_count},
        "peter": {"id": PETER_ID, "count": peter_count}
    }


def get_less_busy_assignee(workload):
    """Vráti ID menej vyťaženého kolegu."""
    if workload["michal"]["count"] <= workload["peter"]["count"]:
        return MICHAL_ID, "Michal Mačai"
    else:
        return PETER_ID, "Peter Gerbel"


def create_task(name, description, assignee_id, priority="high", due_date=None):
    """Vytvorí task v ClickUpe."""
    data = {
        "name": name,
        "markdown_description": description,
        "assignees": [int(assignee_id)],
        "priority": 2 if priority == "high" else 3,
        "status": "to do"
    }
    if due_date:
        data["due_date"] = int(due_date.timestamp() * 1000)

    result = clickup_post(f"list/{CLICKUP_LIST_ID}/task", data)
    task_id = result.get("id")
    logger.info(f"Task vytvorený: {name} (ID: {task_id})")
    return task_id


def add_comment_to_task(task_id, comment):
    """Pridá komentár k tasku."""
    clickup_post(f"task/{task_id}/comment", {"comment_text": comment})
    logger.info(f"Komentár pridaný k tasku {task_id}")


def send_dm_clickup(channel_id, message):
    """Pošle DM správu v ClickUpe."""
    clickup_post(f"chat/channel/{channel_id}/message", {"content": message})
    logger.info(f"DM odoslaný do kanála {channel_id}")


def get_tasks_with_upcoming_deadlines(days=7):
    """Vráti tasky s deadlinom v najbližších X dňoch."""
    now = datetime.now()
    due_before = now + timedelta(days=days)

    due_before_ms = int(due_before.timestamp() * 1000)
    now_ms = int(now.timestamp() * 1000)

    result = clickup_get(
        f"team/{CLICKUP_TEAM_ID}/task?"
        f"statuses[]=to do&statuses[]=in progress"
        f"&due_date_gt={now_ms}&due_date_lt={due_before_ms}"
    )
    return result.get("tasks", [])


# ============================================================
# CLAUDE - AI ROZHODOVANIE
# ============================================================

def analyze_email_with_claude(email_subject, email_body):
    """Použije Claude na analýzu emailu."""
    prompt = f"""Si asistent eventovej agentury BigAgency. Analyzuj tuto spravu a urc ci je to realny dopyt od klienta.

DOLEZITE: Vela sprav pride cez webovy formular z adresy noreply@bigagency.sk - to je normalne! V takom pripade citaj obsah spravy kde je skutocny text od klienta.

Predmet: {email_subject}
Obsah spravy: {email_body}

Realny dopyt od klienta je ak obsah obsahuje:
- Zmysluplny popis co klient chce (event, prenajom, organizacia podujatia...)
- Realne meno osoby alebo nazov firmy
- Je napisany po slovensky, cesky alebo anglicky

IGNORUJ ak:
- Obsah je po rusky alebo ukrainsky (spam)
- Je to newsletter, reklama, cold outreach
- Je to notifikacia z Profesia.sk, bankovy vypis, systemova sprava
- Brigady, pracovne ponuky, HR veci

DOLEZITE: Odpoved MUSI byt iba ciste JSON bez backticks:
{{"is_real_request": true/false, "client_name": "meno klienta alebo nazov firmy", "client_email": "email klienta", "event_description": "kratky popis co chcu", "event_date": "datum ak je zname alebo null", "task_name": "nazov tasku pre ClickUp"}}"""

    response = anthropic.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    # Odstráni backticks ak sú
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


def generate_task_description(analysis, email_body):
    """Vygeneruje podrobný popis tasku."""
    today = datetime.now().strftime("%d.%m.%Y")
    assignee_name = analysis.get("assignee_name", "")

    desc = f"""📧 **Dopyt prijatý:** {today}
👤 **Klient:** {analysis.get('client_name', 'N/A')} ({analysis.get('client_email', 'N/A')})

📋 **Popis dopytu:**
{analysis.get('event_description', 'N/A')}

📅 **Termín:** {analysis.get('event_date') or 'neuvedený'}

📝 **Celý text správy:**
{email_body[:1500]}

---
@{assignee_name} prosím spracuj túto ponuku a odpovedz klientovi."""

    return desc


# ============================================================
# HLAVNÉ ÚLOHY
# ============================================================

def process_info_emails():
    """Hlavná úloha: spracuje emaily z INFO Requesty."""
    logger.info("🔍 Spracúvam INFO Requesty emaily...")

    try:
        emails = get_emails_from_folder("INFO Requesty", limit=15)
        workload = get_team_workload()

        processed = 0
        for email in emails:
            subject = email.get("subject", "")
            body = email.get("bodyPreview", "") + "\n" + email.get("body", {}).get("content", "")

            # Analýza emailu
            try:
                analysis = analyze_email_with_claude(subject, body[:3000])
            except Exception as e:
                logger.error(f"Chyba pri analýze emailu: {e}")
                continue

            if not analysis.get("is_real_request"):
                logger.info(f"Email ignorovaný (nie je request): {subject[:50]}")
                continue

            # Priraď kolegu
            assignee_id, assignee_name = get_less_busy_assignee(workload)
            analysis["assignee_name"] = assignee_name

            # Aktualizuj workload
            if assignee_id == MICHAL_ID:
                workload["michal"]["count"] += 1
            else:
                workload["peter"]["count"] += 1

            # Vytvor task
            description = generate_task_description(analysis, body[:1500])
            task_id = create_task(
                name=analysis.get("task_name", subject[:100]),
                description=description,
                assignee_id=assignee_id,
                priority="high"
            )

            if task_id:
                processed += 1
                logger.info(f"✅ Task vytvorený pre: {analysis.get('client_name')} → {assignee_name}")

        logger.info(f"📊 Spracovaných {processed} nových requestov")

    except Exception as e:
        logger.error(f"Chyba pri spracovaní emailov: {e}")


def check_deadlines():
    """Skontroluje deadliny a napíše komentáre do taskov."""
    logger.info("📅 Kontrolujem deadliny...")

    try:
        tasks = get_tasks_with_upcoming_deadlines(days=7)

        for task in tasks:
            task_id = task.get("id")
            task_name = task.get("name")
            due_date_ms = task.get("due_date")

            if not due_date_ms:
                continue

            due_date = datetime.fromtimestamp(int(due_date_ms) / 1000)
            days_left = (due_date - datetime.now()).days

            assignees = task.get("assignees", [])
            assignee_names = [a.get("username", "kolega") for a in assignees]

            if days_left <= 1:
                urgency = "🚨 URGENTNÉ"
            elif days_left <= 3:
                urgency = "⚠️ BLÍŽI SA DEADLINE"
            else:
                urgency = "📅 Pripomienka"

            comment = (
                f"{urgency}\n\n"
                f"Deadline pre tento task je **{due_date.strftime('%d.%m.%Y')}** "
                f"(zostáva {days_left} {'deň' if days_left == 1 else 'dni' if days_left < 5 else 'dní'}).\n\n"
                f"Ako to vyzerá s prípravou? Je potrebná nejaká pomoc?"
            )

            add_comment_to_task(task_id, comment)
            logger.info(f"⏰ Komentár pridaný k tasku: {task_name} (zostáva {days_left} dní)")

    except Exception as e:
        logger.error(f"Chyba pri kontrole deadlinov: {e}")


def weekly_report():
    """Pondelkový týždenný report."""
    logger.info("📊 Generujem týždenný report...")

    try:
        # Všetky otvorené tasky
        all_tasks = clickup_get(
            f"team/{CLICKUP_TEAM_ID}/task?statuses[]=to do&statuses[]=in progress"
        ).get("tasks", [])

        michal_tasks = [t for t in all_tasks if any(a["id"] == int(MICHAL_ID) for a in t.get("assignees", []))]
        peter_tasks = [t for t in all_tasks if any(a["id"] == int(PETER_ID) for a in t.get("assignees", []))]
        unassigned = [t for t in all_tasks if not t.get("assignees")]

        # Tasky s deadlinom tento týždeň
        now = datetime.now()
        week_end = now + timedelta(days=7)
        urgent_tasks = [
            t for t in all_tasks
            if t.get("due_date") and datetime.fromtimestamp(int(t["due_date"]) / 1000) < week_end
        ]

        report = f"""📊 **TÝŽDENNÝ REPORT – {now.strftime('%d.%m.%Y')}**

**Celkový prehľad:**
- Michal Mačai: {len(michal_tasks)} otvorených taskov
- Peter Gerbel: {len(peter_tasks)} otvorených taskov
- Bez assignee: {len(unassigned)} taskov
- Celkom: {len(all_tasks)} otvorených taskov

**⚠️ Deadliny tento týždeň ({len(urgent_tasks)}):**
"""
        for task in urgent_tasks[:10]:
            due = datetime.fromtimestamp(int(task["due_date"]) / 1000)
            assignees = ", ".join([a.get("username", "?") for a in task.get("assignees", [])])
            report += f"- {task['name']} → {due.strftime('%d.%m')} ({assignees})\n"

        if unassigned:
            report += f"\n**❗ Tasky bez priradeného (treba riešiť):**\n"
            for task in unassigned[:5]:
                report += f"- {task['name']}\n"

        logger.info(report)
        # Tu môžeme pridať posielanie reportu emailom alebo DM

    except Exception as e:
        logger.error(f"Chyba pri generovaní reportu: {e}")


# ============================================================
# SCHEDULER
# ============================================================

def setup_schedule():
    """Nastaví rozvrh úloh."""

    # Každý deň o 8:30 - kontrola deadlinov (pred poradou)
    schedule.every().day.at("08:30").do(check_deadlines)

    # Každý deň o 9:00 - spracovanie emailov
    schedule.every().day.at("09:00").do(process_info_emails)

    # Každý deň o 16:00 - spracovanie emailov znova
    schedule.every().day.at("16:00").do(process_info_emails)

    # Každý pondelok o 8:00 - týždenný report
    schedule.every().monday.at("08:00").do(weekly_report)

    logger.info("✅ Scheduler nastavený:")
    logger.info("  - 08:00 pondelok: Týždenný report")
    logger.info("  - 08:30 denne: Kontrola deadlinov")
    logger.info("  - 09:00 denne: Spracovanie INFO emailov")
    logger.info("  - 16:00 denne: Spracovanie INFO emailov")


def main():
    """Hlavná funkcia."""
    logger.info("🚀 BigAgency AI Agent spustený!")

    # Skontroluj premenné prostredia
    required_vars = ["ANTHROPIC_API_KEY", "CLICKUP_API_KEY", "MS_CLIENT_ID", "MS_CLIENT_SECRET", "MS_TENANT_ID"]
    missing = [var for var in required_vars if not os.environ.get(var)]
    if missing:
        logger.error(f"Chýbajú environment variables: {missing}")
        return

    setup_schedule()

    # Spusti hneď pri štarte (voliteľné - zakomentuj ak nechceš)
    # process_info_emails()

    # Hlavná slučka
    logger.info("⏰ Čakám na naplánované úlohy...")
    while True:
        schedule.run_pending()
        time.sleep(60)  # Kontroluje každú minútu


if __name__ == "__main__":
    main()
