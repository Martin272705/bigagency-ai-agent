import os,json,re,logging,schedule,time,requests
from datetime import datetime,timedelta
from anthropic import Anthropic

logging.basicConfig(level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
logger=logging.getLogger(__name__)
anthropic=Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
CLICKUP_API_KEY=os.environ.get("CLICKUP_API_KEY")
CLICKUP_TEAM_ID="90152385133"
CLICKUP_LIST_ID="901521595520"
MS_CLIENT_ID=os.environ.get("MS_CLIENT_ID")
MS_CLIENT_SECRET=os.environ.get("MS_CLIENT_SECRET")
MS_TENANT_ID=os.environ.get("MS_TENANT_ID")
MS_USER_EMAIL="cano@bigagency.sk"
MICHAL_ID="106588503"
PETER_ID="106591392"

def strip_html(t):
    if not t: return ""
    t=t.replace("&nbsp;"," ").replace("&amp;","&").replace("&lt;","<").replace("&gt;",">")
    t=re.sub(r'<style[^>]*>.*?</style>',' ',t,flags=re.DOTALL|re.IGNORECASE)
    t=re.sub(r'<script[^>]*>.*?</script>',' ',t,flags=re.DOTALL|re.IGNORECASE)
    t=re.sub(r'<br\s*/?>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'</p>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'</div>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'<[^>]+>','',t)
    return '\n'.join(l.strip() for l in t.split('\n') if l.strip())

def get_ms_token():
    return requests.post(f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type":"client_credentials","client_id":MS_CLIENT_ID,
              "client_secret":MS_CLIENT_SECRET,"scope":"https://graph.microsoft.com/.default"}).json().get("access_token")

def get_unread_emails_from_folder(folder_name="INFO Requesty",limit=15):
    token=get_ms_token()
    h={"Authorization":f"Bearer {token}"}
    folders=requests.get(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders",headers=h).json()
    fid=next((f["id"] for f in folders.get("value",[]) if f["displayName"].lower()==folder_name.lower()),None)
    if not fid: logger.error(f"Priecinok '{folder_name}' nenajdeny!"); return []
    url=(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders/{fid}/messages"
         f"?\$filter=isRead eq false&\$top={limit}&\$orderby=receivedDateTime desc"
         f"&\$select=id,subject,bodyPreview,body,from,sender,replyTo,receivedDateTime")
    emails=requests.get(url,headers=h).json().get("value",[])
    logger.info(f"Najdenych {len(emails)} neprecitanych emailov")
    return emails

def mark_email_as_read(eid):
    token=get_ms_token()
    requests.patch(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/messages/{eid}",
        headers={"Authorization":f"Bearer {token}","Content-Type":"application/json"},json={"isRead":True})

def get_client_contact(email,clean_body):
    from_addr=email.get("from",{}).get("emailAddress",{}).get("address","")
    from_name=email.get("from",{}).get("emailAddress",{}).get("name","")
    if from_addr and "noreply" not in from_addr.lower():
        return from_addr,from_name
    emails_in_body=re.findall(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',clean_body)
    for e in emails_in_body:
        if not any(x in e.lower() for x in ['noreply','no-reply','bigagency','microsoft','outlook']):
            return e,""
    for r in email.get("replyTo",[]):
        addr=r.get("emailAddress",{}).get("address","")
        name=r.get("emailAddress",{}).get("name","")
        if addr and "noreply" not in addr.lower(): return addr,name
    return "",""

def clickup_get(ep): return requests.get(f"https://api.clickup.com/api/v2/{ep}",headers={"Authorization":CLICKUP_API_KEY}).json()
def clickup_post(ep,d): return requests.post(f"https://api.clickup.com/api/v2/{ep}",headers={"Authorization":CLICKUP_API_KEY,"Content-Type":"application/json"},json=d).json()

def task_already_exists(task_name):
    """Zabrani duplikatom - skontroluje ci task s podobnym nazvom uz existuje."""
    try:
        existing=clickup_get(f"list/{CLICKUP_LIST_ID}/task?page=0")
        new_key=task_name.lower().strip()[:60]
        for t in existing.get("tasks",[]):
            existing_key=t.get("name","").lower().strip()[:60]
            if existing_key==new_key:
                logger.warning(f"Duplikat preskoceny: '{t.get('name')}'")
                return True
    except Exception as e:
        logger.error(f"Chyba kontroly duplikatov: {e}")
    return False

def get_team_workload():
    m=len(clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={MICHAL_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    p=len(clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={PETER_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    logger.info(f"Vytazenost: Michal={m}, Peter={p}")
    return {"michal":{"id":MICHAL_ID,"count":m},"peter":{"id":PETER_ID,"count":p}}

def get_less_busy_assignee(w):
    return (MICHAL_ID,"Michal Macai") if w["michal"]["count"]<=w["peter"]["count"] else (PETER_ID,"Peter Gerbel")

def create_task(name,desc,aid,priority="high",due_date=None):
    d={"name":name,"markdown_description":desc,"assignees":[int(aid)],"priority":2 if priority=="high" else 3,"status":"to do"}
    if due_date: d["due_date"]=int(due_date.timestamp()*1000)
    r=clickup_post(f"list/{CLICKUP_LIST_ID}/task",d)
    logger.info(f"Task vytvoreny: {name} (ID:{r.get('id')})")
    return r.get("id")

def add_comment_to_task(tid,c): clickup_post(f"task/{tid}/comment",{"comment_text":c})

def get_tasks_with_upcoming_deadlines(days=7):
    now=datetime.now(); db=now+timedelta(days=days)
    return clickup_get(f"team/{CLICKUP_TEAM_ID}/task?statuses[]=to do&statuses[]=in progress&due_date_gt={int(now.timestamp()*1000)}&due_date_lt={int(db.timestamp()*1000)}").get("tasks",[])

def analyze_email_with_claude(subject,body,sender_email,sender_name):
    p=f"""Si asistent eventovej agentury BigAgency. Analyzuj spravu.
Odosielatel: {sender_name} <{sender_email}>
Predmet: {subject}
Obsah: {body}
Realny dopyt = popis eventu po SK/CZ/EN.
IGNORUJ: rustinu/ukrajinskunu, newslettery, Profesia.sk, bankove vypisy, brigady.
JSON bez backticks: {{"is_real_request":true/false,"client_name":"meno","event_description":"popis","event_date":"datum alebo null","task_name":"nazov tasku"}}"""
    r=anthropic.messages.create(model="claude-sonnet-4-5-20250929",max_tokens=1024,messages=[{"role":"user","content":p}])
    raw=r.content[0].text.strip().replace("```json","").replace("```","").strip()
    return json.loads(raw)

def generate_task_description(analysis,body,sender_email,sender_name):
    today=datetime.now().strftime("%d.%m.%Y")
    aname=analysis.get("assignee_name","")
    cname=analysis.get("client_name",sender_name or "N/A")
    email_str=sender_email if sender_email else "NEPODARILO SA ZISTIT"
    return f"""## Kontakt na klienta
**Meno:** {cname}
**Email:** {email_str}

---

## Popis dopytu
{analysis.get('event_description','N/A')}

**Termin:** {analysis.get('event_date') or 'neuvedeny'}

---

## Kompletny text spravy
{body[:3000]}

---
Dopyt prijaty: {today}
@{aname} prosim spracuj tuto ponuku a odpovedz klientovi na: {email_str}"""

def process_info_emails():
    logger.info("Spracuvam INFO Requesty emaily...")
    try:
        emails=get_unread_emails_from_folder("INFO Requesty",limit=15)
        if not emails: logger.info("Ziadne nove emaily."); return
        workload=get_team_workload(); processed=0; skipped=0
        for email in emails:
            eid=email.get("id",""); subject=email.get("subject","")
            html_body=email.get("body",{}).get("content","")
            clean_body=strip_html(html_body) if html_body else email.get("bodyPreview","")
            sender_email,sender_name=get_client_contact(email,clean_body)
            try: analysis=analyze_email_with_claude(subject,clean_body[:3000],sender_email,sender_name)
            except Exception as e: logger.error(f"Analyza: {e}"); mark_email_as_read(eid); continue
            if not analysis.get("is_real_request"):
                logger.info(f"Ignorovany: {subject[:50]}"); mark_email_as_read(eid); continue
            task_name=analysis.get("task_name",subject[:100])
            if task_already_exists(task_name):
                skipped+=1; mark_email_as_read(eid); continue
            aid,aname=get_less_busy_assignee(workload)
            analysis["assignee_name"]=aname
            if aid==MICHAL_ID: workload["michal"]["count"]+=1
            else: workload["peter"]["count"]+=1
            desc=generate_task_description(analysis,clean_body[:3000],sender_email,sender_name)
            tid=create_task(task_name,desc,aid,"high")
            if tid: processed+=1; logger.info(f"OK: {sender_email} -> {aname}"); mark_email_as_read(eid)
        logger.info(f"Spracovanych {processed}, preskoceno duplikatov: {skipped}")
    except Exception as e: logger.error(f"Chyba: {e}")

def check_deadlines():
    logger.info("Deadliny...")
    try:
        for t in get_tasks_with_upcoming_deadlines(7):
            dms=t.get("due_date")
            if not dms: continue
            dd=datetime.fromtimestamp(int(dms)/1000); dl=(dd-datetime.now()).days
            u="URGENTNE" if dl<=1 else ("BLIZI SA" if dl<=3 else "Pripomienka")
            add_comment_to_task(t["id"],f"{u}\nDeadline: {dd.strftime('%d.%m.%Y')} ({dl} dni)")
    except Exception as e: logger.error(f"Deadline: {e}")

def weekly_report():
    logger.info("Tyzdenny report...")
    try:
        all_t=clickup_get(f"team/{CLICKUP_TEAM_ID}/task?statuses[]=to do&statuses[]=in progress").get("tasks",[])
        now=datetime.now(); we=now+timedelta(days=7)
        urg=[t for t in all_t if t.get("due_date") and datetime.fromtimestamp(int(t["due_date"])/1000)<we]
        mc=len([t for t in all_t if any(a["id"]==int(MICHAL_ID) for a in t.get("assignees",[]))])
        pc=len([t for t in all_t if any(a["id"]==int(PETER_ID) for a in t.get("assignees",[]))])
        logger.info(f"Report {now.strftime('%d.%m.%Y')}: M={mc} P={pc} Urgent={len(urg)}")
    except Exception as e: logger.error(f"Report: {e}")

def setup_schedule():
    schedule.every().day.at("08:30").do(check_deadlines)
    schedule.every().day.at("09:00").do(process_info_emails)
    schedule.every().day.at("16:00").do(process_info_emails)
    schedule.every().monday.at("08:00").do(weekly_report)
    logger.info("Scheduler: 08:30 deadliny | 09:00+16:00 emaily | pon 08:00 report")

def main():
    logger.info("BigAgency AI Agent spusteny!")
    missing=[v for v in ["ANTHROPIC_API_KEY","CLICKUP_API_KEY","MS_CLIENT_ID","MS_CLIENT_SECRET","MS_TENANT_ID"] if not os.environ.get(v)]
    if missing: logger.error(f"Chybaju: {missing}"); return
    setup_schedule()
    logger.info("Cakam na ulohy...")
    while True: schedule.run_pending(); time.sleep(60)

if __name__=="__main__": main()
