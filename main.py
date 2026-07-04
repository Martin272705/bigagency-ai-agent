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
MARTIN_ID="284426112"
STANISLAV_ID="106588288"

def strip_html(t):
    if not t: return ""
    t=t.replace("&nbsp;"," ").replace("&amp;","&").replace("&lt;","<").replace("&gt;",">")
    t=re.sub(r'<style[^>]*>.*?</style>',' ',t,flags=re.DOTALL|re.IGNORECASE)
    t=re.sub(r'<script[^>]*>.*?</script>',' ',t,flags=re.DOTALL|re.IGNORECASE)
    t=re.sub(r'<br\s*/?>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'</p>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'</div>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'<li[^>]*>','\n- ',t,flags=re.IGNORECASE)
    t=re.sub(r'</li>','',t,flags=re.IGNORECASE)
    t=re.sub(r'<(?:ul|ol)[^>]*>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'</(?:ul|ol)>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'</tr>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'<[^>]+>','',t)
    return '\n'.join(l.strip() for l in t.split('\n') if l.strip())

def get_ms_token():
    return requests.post(f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type":"client_credentials","client_id":MS_CLIENT_ID,
        "client_secret":MS_CLIENT_SECRET,"scope":"https://graph.microsoft.com/.default"}).json().get("access_token")

def get_emails_from_folder(folder_name="INFO Requesty",limit=50,days_back=30):
    token=get_ms_token()
    h={"Authorization":f"Bearer {token}"}
    folders=requests.get(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders",headers=h).json()
    fid=next((f["id"] for f in folders.get("value",[]) if f["displayName"].lower()==folder_name.lower()),None)
    if not fid: logger.error(f"Priecinok '{folder_name}' nenajdeny!"); return []
    since=(datetime.now()-timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url=(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders/{fid}/messages"
        f"?\$filter=receivedDateTime ge {since}&\$top={limit}&\$orderby=receivedDateTime desc"
        f"&\$select=id,subject,bodyPreview,body,from,sender,replyTo,receivedDateTime,internetMessageId")
    emails=requests.get(url,headers=h).json().get("value",[])
    logger.info(f"Najdenych {len(emails)} emailov za poslednych {days_back} dni")
    return emails

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

def normalize(s):
    """Normalizuje nazov na porovnanie - male pismena, bez diakritiky a prebytocnych medzier."""
    s=s.lower().strip()
    for a,b in [('ÃÂ¡','a'),('ÃÂ','c'),('ÃÂ','d'),('ÃÂ©','e'),('ÃÂ­','i'),('ÃÂ¾','l'),('ÃÂº','l'),
                ('ÃÂ','n'),('ÃÂ³','o'),('ÃÂ´','o'),('ÃÂ','r'),('ÃÂ¡','s'),('ÃÂ¥','t'),('ÃÂº','u'),('ÃÂ½','y'),('ÃÂ¾','z')]:
        s=s.replace(a,b)
    return re.sub(r'\s+',' ',s)

def task_already_exists(msg_id, task_name):
    """Dedup: kontroluje MSG_ID v popise aj nazov tasku (pre starstie tasky bez MSG_ID)."""
    norm_name=normalize(task_name)[:80] if task_name else ""
    try:
        page=0
        while True:
            tasks=clickup_get(f"list/{CLICKUP_LIST_ID}/task?page={page}&include_closed=true").get("tasks",[])
            if not tasks: break
            for t in tasks:
                desc=(t.get("description") or "")
                # Primarna kontrola: MSG_ID (exactne, unikatne pre kazdy email)
                if msg_id and f"[MSG_ID: {msg_id}]" in desc:
                    logger.warning(f"Duplikat MSG_ID - preskakujem: {msg_id[:50]}")
                    return True
                # Sekundarna kontrola: nazov tasku (pre tasky vytvorene pred zavedenim MSG_ID)
                if norm_name and normalize(t.get("name",""))[:80]==norm_name:
                    logger.warning(f"Duplikat nazov - preskakujem: {task_name[:60]}")
                    return True
            if len(tasks)<100: break
            page+=1
    except Exception as e:
        logger.error(f"Chyba dedup: {e}")
    return False

def get_team_workload():
    """Pocita len tasky v zozname Cenove ponuky (To do + In progress)."""
    base=f"list/{CLICKUP_LIST_ID}/task?include_closed=false"
    all_tasks=clickup_get(base).get("tasks",[])
    active=[t for t in all_tasks if t.get("status",{}).get("status","").lower() in ("to do","in progress")]
    def count(uid): return len([t for t in active if any(str(a.get("id"))==str(uid) for a in t.get("assignees",[]))])
    m=count(MICHAL_ID); ma=count(MARTIN_ID); s=count(STANISLAV_ID)
    logger.info(f"Vytazenost (Cenove ponuky): Michal={m}, Martin={ma}, Stanislav={s}")
    return {"michal":{"id":MICHAL_ID,"count":m},"martin":{"id":MARTIN_ID,"count":ma},"stanislav":{"id":STANISLAV_ID,"count":s}}

def get_less_busy_assignee(w):
    people=[("michal",MICHAL_ID,"Michal Macai"),("martin",MARTIN_ID,"Martin Cano"),("stanislav",STANISLAV_ID,"Stanislav Kois")]
    least=min(people,key=lambda x:w[x[0]]["count"])
    return least[1],least[2]

def create_task(name,desc,aid,priority="high",due_date=None):
    d={"name":name,"markdown_description":desc,"assignees":[int(aid)],"priority":2 if priority=="high" else 3,"status":"to do"}
    if due_date: d["due_date"]=int(due_date.timestamp()*1000)
    r=clickup_post(f"list/{CLICKUP_LIST_ID}/task",d)
    logger.info(f"Task vytvoreny: {name} (ID:{r.get('id')})")
    return r.get("id")

def add_comment_to_task(tid,c): clickup_post(f"task/{tid}/comment",{"comment_text":c})


def has_stale_reminder(tid):
    try:
        comments=clickup_get(f"task/{tid}/comment").get("comments",[])
        return any("[AUTO-REMINDER]" in (cm.get("comment_text") or "") for cm in comments)
    except Exception as e:
        logger.error(f"Chyba has_stale_reminder: {e}")
        return False

def check_stale_tasks():
    logger.info("Kontrolujem stale tasky (48h)...")
    try:
        cutoff_ms=int((datetime.now()-timedelta(hours=48)).timestamp()*1000)
        tasks=clickup_get(f"list/{CLICKUP_LIST_ID}/task?statuses[]=to do&date_created_lt={cutoff_ms}").get("tasks",[])
        reminded=0
        for task in tasks:
            tid=task.get("id"); tname=task.get("name","")
            if has_stale_reminder(tid):
                logger.info(f"Reminder uz existuje: {tname}"); continue
            assignees=task.get("assignees",[])
            names=", ".join(a.get("username","kolega") for a in assignees) if assignees else "tim"
            created_ms=task.get("date_created")
            hours_old=int((datetime.now()-datetime.fromtimestamp(int(created_ms)/1000)).total_seconds()/3600) if created_ms else 48
            age_str=f"{hours_old//24} dni" if hours_old>=48 else f"{hours_old} hodin"
            comment=f"[AUTO-REMINDER]\n\nAhoj {names},\n\ntento dopyt caka na spracovanie uz {age_str}. Ako to ide? Potrebujes s niecim pomoc?\n\nAk si ho uz zobral/a do riesenia, nezabudni prepnut status na In Progress."
            add_comment_to_task(tid,comment)
            reminded+=1; logger.info(f"Reminder: {tname} -> {names}")
        logger.info(f"Stale remindery: {reminded}")
    except Exception as e:
        logger.error(f"Chyba check_stale_tasks: {e}")

def get_tasks_with_upcoming_deadlines(days=7):
    now=datetime.now(); db=now+timedelta(days=days)
    return clickup_get(f"team/{CLICKUP_TEAM_ID}/task?statuses[]=to do&statuses[]=in progress&due_date_gt={int(now.timestamp()*1000)}&due_date_lt={int(db.timestamp()*1000)}").get("tasks",[])

def extract_first_json(text):
    """Extrahuje prvy kompletny JSON objekt z textu - robustne aj ked Claude prida extra text."""
    start=text.find('{')
    if start==-1: raise ValueError("Ziadny JSON v odpovedi")
    depth=0
    for i,c in enumerate(text[start:],start):
        if c=='{': depth+=1
        elif c=='}':
            depth-=1
            if depth==0: return json.loads(text[start:i+1])
    raise ValueError("Nekompletny JSON v odpovedi")

def analyze_email_with_claude(subject,body,sender_email,sender_name):
    p=f"""Si asistent eventovej agentury BigAgency. Analyzuj spravu.
Odosielatel: {sender_name} <{sender_email}>
Predmet: {subject}
Obsah: {body}

BigAgency je eventova agentura - organizuje eventy PRE klientov a pozicuje/prenajima vybavenie.

REALNY DOPYT (is_real_request=true) = klient alebo firma chce aby BigAgency nieco spravila PRE NICH:
- chce zorganizovat event, firemnu akciu, teambuilding, vianoce
- chce prenajat vybavenie (manez, ruske koleso, skaciacie hrady, ninja draha, tanecna podlaha, atrakcie...)
- pyta sa na cenu/ponuku za sluzby alebo vybavenie BigAgency
- agentura, firma alebo jednotlivec hlada BigAgency ako DODAVATELA pre ich event alebo projekt
- POZOR: aj firma z ineho odvetvia (osvetlenie, IT, vyroba...) moze byt klientom - rozhoduje to, ci PYTA sluzby od BigAgency, nie co tato firma robi

IGNORUJ (is_real_request=false) = niekto ponuka nieco BigAgency, alebo email nesuvisi s obchodom:
- hotely, priestory, lokality ktore PONUKAJU svoju lokalu na eventy BigAgency (oni su dodavatelia)
- vendori, dodavatelia ktori PONUKAJU svoje produkty alebo sluzby BigAgency
- konferencie, veltrhy ktore POZYVAJU BigAgency ako navstevnika alebo vystavovatelov
- uchadzaci o pracu, brigady, staze
- spam (rustina, turectina, loterie, kasino, kryptomeny)
- faktury, platobne pripomienky
- newslettery, marketing ktory nikto nepytal

JSON bez backticks: {{"is_real_request":true/false,"client_name":"meno alebo nazov firmy","event_description":"popis co chcu","event_date":"datum alebo null","task_name":"nazov tasku"}}"""
    r=anthropic.messages.create(model="claude-sonnet-5",max_tokens=1024,messages=[{"role":"user","content":p}])
    raw=r.content[0].text.strip().replace("```json","").replace("```","").strip()
    return extract_first_json(raw)

def generate_task_description(analysis,body,sender_email,sender_name,msg_id=""):
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
{body}

---
Dopyt prijaty: {today}
[MSG_ID: {msg_id}]
@{aname} prosim spracuj tuto ponuku a odpovedz klientovi na: {email_str}"""

def process_info_emails():
    logger.info("Spracuvam INFO Requesty emaily (vsetky za 30 dni)...")
    try:
        emails=get_emails_from_folder("INFO Requesty",limit=50,days_back=30)
        if not emails: logger.info("Ziadne emaily."); return
        workload=get_team_workload(); processed=0; skipped=0
        for email in emails:
            eid=email.get("id",""); subject=email.get("subject","")
            html_body=email.get("body",{}).get("content","")
            clean_body=strip_html(html_body) if html_body else email.get("bodyPreview","")
            sender_email,sender_name=get_client_contact(email,clean_body)
            msg_id=email.get("internetMessageId","") or eid
            try: analysis=analyze_email_with_claude(subject,clean_body[:3000],sender_email,sender_name)
            except Exception as e: logger.error(f"Analyza zlyhala: {e}"); continue
            if not analysis.get("is_real_request"):
                logger.info(f"Ignorovany (nie realny dopyt): {subject[:60]}"); continue
            task_name=analysis.get("task_name",subject[:100])
            if task_already_exists(msg_id, task_name):
                skipped+=1; continue
            aid,aname=get_less_busy_assignee(workload)
            analysis["assignee_name"]=aname
            if aid==MICHAL_ID: workload["michal"]["count"]+=1
            elif aid==MARTIN_ID: workload["martin"]["count"]+=1
            else: workload["stanislav"]["count"]+=1
            else: workload["stanislav"]["count"]+=1
            desc=generate_task_description(analysis,clean_body,sender_email,sender_name,msg_id)
            tid=create_task(task_name,desc,aid,"high")
            if tid:
                processed+=1; logger.info(f"OK: {sender_email} -> {aname}")
            else:
                logger.error(f"CHYBA vytvorenia tasku pre: {sender_email}")
        logger.info(f"Hotovo: {processed} novych taskov, {skipped} duplikatov preskoceno")
    except Exception as e: logger.error(f"Chyba procesu: {e}")

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
        sc=len([t for t in all_t if any(a["id"]==int(STANISLAV_ID) for a in t.get("assignees",[]))])
        logger.info(f"Report {now.strftime('%d.%m.%Y')}: Michal={mc} Stanislav={sc} Urgent={len(urg)}")
    except Exception as e: logger.error(f"Report: {e}")

def setup_schedule():
    schedule.every().day.at("08:30").do(check_deadlines)
    schedule.every().day.at("09:00").do(process_info_emails)
    schedule.every().day.at("10:00").do(check_stale_tasks)
    schedule.every().day.at("16:00").do(process_info_emails)
    schedule.every().monday.at("08:00").do(weekly_report)
    logger.info("Scheduler: 08:30 deadliny | 09:00+16:00 emaily | pon 08:00 report")

def main():
    logger.info("BigAgency AI Agent spusteny!")
    missing=[v for v in ["ANTHROPIC_API_KEY","CLICKUP_API_KEY","MS_CLIENT_ID","MS_CLIENT_SECRET","MS_TENANT_ID"] if not os.environ.get(v)]
    if missing: logger.error(f"Chybaju env var: {missing}"); return
    setup_schedule()
    logger.info("Cakam na ulohy... (prve spracovanie emailov o 09:00)")
    while True: schedule.run_pending(); time.sleep(60)

if __name__=="__main__": main()
