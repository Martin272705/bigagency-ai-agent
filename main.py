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
ASSIGNEE_EMAILS={"106588503":"macai@bigagency.sk","284426112":"cano@bigagency.sk","106588288":"kois@bigagency.sk"}

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

def get_emails_from_folder(folder_name="INFO Requesty",limit=50):
    token=get_ms_token()
    h={"Authorization":f"Bearer {token}"}
    folders=requests.get(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders",headers=h).json()
    fid=next((f["id"] for f in folders.get("value",[]) if f["displayName"].lower()==folder_name.lower()),None)
    if not fid: logger.error(f"Priecinok '{folder_name}' nenajdeny!"); return []
    since=(datetime.now()-timedelta(hours=25)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url=(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders/{fid}/messages"
         f"?$filter=receivedDateTime ge {since}&$top={limit}&$orderby=receivedDateTime desc"
         f"&$select=id,subject,bodyPreview,body,from,sender,replyTo,receivedDateTime,internetMessageId")
    emails=requests.get(url,headers=h).json().get("value",[])
    logger.info(f"Najdenych {len(emails)} emailov za poslednych 25 hodin")
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

def start_time_tracking(task_id):
    try:
        r=requests.post(f"https://api.clickup.com/api/v2/team/{CLICKUP_TEAM_ID}/time_entries/start",headers={"Authorization":CLICKUP_API_KEY,"Content-Type":"application/json"},json={"tid":task_id,"billable":False})
        if r.status_code in(200,201): logger.info(f"Time tracking: {task_id}"); return True
        logger.warning(f"Time tracking failed: {r.status_code}"); return False
    except Exception as e: logger.error(f"start_time_tracking: {e}"); return False

def send_email_via_graph(to_email,subject,body_text):
    try:
        token=get_ms_token()
        msg={"message":{"subject":subject,"body":{"contentType":"Text","content":body_text},"toRecipients":[{"emailAddress":{"address":to_email}}]}}
        r=requests.post(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/sendMail",headers={"Authorization":f"Bearer {token}","Content-Type":"application/json"},json=msg)
        if r.status_code==202: logger.info(f"Email odoslany: {to_email}"); return True
        logger.warning(f"Email failed: {r.status_code} {r.text[:200]}"); return False
    except Exception as e: logger.error(f"send_email_via_graph: {e}"); return False

def task_already_exists(msg_id,sender_email):
    # POZNAMKA: ClickUp search API nehladá v popisoch taskov - len v názvoch.
    # Preto čítame tasky priamo z nášho listu a kontrolujeme popis manuálne.
    try:
        cutoff=int((datetime.now()-timedelta(days=90)).timestamp()*1000)
        page=0
        while True:
            r=requests.get(
                f"https://api.clickup.com/api/v2/list/{CLICKUP_LIST_ID}/task",
                headers={"Authorization":CLICKUP_API_KEY},
                params={"date_created_gt":cutoff,"include_closed":"true","subtasks":"false","page":page}
            )
            tasks=r.json().get("tasks",[])
            if not tasks: break
            for task in tasks:
                desc=task.get("description","")
                if msg_id and msg_id in desc:
                    logger.warning(f"Duplikat (MSG_ID v popise) - preskakujem: {task.get('name','')[:50]}")
                    return True
                if sender_email and "noreply" not in sender_email.lower():
                    if sender_email.lower() in desc.lower():
                        logger.warning(f"Duplikat (email {sender_email} v popise) - preskakujem: {task.get('name','')[:50]}")
                        return True
            if len(tasks)<100: break
            page+=1
    except Exception as e:
        logger.error(f"Chyba dedup: {e}")
    return False

def get_team_workload():
    m=len(clickup_get(f"list/{CLICKUP_LIST_ID}/task?assignees[]={MICHAL_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    ma=len(clickup_get(f"list/{CLICKUP_LIST_ID}/task?assignees[]={MARTIN_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    s=len(clickup_get(f"list/{CLICKUP_LIST_ID}/task?assignees[]={STANISLAV_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    logger.info(f"Vytazenost: Michal={m}, Martin={ma}, Stanislav={s}")
    return {"michal":{"id":MICHAL_ID,"count":m},"martin":{"id":MARTIN_ID,"count":ma},"stanislav":{"id":STANISLAV_ID,"count":s}}

def get_less_busy_assignee(w):
    person=min(w,key=lambda k:w[k]["count"])
    names={"michal":"Michal Macai","martin":"Martin Cano","stanislav":"Stanislav Kois"}
    return w[person]["id"],names[person],person

def create_task(name,desc,aid,priority="high",due_date=None):
    d={"name":name,"markdown_description":desc,"assignees":[int(aid)],"priority":2 if priority=="high" else 3,"status":"to do"}
    if due_date: d["due_date"]=int(due_date.timestamp()*1000)
    r=clickup_post(f"list/{CLICKUP_LIST_ID}/task",d)
    tid=r.get("id")
    logger.info(f"Task vytvoreny: {name} (ID:{tid})")
    if tid: start_time_tracking(tid)
    return tid

def add_comment_to_task(tid,c): clickup_post(f"task/{tid}/comment",{"comment_text":c})

def send_status_email(subject,body):
    try:
        token=get_ms_token()
        requests.post(
            f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/sendMail",
            headers={"Authorization":f"Bearer {token}","Content-Type":"application/json"},
            json={"message":{"subject":subject,"body":{"contentType":"Text","content":body},"toRecipients":[{"emailAddress":{"address":"cano@bigagency.sk"}}]}}
        )
        logger.info(f"Notifikacia odoslana: {subject}")
    except Exception as e:
        logger.error(f"send_status_email chyba: {e}")

def check_stale_tasks():
    logger.info("Kontrola stale taskov (48h)...")
    try:
        cutoff=int((datetime.now()-timedelta(hours=48)).timestamp()*1000)
        tasks=clickup_get(f"list/{CLICKUP_LIST_ID}/task?statuses[]=to do&date_created_lt={cutoff}").get("tasks",[])
        reminded=0
        for task in tasks:
            tid=task.get("id")
            comments=clickup_get(f"task/{tid}/comment").get("comments",[])
            if any("[AUTO-REMINDER]" in c.get("comment_text","") for c in comments):
                continue
            assignees=task.get("assignees",[])
            names=", ".join([f"@{a.get('username','')}" for a in assignees])
            add_comment_to_task(tid,f"[AUTO-REMINDER] Tento dopyt caka uz viac ako 48 hodin. {names} prosim spracuj ponuku a odpoved klientovi.")
            reminded+=1
            logger.info(f"AUTO-REMINDER: {task.get('name','')}")
        logger.info(f"Stale check hotovy, {reminded} reminderov pridanych")
    except Exception as e:
        logger.error(f"check_stale_tasks chyba: {e}")

def analyze_email_with_claude(subject,body,sender_email,sender_name):
    p=f"""Si asistent eventovej agentury BigAgency na Slovensku. Analyzuj email a urc ci ide o realny klientsky dopyt.

Odosielatel: {sender_name} <{sender_email}>
Predmet: {subject}
Obsah: {body}

VYTVOR TASK ak:
- Klient chce aby BigAgency nieco spravila PRE NEHO (event, prenajom vybavenia, cenova ponuka)
- Email moze byt v slovenčine AJ v angličtine - oboje je OK!

IGNORUJ (is_real_request: false):
- Spam (rustina, turectina, nepochopitelny obsah)
- Faktury ktore dostava BigAgency od dodavatelov (poznas ich podla: "faktura", "invoice", "platba", obsahuju cislo faktury)
- Dodavatelia/vendori/hotely/priestory ktori PONUKAJU svoje sluzby BigAgency - obrateny smer!
- Brigadnici a uchadzaci o pracu (hostesky, technici, pomocnici)
- Newslettery, marketingove emaily, hromadne rozosielky
- Podakovanie za ponuku alebo odmietnutie ponuky
- Emaily kde BigAgency je len v kopia (CC) a dopyt nie je priamo na BigAgency

Odpoved ako JSON bez backticks:
{{"is_real_request":true/false,"client_name":"meno klienta","event_description":"strucny popis dopytu","event_date":"datum alebo null","task_name":"nazov tasku v ClickUp"}}"""
    r=anthropic.messages.create(model="claude-sonnet-5",max_tokens=1024,messages=[{"role":"user","content":p}])
    text_block=next((b for b in r.content if hasattr(b,'text')),None)
    if not text_block: raise ValueError("No text block in Claude response")
    raw=text_block.text.strip()
    m=re.search(r'\{.*\}',raw,re.DOTALL)
    if m: raw=m.group(0)
    return json.loads(raw)

def generate_task_description(analysis,body,sender_email,sender_name,msg_id=""):
    today=datetime.now().strftime("%d.%m.%Y")
    aname=analysis.get("assignee_name","")
    cname=analysis.get("client_name",sender_name or "N/A")
    email_str=sender_email if sender_email else "NEPODARILO SA ZISTIT"
    return (
        "## Kontakt na klienta\n"
        f"**Meno:** {cname}\n"
        f"**Email:** {email_str}\n\n"
        "---\n\n"
        "## Popis dopytu\n"
        f"{analysis.get('event_description','N/A')}\n\n"
        f"**Termin:** {analysis.get('event_date') or 'neuvedeny'}\n\n"
        "---\n\n"
        "## Kompletny text spravy\n"
        f"{body}\n\n"
        "---\n"
        f"Dopyt prijaty: {today}\n"
        f"[MSG_ID: {msg_id}]\n"
        f"@{aname} prosim spracuj tuto ponuku a odpovedz klientovi na: {email_str}"
    )

def process_info_emails():
    cas=datetime.now().strftime("%H:%M")
    logger.info("Spracuvam INFO Requesty emaily...")
    try:
        emails=get_emails_from_folder("INFO Requesty",limit=50)
        if not emails:
            logger.info("Ziadne emaily.")
            send_status_email(f"✅ Agent {cas} — žiadne nové emaily","Za posledných 25 hodín neboli žiadne nové emaily v INFO Requesty.")
            return
        workload=get_team_workload()
        processed=0
        skipped=0
        errors=0
        for email in emails:
            eid=email.get("id","")
            subject=email.get("subject","")
            html_body=email.get("body",{}).get("content","")
            clean_body=strip_html(html_body) if html_body else email.get("bodyPreview","")
            sender_email,sender_name=get_client_contact(email,clean_body)
            try:
                analysis=analyze_email_with_claude(subject,clean_body[:3000],sender_email,sender_name)
            except Exception as e:
                logger.error(f"Analyza: {e}")
                errors+=1
                continue
            if not analysis.get("is_real_request"):
                logger.info(f"Ignorovany: {subject[:50]}")
                continue
            task_name=analysis.get("task_name",subject[:100])
            msg_id=email.get("internetMessageId","") or eid
            if task_already_exists(msg_id,sender_email):
                skipped+=1
                continue
            aid,aname,person_key=get_less_busy_assignee(workload)
            analysis["assignee_name"]=aname
            workload[person_key]["count"]+=1
            desc=generate_task_description(analysis,clean_body,sender_email,sender_name,msg_id)
            tid=create_task(task_name,desc,aid,"high")
            if tid:
                processed+=1
                logger.info(f"OK: {sender_email} -> {aname}")
            else:
                logger.error(f"CHYBA: task pre {sender_email} sa nepodarilo vytvorit")
                errors+=1
        logger.info(f"Spracovanych {processed}, preskoceno duplikatov: {skipped}")
        status="✅" if errors==0 else "⚠️"
        send_status_email(
            f"{status} Agent {cas} — nové tasky: {processed}, preskočených: {skipped}",
            f"Beh {cas}:\n\nNové ClickUp tasky: {processed}\nPreskočené (duplikáty): {skipped}\nChyby pri analýze: {errors}\nCelkovo emailov: {len(emails)}"
        )
    except Exception as e:
        logger.error(f"Chyba: {e}")
        send_status_email(f"❌ Agent {cas} CHYBA",f"Agent padol s chybou:\n{e}\n\nSkontroluj Railway logy.")

def setup_schedule():
    schedule.every().day.at("09:00").do(process_info_emails)
    schedule.every().day.at("10:00").do(check_stale_tasks)
    schedule.every().day.at("16:00").do(process_info_emails)
    logger.info("Scheduler: 09:00+16:00 emaily | 10:00 stale tasky")

def main():
    logger.info("BigAgency AI Agent spusteny!")
    missing=[v for v in ["ANTHROPIC_API_KEY","CLICKUP_API_KEY","MS_CLIENT_ID","MS_CLIENT_SECRET","MS_TENANT_ID"] if not os.environ.get(v)]
    if missing: logger.error(f"Chybaju: {missing}"); return
    setup_schedule()
    logger.info("Cakam na ulohy...")
    while True: schedule.run_pending(); time.sleep(60)

if __name__=="__main__": main()
