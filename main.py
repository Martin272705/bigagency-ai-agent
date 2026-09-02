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
PROFIPONUKA_API_KEY=os.environ.get("PROFIPONUKA_API_KEY")
PROFIPONUKA_BASE="https://app.profiponuka.sk/api"
PP_CURRENCY_EUR=2512
PP_LANGUAGE_SK=2282
PP_STATUS_CREATED=8505

def strip_html(t):
    if not t: return ""
    t=t.replace("&nbsp;"," ").replace("&amp;","&").replace("&lt;","<").replace("&gt;",">")
    t=re.sub(r'<style[^>]*>.*?</style>',' ',t,flags=re.DOTALL|re.IGNORECASE)
    t=re.sub(r'<script[^>]*>.*?</script>',' ',t,flags=re.DOTALL|re.IGNORECASE)
    t=re.sub(r'<br\s*/?>',' \n',t,flags=re.IGNORECASE)
    t=re.sub(r'</p>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'</div>','\n',t,flags=re.IGNORECASE)
    t=re.sub(r'<li[^>]*>','\n• ',t,flags=re.IGNORECASE)
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

def get_unread_emails_from_folder(folder_name="INFO Requesty",limit=15):
    token=get_ms_token()
    h={"Authorization":f"Bearer {token}"}
    folders=requests.get(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders",headers=h).json()
    fid=next((f["id"] for f in folders.get("value",[]) if f["displayName"].lower()==folder_name.lower()),None)
    if not fid: logger.error(f"Priecinok '{folder_name}' nenajdeny!"); return []
    since=(datetime.now()-timedelta(hours=26)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url=(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/mailFolders/{fid}/messages"
         f"?$filter=receivedDateTime ge {since}&$top={limit}&$orderby=receivedDateTime desc"
         f"&$select=id,subject,bodyPreview,body,from,sender,replyTo,receivedDateTime,internetMessageId")
    emails=requests.get(url,headers=h).json().get("value",[])
    logger.info(f"Najdenych {len(emails)} emailov za poslednych 26 hodin")
    return emails

def mark_email_as_read(eid,retries=3):
    for attempt in range(retries):
        try:
            token=get_ms_token()
            r=requests.patch(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/messages/{eid}",
                headers={"Authorization":f"Bearer {token}","Content-Type":"application/json"},json={"isRead":True})
            if r.status_code in (200,204):
                return True
            logger.warning(f"mark_as_read pokus {attempt+1}: HTTP {r.status_code}")
        except Exception as e:
            logger.warning(f"mark_as_read pokus {attempt+1} zlyhal: {e}")
        time.sleep(2)
    logger.error(f"Nepodarilo sa oznacit email ako precitany po {retries} pokusoch: {eid}")
    return False

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

def pp_get(ep,params=None):
    try:
        r=requests.get(f"{PROFIPONUKA_BASE}/{ep}",headers={"Authorization":PROFIPONUKA_API_KEY},params=params)
        return r.json()
    except Exception as e: logger.error(f"pp_get {ep}: {e}"); return {}

def pp_post(ep,data):
    try:
        r=requests.post(f"{PROFIPONUKA_BASE}/{ep}",headers={"Authorization":PROFIPONUKA_API_KEY},data=data)
        return r.json()
    except Exception as e: logger.error(f"pp_post {ep}: {e}"); return {}

def pp_find_or_create_customer(email,name):
    try:
        res=pp_get("customer",{"email":email})
        for c in res.get("records",[]):
            if c.get("email","").lower()==email.lower():
                logger.info(f"PP zakaznik najdeny: {c['id']}")
                return c["id"]
        ctype="COMPANY" if any(x in name.lower() for x in ["s.r.o","a.s.","spol","ltd","gmbh","a. s."]) else "PERSON"
        res=pp_post("customer",{"name":name,"email":email,"type":ctype})
        cid=res.get("id")
        logger.info(f"PP zakaznik vytvoreny: {cid}")
        return cid
    except Exception as e: logger.error(f"pp_find_or_create_customer: {e}"); return None

def pp_create_draft_quote(customer_id,description):
    try:
        today=datetime.now().strftime("%Y-%m-%d")
        valid_until=(datetime.now()+timedelta(days=30)).strftime("%Y-%m-%d")
        data={"idCustomer":customer_id,"idState":PP_STATUS_CREATED,
              "idCurrency":PP_CURRENCY_EUR,"idLanguage":PP_LANGUAGE_SK,
              "date":today,"dateValid":valid_until,"description":description}
        res=pp_post("price-quote",data)
        qid=res.get("id")
        if qid:
            url=f"https://app.profiponuka.sk/price-quote/{qid}"
            logger.info(f"PP ponuka vytvorena: {qid}")
            return qid,url
        logger.warning(f"PP ponuka failed: {res}")
        return None,None
    except Exception as e: logger.error(f"pp_create_draft_quote: {e}"); return None,None

def start_time_tracking(task_id):
    try:
        r=requests.post(f"https://api.clickup.com/api/v2/team/{CLICKUP_TEAM_ID}/time_entries/start",
            headers={"Authorization":CLICKUP_API_KEY,"Content-Type":"application/json"},
            json={"tid":task_id,"billable":False})
        if r.status_code in(200,201): logger.info(f"Time tracking: {task_id}"); return True
        logger.warning(f"Time tracking failed: {r.status_code}"); return False
    except Exception as e: logger.error(f"start_time_tracking: {e}"); return False

def send_email_via_graph(to_email,subject,body_text):
    try:
        token=get_ms_token()
        msg={"message":{"subject":subject,"body":{"contentType":"Text","content":body_text},"toRecipients":[{"emailAddress":{"address":to_email}}]}}
        r=requests.post(f"https://graph.microsoft.com/v1.0/users/{MS_USER_EMAIL}/sendMail",
            headers={"Authorization":f"Bearer {token}","Content-Type":"application/json"},json=msg)
        if r.status_code==202: logger.info(f"Email odoslany: {to_email}"); return True
        logger.warning(f"Email failed: {r.status_code} {r.text[:200]}"); return False
    except Exception as e: logger.error(f"send_email_via_graph: {e}"); return False

def task_already_exists(msg_id,sender_email):
    try:
        page=0
        while True:
            tasks=clickup_get(f"list/{CLICKUP_LIST_ID}/task?page={page}&include_closed=true").get("tasks",[])
            if not tasks: break
            for t in tasks:
                desc=(t.get("description") or "")
                if msg_id and f"[MSG_ID: {msg_id}]" in desc:
                    logger.warning(f"Duplikat (MSG_ID) - preskakujem")
                    return True
                if sender_email and sender_email.lower() in desc.lower():
                    logger.warning(f"Duplikat (email {sender_email}) - preskakujem")
                    return True
            if len(tasks)<100: break
            page+=1
    except Exception as e:
        logger.error(f"Chyba dedup: {e}")
    return False

def get_team_workload():
    m=len(clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={MICHAL_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    ma=len(clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={MARTIN_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    s=len(clickup_get(f"team/{CLICKUP_TEAM_ID}/task?assignees[]={STANISLAV_ID}&statuses[]=to do&statuses[]=in progress").get("tasks",[]))
    logger.info(f"Vytazenost: Michal={m}, Martin={ma}, Stanislav={s}")
    return {"michal":{"id":MICHAL_ID,"count":m},"martin":{"id":MARTIN_ID,"count":ma},"stanislav":{"id":STANISLAV_ID,"count":s}}

def get_less_busy_assignee(w):
    people=[("michal",MICHAL_ID,"Michal Macai"),("martin",MARTIN_ID,"Martin Cano"),("stanislav",STANISLAV_ID,"Stanislav Kois")]
    least=min(people,key=lambda x:w[x[0]]["count"])
    return least[1],least[2]

def create_task(name,desc,aid,priority="high",due_date=None):
    d={"name":name,"markdown_description":desc,"assignees":[int(aid)],"priority":2 if priority=="high" else 3,"status":"to do"}
    if due_date: d["due_date"]=int(due_date.timestamp()*1000)
    r=clickup_post(f"list/{CLICKUP_LIST_ID}/task",d)
    tid=r.get("id")
    logger.info(f"Task vytvoreny: {name} (ID:{tid})")
    if tid: start_time_tracking(tid)
    return tid

def add_comment_to_task(tid,c,assignee_id=None):
    d={"comment_text":c,"notify_all":True}
    if assignee_id: d["assignee"]=int(assignee_id)
    clickup_post(f"task/{tid}/comment",d)

def get_tasks_with_upcoming_deadlines(days=7):
    now=datetime.now(); db=now+timedelta(days=days)
    return clickup_get(f"team/{CLICKUP_TEAM_ID}/task?statuses[]=to do&statuses[]=in progress&due_date_gt={int(now.timestamp()*1000)}&due_date_lt={int(db.timestamp()*1000)}").get("tasks",[])

def analyze_email_with_claude(subject,body,sender_email,sender_name):
        p=f"""Si asistent eventovej agentury BigAgency. Analyzuj spravu.
Odosielatel: {sender_name} <{sender_email}>
Predmet: {subject}
Obsah: {body}
BigAgency je eventova agentura - organizuje eventy PRE klientov a pozicuje vybavenie.

REALNY DOPYT (is_real_request=true) = klient chce aby BigAgency nieco spravila PRE NEHO:
- chce zorganizovat event, firemnu akciu, teambuilding, sportovy turnaj, konferenciu
- chce prenajat vybavenie (skaciacie hrady, ninja draha, tanecna podlaha, stoly, stolicky...)
- pyta sa na cenu/ponuku za sluzby BigAgency
- firma/korporat hlada dodavatela na organizaciu eventu - aj ked pisu TENDER znamena ze hladaju dodavatela
- email z weboveho formulara (subject zacina "WEB kontakt") - vzdy realny dopyt pokial telo emailu nie je spam
- klient chce aby mu BigAgency zavolala (Zavolame Vam)

IGNORUJ (is_real_request=false):
- hotely/priestory ponukaju svoju lokalu BigAgency (oni predavaju, nie kupuju)
- konferencie/veltrhy pozyvaju BigAgency ako navstevnika/kupujuceho
- vendori/dodavatelia ponukaju svoje produkty/sluzby BigAgency
- brigady, uchadzaci o pracu
- spam (rustina, turectina, arabcina, loterie, kasino, darknet)
- faktury, bankove vypisy, automaticke notifikacie systemov
- newslettery ktore nikto nepytal
- app notifikacie ako "is in your phone contacts" alebo podobne

create_quote=true ak ide o konkretny dopyt na prenajom vybavenia alebo event kde mozno hned pripravit ponuku. false ak ide o TENDER alebo klient len pyta vseobecne info.

JSON bez backticks: {{"is_real_request":true/false,"create_quote":true/false,"client_name":"meno alebo nazov firmy","event_description":"popis co chcu","event_date":"datum alebo null","task_name":"nazov tasku"}}"""
    r=anthropic.messages.create(model="claude-sonnet-4-5-20250929",max_tokens=1024,messages=[{"role":"user","content":p}])
    raw=r.content[0].text.strip().replace("```json","").replace("```","").strip()
    obj,_=json.JSONDecoder().raw_decode(raw)
    return obj

def generate_task_description(analysis,body,sender_email,sender_name,msg_id="",quote_url=None):
    today=datetime.now().strftime("%d.%m.%Y")
    aname=analysis.get("assignee_name","")
    cname=analysis.get("client_name",sender_name or "N/A")
    email_str=sender_email if sender_email else "NEPODARILO SA ZISTIT"
    quote_section=f"\n---\n\n## Cenová ponuka v ProfiPonuke\n[Otvoriť a doplniť položky]({quote_url})" if quote_url else ""
    return f"""## Kontakt na klienta
**Meno:** {cname}
**Email:** {email_str}

---

## Popis dopytu
{analysis.get('event_description','N/A')}

**Termin:** {analysis.get('event_date') or 'neuvedeny'}
{quote_section}

---

## Kompletny text spravy
{body[:3000]}

---
Dopyt prijaty: {today}
[MSG_ID: {msg_id}]
@{aname} prosim spracuj tuto ponuku a odpovedz klientovi na: {email_str}"""

def process_info_emails():
    logger.info("Spracuvam INFO Requesty emaily...")
    try:
        emails=get_unread_emails_from_folder("INFO Requesty",limit=15)
        if not emails: logger.info("Ziadne nove emaily."); return
        workload=get_team_workload(); processed=0; skipped=0; errors=0
        for email in emails:
            eid=email.get("id",""); subject=email.get("subject","")
            html_body=email.get("body",{}).get("content","")
            clean_body=strip_html(html_body) if html_body else email.get("bodyPreview","")
            sender_email,sender_name=get_client_contact(email,clean_body)
            try: analysis=analyze_email_with_claude(subject,clean_body[:3000],sender_email,sender_name)
            except Exception as e: logger.error(f"Analyza: {e}"); mark_email_as_read(eid); errors+=1; continue
            if not analysis.get("is_real_request"):
                logger.info(f"Ignorovany: {subject[:50]}"); mark_email_as_read(eid); continue
            task_name=analysis.get("task_name",subject[:100])
            msg_id=email.get("internetMessageId","") or eid
            if task_already_exists(msg_id,sender_email):
                skipped+=1; mark_email_as_read(eid); continue
            aid,aname=get_less_busy_assignee(workload)
            analysis["assignee_name"]=aname
            if aid==MICHAL_ID: workload["michal"]["count"]+=1
            elif aid==MARTIN_ID: workload["martin"]["count"]+=1
            else: workload["stanislav"]["count"]+=1
            # ProfiPonuka
            quote_url=None
            if analysis.get("create_quote") and PROFIPONUKA_API_KEY:
                cname=analysis.get("client_name",sender_name or "")
                cid=pp_find_or_create_customer(sender_email,cname) if sender_email else None
                if cid:
                    _,quote_url=pp_create_draft_quote(cid,analysis.get("event_description",""))
            desc=generate_task_description(analysis,clean_body[:3000],sender_email,sender_name,msg_id,quote_url)
            tid=create_task(task_name,desc,aid,"high")
            if tid:
                processed+=1; logger.info(f"OK: {sender_email} -> {aname}" + (f" | PP: {quote_url}" if quote_url else ""))
                # Email kolegovi
                assignee_email=ASSIGNEE_EMAILS.get(str(aid))
                if assignee_email:
                    task_url=f"https://app.clickup.com/t/{tid}"
                    aname_short=aname.split()[0]
                    pp_line=f"\nCenová ponuka (doplňte položky): {quote_url}" if quote_url else ""
                    body_email=f"Ahoj {aname_short},\n\nBol ti priradený nový dopyt od klienta.\n\nZákazník: {analysis.get('client_name',sender_email)}\nPopis: {analysis.get('event_description','')[:200]}\n\nTask v ClickUp: {task_url}{pp_line}\n\nBigAgency AI Agent"
                    send_email_via_graph(assignee_email,f"Nový dopyt: {task_name}",body_email)
                mark_email_as_read(eid)
            else:
                logger.error(f"CHYBA: task pre {sender_email} sa nepodarilo vytvorit - email zostava unread"); errors+=1
        logger.info(f"Spracovanych {processed}, preskoceno duplikatov: {skipped}")
        cas=datetime.now().strftime("%H:%M")
        icon="✅" if errors==0 else "⚠️"
        subj=f"{icon} Agent {cas} — nov\xe9 tasky: {processed}, preskočen\xfdch: {skipped}"
        body_sum=f"Beh {cas}:\n\nNov\xe9 ClickUp tasky: {processed}\nPreskočen\xe9 (duplik\xe1ty): {skipped}\nChyby pri anal\xfdze: {errors}\nCelkovo emailov: {len(emails)}"
        send_email_via_graph(MS_USER_EMAIL,subj,body_sum)
    except Exception as e: logger.error(f"Chyba: {e}")

def check_unprocessed_tasks():
    """24h a 72h reminder pre tasky v to do statuse."""
    logger.info("Kontrolujem nespracovane tasky...")
    try:
        tasks=clickup_get(f"list/{CLICKUP_LIST_ID}/task?statuses[]=to do&include_closed=false").get("tasks",[])
        r24=0;r72=0
        for task in tasks:
            tid=task.get("id");tname=task.get("name","");assignees=task.get("assignees",[])
            task_url=f"https://app.clickup.com/t/{tid}"
            hours_old=(datetime.now()-datetime.fromtimestamp(int(task.get("date_created",0))/1000)).total_seconds()/3600
            if hours_old<24: continue
            ctexts=[(c.get("comment_text") or "") for c in clickup_get(f"task/{tid}/comment").get("comments",[])]
            has24=any("24h REMINDER" in t for t in ctexts)
            has72=any("72h REMINDER" in t for t in ctexts)
            if hours_old>=72 and not has72:
                add_comment_to_task(tid,f"Dopyt caka uz {int(hours_old)} hodin! Klient stale caka na odpoved. Riesit urgentne. [72h REMINDER]")
                for a in assignees:
                    em=ASSIGNEE_EMAILS.get(str(a.get("id","")))
                    if em:
                        aname=a.get("username","").split()[0]
                        body="\n".join([f"Ahoj {aname},","",f"Dopyt '{tname}' caka uz {int(hours_old)} hodin bez spracovania!","","Klient stale caka na odpoved. Prosim riesit urgentne.","","Ak si uz zacel/a, zmen prosim status v ClickUp na 'In Progress'.",f"Link: {task_url}","","BigAgency AI Agent"])
                        send_email_via_graph(em,f"URGENTNE - Nespracovany dopyt: {tname}",body)
                r72+=1
            elif hours_old>=24 and not has24:
                add_comment_to_task(tid,f"Dopyt caka na spracovanie uz {int(hours_old)} hodin a stale je v stave to do. [24h REMINDER]")
                for a in assignees:
                    em=ASSINGEE_EMAILS.get(str(a.get("id","")))
                    if em:
                        aname=a.get("username","").split()[0]
                        body="\n".join([f"Ahoj {aname},","",f"Dopyt '{tname}' v ClickUp caka uz {int(hours_old)} hodin na spracovanie.","","Prosim skontroluj a odpovedz klientovi co najskor.",f"Link: {task_url}","","BigAgency AI Agent"])
                        send_email_via_graph(em,f"Nespracovany dopyt v ClickUp: {tname}",body)
                r24+=1
        logger.info(f"Remindery: 24h={r24}, 72h={r72}")
    except Exception as e: logger.error(f"check_unprocessed_tasks: {e}")

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
        ma=len([t for t in all_t if any(a["id"]==int(MARTIN_ID) for a in t.get("assignees",[]))])
        sc=len([t for t in all_t if any(a["id"]==int(STANISLAV_ID) for a in t.get("assignees",[]))])
        logger.info(f"Report {now.strftime('%d.%m.%Y')}: Michal={mc} Martin={ma} Stanislav={sc} Urgent={len(urg)}")
    except Exception as e: logger.error(f"Report: {e}")

def setup_schedule():
    schedule.every().day.at("08:30").do(check_deadlines)
    schedule.every().day.at("09:00").do(process_info_emails)
    schedule.every().day.at("10:00").do(check_unprocessed_tasks)
    schedule.every().day.at("16:00").do(process_info_emails)
    schedule.every().monday.at("08:00").do(weekly_report)
    logger.info("Scheduler: 08:30 deadliny | 09:00+16:00 emaily | 10:00 remindery | pon 08:00 report")

def main():
    logger.info("BigAgency AI Agent spusteny!")
    missing=[v for v in ["ANTHROPIC_API_KEY","CLICKUP_API_KEY","MS_CLIENT_ID","MS_CLIENT_SECRET","MS_TENANT_ID"] if not os.environ.get(v)]
    if missing: logger.error(f"Chybaju: {missing}"); return
    if not PROFIPONUKA_API_KEY: logger.warning("PROFIPONUKA_API_KEY nie je nastaveny - PP integrácia vypnuta")
    setup_schedule()
    logger.info("Cakam na ulohy...")
    while True: schedule.run_pending(); time.sleep(60)

if __name__=="__main__": main()
