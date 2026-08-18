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
