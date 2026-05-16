"""Canal Outlook / Microsoft 365.

Stratégie : on ne répond PAS automatiquement. On crée un BROUILLON (draft) dans
la boîte de réponse à chaque mail entrant. Alex relit, ajuste, envoie d'un clic.

Flow :
  1. App Azure AD (créée à part) avec permissions Mail.ReadWrite déléguées ou
     applicatives (selon mode tenant).
  2. Au démarrage, on s'authentifie via MSAL (client credentials flow).
  3. On poll la boîte ou on s'abonne aux notifications Graph (push).
  4. Pour chaque nouveau message non lu sans brouillon associé : on appelle
     ClaudeClient.draft_email_reply() et on crée le draft via Graph API.

Note : ce fichier contient le squelette. L'auth Graph et le polling/webhook
seront branchés une fois les credentials Azure AD configurés.
"""

from dataclasses import dataclass
import httpx
import msal

from src.config import settings
from src.core.claude_client import ClaudeClient


GRAPH_API = "https://graph.microsoft.com/v1.0"


@dataclass
class IncomingEmail:
    message_id: str
    subject: str
    body_text: str
    sender_name: str
    sender_email: str
    conversation_id: str


class OutlookChannel:
    def __init__(self, claude: ClaudeClient):
        self.claude = claude
        self._token: str | None = None
        self._msal_app: msal.ConfidentialClientApplication | None = None

    def _ensure_msal(self) -> msal.ConfidentialClientApplication:
        if self._msal_app is None:
            self._msal_app = msal.ConfidentialClientApplication(
                client_id=settings.ms_client_id,
                client_credential=settings.ms_client_secret,
                authority=f"https://login.microsoftonline.com/{settings.ms_tenant_id}",
            )
        return self._msal_app

    async def _get_token(self) -> str:
        app = self._ensure_msal()
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            raise RuntimeError(f"MS Graph auth failed: {result.get('error_description')}")
        self._token = result["access_token"]
        return self._token

    async def list_unread(self, top: int = 10) -> list[IncomingEmail]:
        token = await self._get_token()
        url = (
            f"{GRAPH_API}/users/{settings.ms_user_email}/mailFolders/inbox/messages"
            f"?$filter=isRead eq false&$top={top}"
            f"&$select=id,subject,bodyPreview,body,from,conversationId"
        )
        async with httpx.AsyncClient(timeout=30) as http:
            resp = await http.get(url, headers={"Authorization": f"Bearer {token}"})
            resp.raise_for_status()
            data = resp.json()

        emails: list[IncomingEmail] = []
        for item in data.get("value", []):
            sender = item.get("from", {}).get("emailAddress", {})
            body = item.get("body", {})
            emails.append(
                IncomingEmail(
                    message_id=item["id"],
                    subject=item.get("subject", ""),
                    body_text=body.get("content", "") if body.get("contentType") == "text"
                    else _html_to_text(body.get("content", "")),
                    sender_name=sender.get("name", ""),
                    sender_email=sender.get("address", ""),
                    conversation_id=item.get("conversationId", ""),
                )
            )
        return emails

    async def create_reply_draft(self, email: IncomingEmail, draft_body: str) -> str:
        """Crée un brouillon de réponse attaché au message original."""
        token = await self._get_token()
        url = (
            f"{GRAPH_API}/users/{settings.ms_user_email}/messages/{email.message_id}/createReply"
        )
        async with httpx.AsyncClient(timeout=30) as http:
            resp = await http.post(url, headers={"Authorization": f"Bearer {token}"})
            resp.raise_for_status()
            draft = resp.json()
            draft_id = draft["id"]

            update_url = f"{GRAPH_API}/users/{settings.ms_user_email}/messages/{draft_id}"
            patch_resp = await http.patch(
                update_url,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                },
                json={"body": {"contentType": "text", "content": draft_body}},
            )
            patch_resp.raise_for_status()
        return draft_id

    async def process_one(self, email: IncomingEmail) -> str:
        draft_body = await self.claude.draft_email_reply(
            subject=email.subject,
            body=email.body_text,
            sender_name=email.sender_name,
            sender_email=email.sender_email,
        )
        return await self.create_reply_draft(email, draft_body)


def _html_to_text(html: str) -> str:
    """Conversion HTML→texte basique. À remplacer par BeautifulSoup si besoin."""
    import re
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"\s+", " ", text)
    return text.strip()
