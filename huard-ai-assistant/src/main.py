from contextlib import asynccontextmanager
from fastapi import FastAPI, Form, Request, HTTPException
from fastapi.responses import PlainTextResponse, JSONResponse

from src.config import settings
from src.core.knowledge_base import KnowledgeBase
from src.core.claude_client import ClaudeClient
from src.core.conversation import ConversationStore
from src.channels.whatsapp_twilio import WhatsAppTwilioChannel
from src.channels.outlook import OutlookChannel


@asynccontextmanager
async def lifespan(app: FastAPI):
    kb = KnowledgeBase(settings.knowledge_dir)
    claude = ClaudeClient(kb)
    conv = ConversationStore()
    app.state.kb = kb
    app.state.claude = claude
    app.state.conv = conv
    app.state.wa_twilio = WhatsAppTwilioChannel(claude, conv)
    app.state.outlook = OutlookChannel(claude)
    yield


app = FastAPI(title="HUARD AI Assistant", version="0.1.0", lifespan=lifespan)


@app.get("/health")
async def health():
    kb: KnowledgeBase = app.state.kb
    return {
        "status": "ok",
        "env": settings.app_env,
        "knowledge": kb.summary(),
        "channels": {
            "outlook_configured": bool(settings.ms_client_id and settings.ms_client_secret),
            "twilio_configured": bool(settings.twilio_account_sid and settings.twilio_auth_token),
            "wa_perso_enabled": settings.wa_perso_enabled,
        },
    }


@app.post("/webhooks/twilio/whatsapp", response_class=PlainTextResponse)
async def twilio_whatsapp_webhook(
    request: Request,
    From: str = Form(...),
    Body: str = Form(...),
):
    """Webhook appelé par Twilio à chaque message WhatsApp entrant.

    Twilio attend une réponse TwiML XML pour répondre dans le même HTTP cycle.
    """
    wa: WhatsAppTwilioChannel = app.state.wa_twilio
    try:
        reply = await wa.handle_incoming(from_number=From, body=Body)
    except Exception as e:
        return PlainTextResponse(
            f"<Response><Message>Désolé, problème technique : {e}</Message></Response>",
            media_type="application/xml",
        )
    escaped = reply.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return PlainTextResponse(
        f"<Response><Message>{escaped}</Message></Response>",
        media_type="application/xml",
    )


@app.post("/outlook/scan-and-draft")
async def outlook_scan_and_draft(top: int = 10):
    """Endpoint à déclencher (cron, bouton) : lit les N derniers mails non lus
    et crée un brouillon de réponse pour chacun."""
    outlook: OutlookChannel = app.state.outlook
    if not (settings.ms_client_id and settings.ms_client_secret):
        raise HTTPException(status_code=400, detail="Microsoft Graph non configuré")

    emails = await outlook.list_unread(top=top)
    results = []
    for email in emails:
        try:
            draft_id = await outlook.process_one(email)
            results.append({
                "message_id": email.message_id,
                "from": email.sender_email,
                "subject": email.subject,
                "draft_id": draft_id,
                "status": "draft_created",
            })
        except Exception as e:
            results.append({
                "message_id": email.message_id,
                "from": email.sender_email,
                "subject": email.subject,
                "status": "error",
                "error": str(e),
            })
    return JSONResponse({"processed": len(results), "results": results})


@app.post("/test/claude")
async def test_claude(request: Request):
    """Endpoint utilitaire pour tester rapidement Claude + la base de connaissances."""
    payload = await request.json()
    question = payload.get("question", "")
    claude: ClaudeClient = app.state.claude
    from src.core.claude_client import ChatMessage
    reply = await claude.reply(
        messages=[ChatMessage(role="user", content=question)],
        max_tokens=800,
    )
    return {"reply": reply}
