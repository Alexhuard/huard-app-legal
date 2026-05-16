"""Canal WhatsApp Business via Twilio (API officielle Meta).

Twilio expose un webhook HTTP : à chaque message reçu sur le numéro WhatsApp
Business HUARD, Twilio POST une requête form-encoded sur notre endpoint. On
répond soit en TwiML synchrone (pour les réponses courtes), soit en appelant
l'API Twilio en async (pour les réponses longues).
"""

from twilio.rest import Client as TwilioClient

from src.config import settings
from src.core.claude_client import ClaudeClient, ChatMessage
from src.core.conversation import ConversationStore


WHATSAPP_INSTRUCTIONS = (
    "# Canal : WhatsApp interne HUARD\n"
    "- Réponses TRÈS courtes, façon SMS : 1 à 4 phrases max.\n"
    "- Pas de markdown, pas de listes à puces complexes.\n"
    "- Tutoiement si l'interlocuteur tutoie, vouvoiement sinon.\n"
    "- Si la question demande plus de 5 lignes de réponse, propose plutôt\n"
    "  d'appeler ou de prendre rdv avec la personne compétente HUARD.\n"
)


class WhatsAppTwilioChannel:
    def __init__(self, claude: ClaudeClient, conversation_store: ConversationStore):
        self.claude = claude
        self.conv = conversation_store
        self._twilio: TwilioClient | None = None

    def _client(self) -> TwilioClient:
        if self._twilio is None:
            self._twilio = TwilioClient(
                settings.twilio_account_sid, settings.twilio_auth_token
            )
        return self._twilio

    async def handle_incoming(self, from_number: str, body: str) -> str:
        """Reçoit un message, met à jour l'historique, génère et renvoie la réponse."""
        self.conv.append("wa_twilio", from_number, ChatMessage(role="user", content=body))
        history = self.conv.history("wa_twilio", from_number)

        reply = await self.claude.reply(
            messages=history,
            channel_instructions=WHATSAPP_INSTRUCTIONS,
            fast=True,
            max_tokens=500,
        )
        self.conv.append("wa_twilio", from_number, ChatMessage(role="assistant", content=reply))
        return reply

    def send_message(self, to_number: str, text: str) -> str:
        message = self._client().messages.create(
            from_=settings.twilio_whatsapp_from,
            to=to_number if to_number.startswith("whatsapp:") else f"whatsapp:{to_number}",
            body=text,
        )
        return message.sid
