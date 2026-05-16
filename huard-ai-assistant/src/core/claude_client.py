from dataclasses import dataclass
from anthropic import AsyncAnthropic

from src.config import settings
from src.core.knowledge_base import KnowledgeBase


SYSTEM_BASE = """Tu es l'assistant IA interne du Groupe HUARD.

Tu réponds au nom d'HUARD, avec un ton professionnel, clair, courtois et concis,
en français par défaut (sauf si l'interlocuteur écrit dans une autre langue).

Règles strictes :
- Tu ne dis JAMAIS que tu es une IA, sauf si on te le demande explicitement.
- Tu ne promets jamais un délai, un prix, ou un engagement contractuel sans qu'il
  soit explicitement présent dans la base de connaissances ci-dessous.
- Si tu ne sais pas répondre avec certitude, tu dis clairement que tu transmets
  la demande à un collaborateur HUARD, et tu ne devines pas.
- Si la question semble urgente, sensible (litige, réclamation, RH, juridique)
  ou hors de ton périmètre, tu n'improvises pas : tu indiques qu'un humain
  prendra le relais rapidement.
- Tu cites les procédures internes uniquement par leur sens, jamais en
  reproduisant des passages confidentiels mot pour mot vers l'extérieur.
"""


@dataclass
class ChatMessage:
    role: str  # "user" ou "assistant"
    content: str


class ClaudeClient:
    def __init__(self, knowledge_base: KnowledgeBase):
        self.client = AsyncAnthropic(api_key=settings.anthropic_api_key)
        self.kb = knowledge_base

    def _system_blocks(self, channel_instructions: str | None) -> list[dict]:
        """Construit le system prompt en 3 blocs, dont 2 sont cacheables.

        Ordre des blocs :
          1. Instructions de base (cacheable, change rarement)
          2. Base de connaissances HUARD (cacheable, change quand on édite les .md)
          3. Instructions spécifiques au canal (non cachées, ex: WhatsApp vs Email)
        """
        blocks: list[dict] = [
            {
                "type": "text",
                "text": SYSTEM_BASE,
                "cache_control": {"type": "ephemeral"},
            },
            {
                "type": "text",
                "text": (
                    "# Base de connaissances HUARD\n\n"
                    "Voici l'ensemble des procédures, FAQ et règles internes.\n"
                    "Appuie-toi exclusivement sur ce contenu pour répondre :\n\n"
                    + self.kb.render_for_prompt()
                ),
                "cache_control": {"type": "ephemeral"},
            },
        ]
        if channel_instructions:
            blocks.append({"type": "text", "text": channel_instructions})
        return blocks

    async def reply(
        self,
        messages: list[ChatMessage],
        channel_instructions: str | None = None,
        fast: bool = False,
        max_tokens: int = 1024,
    ) -> str:
        model = settings.claude_model_fast if fast else settings.claude_model
        api_messages = [{"role": m.role, "content": m.content} for m in messages]

        resp = await self.client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=self._system_blocks(channel_instructions),
            messages=api_messages,
        )
        parts: list[str] = []
        for block in resp.content:
            if getattr(block, "type", None) == "text":
                parts.append(block.text)
        return "".join(parts).strip()

    async def draft_email_reply(
        self,
        subject: str,
        body: str,
        sender_name: str | None = None,
        sender_email: str | None = None,
        thread_context: str | None = None,
    ) -> str:
        instructions = (
            "# Canal : Email Outlook (pré-réponse)\n"
            "Tu rédiges un BROUILLON de réponse email professionnel pour Alex HUARD.\n"
            "- Format : texte simple, sans markdown, prêt à coller dans Outlook.\n"
            "- Démarre par une formule d'appel adaptée (Bonjour [Prénom], etc.).\n"
            "- Termine par la signature standard HUARD si disponible dans templates/.\n"
            "- Si la demande nécessite une info que tu n'as pas, écris explicitement\n"
            "  entre crochets : [À COMPLÉTER PAR ALEX : ...]\n"
            "- Ne signe JAMAIS d'engagement (prix, délai, contrat) sans donnée fiable."
        )
        user_payload = (
            f"De : {sender_name or '?'} <{sender_email or '?'}>\n"
            f"Sujet : {subject}\n\n"
            f"--- Corps du mail reçu ---\n{body}\n"
        )
        if thread_context:
            user_payload += f"\n--- Contexte du fil ---\n{thread_context}\n"
        user_payload += "\n--- Rédige le brouillon de réponse ci-dessous ---"

        return await self.reply(
            messages=[ChatMessage(role="user", content=user_payload)],
            channel_instructions=instructions,
            max_tokens=1500,
        )
