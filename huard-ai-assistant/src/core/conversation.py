from src.core.claude_client import ChatMessage


class ConversationStore:
    """Mémoire conversationnelle simple en RAM, indexée par (canal, id_externe).

    Pour la prod on remplacera par SQLite (voir src/storage/db.py). Pour
    démarrer rapidement, du in-memory suffit côté POC.
    """

    def __init__(self, max_turns: int = 20):
        self.max_turns = max_turns
        self._store: dict[str, list[ChatMessage]] = {}

    def _key(self, channel: str, external_id: str) -> str:
        return f"{channel}:{external_id}"

    def history(self, channel: str, external_id: str) -> list[ChatMessage]:
        return list(self._store.get(self._key(channel, external_id), []))

    def append(self, channel: str, external_id: str, message: ChatMessage) -> None:
        key = self._key(channel, external_id)
        history = self._store.setdefault(key, [])
        history.append(message)
        if len(history) > self.max_turns * 2:
            self._store[key] = history[-self.max_turns * 2:]

    def reset(self, channel: str, external_id: str) -> None:
        self._store.pop(self._key(channel, external_id), None)
