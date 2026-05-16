from pathlib import Path
from dataclasses import dataclass


@dataclass
class KnowledgeDoc:
    relative_path: str
    category: str
    content: str


class KnowledgeBase:
    """Charge tous les fichiers markdown du dossier knowledge/ et les assemble
    en un seul bloc cacheable côté Claude (prompt caching).

    Catégories attendues :
      - general/      : ton, style, règles transverses
      - procedures/   : procédures internes HUARD
      - faq/          : questions/réponses récurrentes
      - templates/    : signatures, formules type
    """

    def __init__(self, root: Path):
        self.root = root
        self.docs: list[KnowledgeDoc] = []
        self._load()

    def _load(self) -> None:
        if not self.root.exists():
            return
        for path in sorted(self.root.rglob("*.md")):
            if path.name.startswith("_"):
                continue
            try:
                content = path.read_text(encoding="utf-8").strip()
            except OSError:
                continue
            if not content:
                continue
            rel = path.relative_to(self.root)
            category = rel.parts[0] if len(rel.parts) > 1 else "general"
            self.docs.append(
                KnowledgeDoc(
                    relative_path=str(rel),
                    category=category,
                    content=content,
                )
            )

    def render_for_prompt(self) -> str:
        """Rend la base sous forme d'un seul bloc texte structuré, prêt à
        être injecté dans le system prompt avec cache_control."""
        if not self.docs:
            return "(Base de connaissances vide pour l'instant.)"

        by_category: dict[str, list[KnowledgeDoc]] = {}
        for doc in self.docs:
            by_category.setdefault(doc.category, []).append(doc)

        order = ["general", "procedures", "faq", "templates"]
        sections: list[str] = []
        for cat in order + [c for c in by_category if c not in order]:
            docs = by_category.get(cat)
            if not docs:
                continue
            sections.append(f"## Catégorie : {cat}\n")
            for doc in docs:
                sections.append(f"### {doc.relative_path}\n{doc.content}\n")
        return "\n".join(sections)

    def summary(self) -> dict:
        return {
            "total_docs": len(self.docs),
            "by_category": {
                cat: sum(1 for d in self.docs if d.category == cat)
                for cat in {d.category for d in self.docs}
            },
        }
