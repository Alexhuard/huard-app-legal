from pathlib import Path
from src.core.knowledge_base import KnowledgeBase


def test_knowledge_base_loads_default_docs():
    kb = KnowledgeBase(Path("knowledge"))
    summary = kb.summary()
    assert summary["total_docs"] >= 3
    assert "general" in summary["by_category"]


def test_render_for_prompt_includes_general_first():
    kb = KnowledgeBase(Path("knowledge"))
    rendered = kb.render_for_prompt()
    assert "Catégorie : general" in rendered
    assert "ton_et_style" in rendered
