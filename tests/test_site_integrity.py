import re
import unittest
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class IdCollector(HTMLParser):
    def __init__(self):
        super().__init__()
        self.ids = []

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if attributes.get("id"):
            self.ids.append(attributes["id"])


class SiteIntegrityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.html = (ROOT / "index.html").read_text(encoding="utf-8")
        cls.script = (ROOT / "script.js").read_text(encoding="utf-8")
        cls.style = (ROOT / "style.css").read_text(encoding="utf-8")

    def test_html_nao_tem_ids_duplicados(self):
        parser = IdCollector()
        parser.feed(self.html)
        duplicates = [item for item, count in Counter(parser.ids).items() if count > 1]
        self.assertEqual(duplicates, [])

    def test_ids_usados_no_javascript_existem_no_html(self):
        parser = IdCollector()
        parser.feed(self.html)
        html_ids = set(parser.ids)
        referenced_ids = set(re.findall(r"getElementById\('([^']+)'\)", self.script))
        self.assertEqual(sorted(referenced_ids - html_ids), [])

    def test_arquivos_estaticos_possuem_versao_de_cache(self):
        self.assertRegex(self.html, r'href="style\.css\?v=[0-9a-f]{12}"')
        self.assertRegex(self.html, r'src="script\.js\?v=[0-9a-f]{12}"')

    def test_css_esta_estruturalmente_balanceado(self):
        self.assertEqual(self.style.count("{"), self.style.count("}"))
        self.assertIn("@media (max-width:1080px)", self.style)
        self.assertIn("@media (max-width:760px)", self.style)
        self.assertIn("@media (max-width:480px)", self.style)


if __name__ == "__main__":
    unittest.main()
