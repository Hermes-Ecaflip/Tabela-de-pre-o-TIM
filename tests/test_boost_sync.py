import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import processar_planilha as app


class BoostSyncTests(unittest.TestCase):
    def test_boost_removido_da_planilha_nao_permanece_no_mapa(self):
        com_boost = pd.DataFrame({
            app.COL_CODIGO: [4010001],
            app.COL_BOOST: [500],
        })
        sem_boost = pd.DataFrame({
            app.COL_CODIGO: [4010001],
            app.COL_BOOST: [None],
        })

        with patch.object(app.pd, "read_excel", return_value=com_boost):
            self.assertEqual(app.carregar_boost(), {"4010001": 500.0})

        with patch.object(app.pd, "read_excel", return_value=sem_boost):
            self.assertEqual(app.carregar_boost(), {})

    def test_versao_do_script_e_trocada_para_evitar_cache_antigo(self):
        html = (
            '<link rel="stylesheet" href="style.css?v=estilo-antigo">'
            '<p class="header-badge" aria-label="Total de itens">1 itens</p>'
            '<time datetime="2026-09-01">01/09/2026</time>'
            '<script src="script.js?v=versao-antiga"></script>'
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            index_path = Path(temp_dir) / "index.html"
            index_path.write_text(html, encoding="utf-8")
            diretorio_anterior = os.getcwd()
            try:
                os.chdir(temp_dir)
                (Path(temp_dir) / "style.css").write_text("body{}", encoding="utf-8")
                app.atualizar_index_html("02/09/2026", 1, "novo-script", "novo-estilo")
            finally:
                os.chdir(diretorio_anterior)

            atualizado = index_path.read_text(encoding="utf-8")

        self.assertIn('src="script.js?v=novo-script"', atualizado)
        self.assertIn('href="style.css?v=novo-estilo"', atualizado)
        self.assertNotIn("versao-antiga", atualizado)

    def test_regeneracao_remove_boost_antigo_do_script(self):
        script_antigo = (
            'const DATA = [{"codigo":"4010001","boost":500}];\n'
            'const OUTRA_CONFIGURACAO = true;\n'
        )
        dados_novos = '[{"codigo":"4010001","boost":null}]'

        with tempfile.TemporaryDirectory() as temp_dir:
            script_path = Path(temp_dir) / "script.js"
            script_path.write_text(script_antigo, encoding="utf-8")
            diretorio_anterior = os.getcwd()
            try:
                os.chdir(temp_dir)
                app.atualizar_script_js(dados_novos)
            finally:
                os.chdir(diretorio_anterior)

            atualizado = script_path.read_text(encoding="utf-8")

        self.assertIn('"boost":null', atualizado)
        self.assertNotIn('"boost":500', atualizado)
        self.assertIn("const OUTRA_CONFIGURACAO = true;", atualizado)


if __name__ == "__main__":
    unittest.main()
