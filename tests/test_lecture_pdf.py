"""
Tolérance aux pannes de moteur/lecture_pdf.py : un PDF illisible (ou tout
autre problème sur UN devis) ne doit jamais interrompre le traitement des
autres — voir CLAUDE.md, "Tolérance aux pannes" (session v1.0).
"""

import shutil

from moteur.lecture_pdf import analyser_devis

from conftest import FIXTURES


def test_pdf_corrompu_n_interrompt_pas_les_autres(tmp_path, capsys):

    shutil.copy(FIXTURES / "ravate.pdf", tmp_path / "ravate.pdf")
    (tmp_path / "corrompu.pdf").write_bytes(b"ceci n'est pas un vrai PDF")

    articles = analyser_devis(tmp_path)

    # Le PDF valide est bien lu malgré le PDF corrompu à côté : jamais de
    # plantage global sur une anomalie locale.
    assert len(articles) > 0

    sortie = capsys.readouterr().out
    assert "corrompu.pdf" in sortie
    assert "PDF illisible" in sortie
    assert "2 PDF trouvé" in sortie
    assert "1 PDF lu(s) sur 2" in sortie
    assert "1 anomalie(s)" in sortie
