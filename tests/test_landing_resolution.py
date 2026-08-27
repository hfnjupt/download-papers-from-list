import importlib.util
import sys
import unittest
from pathlib import Path


SCRIPT = Path(__file__).parents[1] / "scripts" / "download_papers_from_list.py"
SPEC = importlib.util.spec_from_file_location("paper_downloader", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
sys.modules["paper_downloader"] = MODULE
SPEC.loader.exec_module(MODULE)


class LandingPageResolutionTests(unittest.TestCase):
    def test_cvf_paper_page_maps_to_pdf(self):
        page_url = (
            "https://openaccess.thecvf.com/content/CVPR2026/html/"
            "Xu_Bidirectional_Cross-Modal_Prompting_for_Event-Frame_Asymmetric_Stereo_CVPR_2026_paper.html"
        )
        expected = (
            "https://openaccess.thecvf.com/content/CVPR2026/papers/"
            "Xu_Bidirectional_Cross-Modal_Prompting_for_Event-Frame_Asymmetric_Stereo_CVPR_2026_paper.pdf"
        )
        self.assertEqual(MODULE.extract_pdf_link(b"<html></html>", page_url), expected)

    def test_pdf_anchor_beats_supplement(self):
        page = b"""
        <h2>Related Material</h2>
        <a href="../supplemental/paper_supp.pdf">[supp]</a>
        <a href="../papers/paper.pdf">[pdf]</a>
        <a href="https://arxiv.org/abs/1234.5678">[arXiv]</a>
        """
        result = MODULE.extract_pdf_link(page, "https://papers.example.org/html/paper.html")
        self.assertEqual(result, "https://papers.example.org/papers/paper.pdf")

    def test_citation_pdf_metadata(self):
        page = b'<meta name="citation_pdf_url" content="/downloads/paper-file">'
        result = MODULE.extract_pdf_link(page, "https://publisher.example.org/article/1")
        self.assertEqual(result, "https://publisher.example.org/downloads/paper-file")

    def test_download_button_data_url(self):
        page = b'<button aria-label="Download PDF" data-url="/article/1/download">Download</button>'
        result = MODULE.extract_pdf_link(page, "https://publisher.example.org/article/1")
        self.assertEqual(result, "https://publisher.example.org/article/1/download")

    def test_embedded_pdf_viewer(self):
        page = b'<iframe src="/viewer/article.pdf"></iframe>'
        result = MODULE.extract_pdf_link(page, "https://publisher.example.org/article/1")
        self.assertEqual(result, "https://publisher.example.org/viewer/article.pdf")


if __name__ == "__main__":
    unittest.main()
