from __future__ import annotations
from dataclasses import dataclass
from typing import List, Tuple
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table


@dataclass
class TextUnit:
	path: Tuple[int, ...]
	text: str
	type: str  # 'paragraph' | 'table_cell'


def extract_text_units(doc_path: str) -> Tuple[List[TextUnit], Document]:
	doc = Document(doc_path)
	units: List[TextUnit] = []

	for p_idx, p in enumerate(doc.paragraphs):
		text = p.text or ""
		if text.strip():
			units.append(TextUnit(path=(p_idx,), text=text, type="paragraph"))

	for t_idx, tbl in enumerate(doc.tables):
		for r_idx, row in enumerate(tbl.rows):
			for c_idx, cell in enumerate(row.cells):
				for p_idx, p in enumerate(cell.paragraphs):
					text = p.text or ""
					if text.strip():
						units.append(TextUnit(path=(t_idx, r_idx, c_idx, p_idx), text=text, type="table_cell"))

	return units, doc


def replace_text_in_document(doc: Document, original_to_translated: List[Tuple[TextUnit, str]]) -> Document:
	for unit, translated in original_to_translated:
		if unit.type == "paragraph" and len(unit.path) == 1:
			p_idx = unit.path[0]
			p: Paragraph = doc.paragraphs[p_idx]
			for _ in range(len(p.runs)):
				p.runs[0]._r.getparent().remove(p.runs[0]._r)
			p.add_run(translated)
		elif unit.type == "table_cell" and len(unit.path) == 4:
			t_idx, r_idx, c_idx, p_idx = unit.path
			tbl: Table = doc.tables[t_idx]
			cell: _Cell = tbl.rows[r_idx].cells[c_idx]
			p: Paragraph = cell.paragraphs[p_idx]
			for _ in range(len(p.runs)):
				p.runs[0]._r.getparent().remove(p.runs[0]._r)
			p.add_run(translated)
	return doc

