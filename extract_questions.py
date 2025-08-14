import json, re
from pathlib import Path

BASE = Path('/Users/nklakshminarasimhamurthy/Desktop/salesforce')
DOCX_PATH = BASE / 'salesforce set imppp.docx'
PDF1_PATH = BASE / 'AgentForce Questions - test - AgentForce Mcqs.pdf'
PDF2_PATH = BASE / 'Salesforce Certified Agentforce Specialist - Exam Questions.pdf'
OUT_PATH = BASE / 'questions.json'

questions = []
docx_rows = []
pdf1_rows_all = []
pdf2_rows_all = []
# Globals used by DOCX flush helper
q_text_accum = []
options = []
correct_idx = None

# --------------------------
# Parse DOCX (correct answers marked as (Selected))
# --------------------------
try:
	import docx  # python-docx
	doc = docx.Document(str(DOCX_PATH))
	lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]

	# Matches bullets like "A. text ... (Selected)"
	option_re = re.compile(r"^[\u2022•\-\s]*([A-E])\.\s*(.*?)(?:\s*\(Selected\))?\s*$")

	def flush_question():
		global q_text_accum, options, correct_idx
		if q_text_accum and options and (correct_idx is not None):
			question_text = ' '.join(q_text_accum).strip()
			item = {'question': question_text, 'options': options[:], 'correct': correct_idx, 'source': 'docx'}
			questions.append(item)
			docx_rows.append(item)
		q_text_accum, options, correct_idx = [], [], None

	for raw in lines:
		if re.match(r'^Question\s+\d+\s+of\s+\d+', raw, flags=re.I):
			flush_question()
			continue
		m = option_re.match(raw)
		if m:
			letter = m.group(1)
			text = m.group(2).strip()
			is_selected = '(Selected)' in raw
			options.append(text)
			if is_selected:
				correct_idx = ord(letter) - ord('A')
			continue
		q_text_accum.append(raw)
	flush_question()
	print(f'DOCX parsed: {len(questions)}')
except Exception as e:
	print('DOCX parse error:', e)

# --------------------------
# Helpers for PDF table parsing
# --------------------------

def normalize_cell(x):
	if x is None: return ''
	s = str(x).strip()
	# collapse internal whitespace
	return re.sub(r"\s+", ' ', s)


def parse_pdf_table_with_headers(pdf_path, header_alias_map):
	"""
	Try to find tables with headers and extract rows.
	header_alias_map: mapping logical name -> list of header aliases
	Returns list of dicts with keys from header_alias_map keys.
	"""
	import pdfplumber
	rows_out = []
	with pdfplumber.open(str(pdf_path)) as pdf:
		for page in pdf.pages:
			try:
				tables = page.extract_tables()
			except Exception:
				tables = []
			for t in tables or []:
				if not t or all(not any(c) for c in t):
					continue
				# Normalize
				t_norm = [[normalize_cell(c) for c in row] for row in t]
				# Find header row index by matching aliases
				header_idx = -1
				for idx, row in enumerate(t_norm[:5]):
					joined = ' | '.join(row).lower()
					matches = 0
					for key, aliases in header_alias_map.items():
						if any(a.lower() in joined for a in aliases):
							matches += 1
					if matches >= max(3, len(header_alias_map)//2):
						header_idx = idx
						break
				if header_idx == -1:
					continue
				headers = t_norm[header_idx]
				# Map column indices to logical keys
				col_map = {}
				for i, h in enumerate(headers):
					hl = h.lower()
					for key, aliases in header_alias_map.items():
						if any(a.lower() in hl for a in aliases):
							col_map.setdefault(key, i)
				# Consume data rows
				for row in t_norm[header_idx+1:]:
					if not any(row):
						continue
					obj = {}
					for key, col in col_map.items():
						obj[key] = row[col] if col < len(row) else ''
					rows_out.append(obj)
	return rows_out

# --------------------------
# Parse PDF1: expect headers like "Questions, optionA..optionE, selectedoption"
# First, generic tables pass (already), then an aggressive table-settings pass
# --------------------------
try:
	pdf1_rows = parse_pdf_table_with_headers(
		PDF1_PATH,
		{
			'question': ['questions','question'],
			'optA': ['option a','optiona','option a.','optiona.','a'],
			'optB': ['option b','optionb','b'],
			'optC': ['option c','optionc','c'],
			'optD': ['option d','optiond','d'],
			'optE': ['option e','optione','e'],
			'answer': ['selectedoption','selected option','answer']
		}
	)
	added = 0
	for r in pdf1_rows:
		qtxt = normalize_cell(r.get('question',''))
		if not qtxt:
			continue
		opts = [normalize_cell(r.get('optA','')), normalize_cell(r.get('optB','')), normalize_cell(r.get('optC','')), normalize_cell(r.get('optD','')), normalize_cell(r.get('optE',''))]
		opts = [o for o in opts if o]
		if len(opts) < 2:
			continue
		ans = normalize_cell(r.get('answer','')).strip()
		m = re.match(r'^([A-Ea-e])$', ans)
		if not m:
			# Sometimes answer cell may contain like 'optionA'
			m = re.search(r'option\s*([A-E])', ans, flags=re.I)
		if not m:
			continue
		letter = m.group(1).upper()
		correct_idx = ord(letter) - ord('A')
		if correct_idx >= len(opts):
			continue
		item = {'question': qtxt, 'options': opts, 'correct': correct_idx, 'source': 'pdf1'}
		questions.append(item)
		pdf1_rows_all.append(item)
		added += 1
	print(f'PDF1 tables parsed: +{added}, total {len(questions)}')
except Exception as e:
	print('PDF1 table parse error:', e)

# Aggressive table extraction with tuned settings
try:
    import pdfplumber
    added = 0
    with pdfplumber.open(str(PDF1_PATH)) as pdf:
        for page in pdf.pages:
            try:
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "min_words_vertical": 1,
                    "min_words_horizontal": 1,
                })
            except Exception:
                tables = []
            for t in tables or []:
                t_norm = [[normalize_cell(c) for c in row] for row in t]
                if not t_norm:
                    continue
                # Heuristic: row length >= 5 => might be [Q, A, B, C, D, E, Answer]
                headers = t_norm[0]
                joined = ' | '.join(headers).lower()
                is_header = ('question' in joined or 'questions' in joined) and ('option' in joined or 'selected' in joined)
                start_idx = 1 if is_header else 0
                for row in t_norm[start_idx:]:
                    cells = [c for c in row if c]
                    if len(cells) < 4:
                        continue
                    # last cell might be answer
                    ans_cell = row[-1]
                    m = re.search(r'(?:^|\b)option\s*([A-E])\b|^([A-E])$', normalize_cell(ans_cell), flags=re.I)
                    if not m:
                        continue
                    letter = (m.group(1) or m.group(2)).upper()
                    # question assumed in first non-empty, options next 3-5 cells before last
                    qtxt = normalize_cell(row[0])
                    # find options among middle cells
                    mid = [normalize_cell(c) for c in row[1:-1] if normalize_cell(c)]
                    if len(mid) < 2:
                        continue
                    opts = mid[:5]
                    correct_idx = ord(letter) - ord('A')
                    if correct_idx >= len(opts):
                        continue
                    item = {'question': qtxt, 'options': opts, 'correct': correct_idx, 'source': 'pdf1'}
                    questions.append(item)
                    pdf1_rows_all.append(item)
                    added += 1
    print(f'PDF1 extra tables parsed: +{added}, total {len(questions)}')
except Exception as e:
    print('PDF1 extra tables error:', e)

# Heuristic table parsing for PDF1 without relying on headers
try:
	import pdfplumber
	added = 0
	with pdfplumber.open(str(PDF1_PATH)) as pdf:
		for page in pdf.pages:
			try:
				tables = page.extract_tables()
			except Exception:
				tables = []
			for t in tables or []:
				if not t or len(t[0]) < 4:
					continue
				t_norm = [[normalize_cell(c) for c in row] for row in t]
				# Identify answer column as the one where most rows look like optionX or single letter
				cols = list(range(len(t_norm[0])))
				best_col = None; best_hits = -1
				for ci in cols:
					hits = 0; total = 0
					for r in t_norm[1:]:
						if ci >= len(r):
							continue
						val = r[ci]
						if not val:
							continue
						total += 1
						if re.fullmatch(r'[A-Ea-e]', val) or re.search(r'option\s*[A-E]', val, flags=re.I):
							hits += 1
					if hits > best_hits and hits >= max(2, total//3):
						best_hits = hits; best_col = ci
				if best_col is None:
					continue
				# Choose question column as the longest text average among non-answer columns
				non_ans = [ci for ci in cols if ci != best_col]
				avg_len = []
				for ci in non_ans:
					lens = [len(r[ci]) for r in t_norm[1:] if ci < len(r) and r[ci]]
					avg_len.append((sum(lens)/len(lens) if lens else 0, ci))
				avg_len.sort(reverse=True)
				if not avg_len:
					continue
				qcol = avg_len[0][1]
				opt_cols = [ci for ci in non_ans if ci != qcol]
				for r in t_norm[1:]:
					# build row
					if qcol >= len(r) or best_col >= len(r):
						continue
					qtxt = r[qcol]
					ans = r[best_col]
					if not qtxt or not ans:
						continue
					m = re.search(r'option\s*([A-E])', ans, flags=re.I) or re.fullmatch(r'([A-Ea-e])', ans)
					if not m:
						continue
					letter = m.group(1).upper()
					opts = []
					for ci in opt_cols:
						if ci < len(r) and r[ci]:
							opts.append(r[ci])
					opts = [o for o in opts if o][:5]
					if len(opts) < 2:
						continue
					correct_idx = ord(letter) - ord('A')
					if correct_idx >= len(opts):
						continue
					item = {'question': qtxt, 'options': opts, 'correct': correct_idx, 'source': 'pdf1-heur'}
					questions.append(item)
					pdf1_rows_all.append(item)
					added += 1
	print(f'PDF1 heuristic tables parsed: +{added}, total {len(questions)}')
except Exception as e:
	print('PDF1 heuristic error:', e)

# --------------------------
# Parse PDF2: expect headers like "Q#, Question, Option A.., Answer"
# First, generic tables (already), then aggressive table pass and fallback text pass
# --------------------------
try:
	pdf2_rows = parse_pdf_table_with_headers(
		PDF2_PATH,
		{
			'qnum': ['q#','q no','no.'],
			'question': ['question','questions'],
			'optA': ['option a','optiona','a'],
			'optB': ['option b','optionb','b'],
			'optC': ['option c','optionc','c'],
			'optD': ['option d','optiond','d'],
			'optE': ['option e','optione','e'],
			'answer': ['answer','correct','selected option']
		}
	)
	added = 0
	for r in pdf2_rows:
		qtxt = normalize_cell(r.get('question',''))
		if not qtxt:
			continue
		opts = [normalize_cell(r.get('optA','')), normalize_cell(r.get('optB','')), normalize_cell(r.get('optC','')), normalize_cell(r.get('optD','')), normalize_cell(r.get('optE',''))]
		opts = [o for o in opts if o]
		if len(opts) < 2:
			continue
		ans = normalize_cell(r.get('answer','')).strip()
		m = re.match(r'^([A-Ea-e])$', ans)
		if not m:
			m = re.search(r'option\s*([A-E])', ans, flags=re.I)
		if not m:
			continue
		letter = m.group(1).upper()
		correct_idx = ord(letter) - ord('A')
		if correct_idx >= len(opts):
			continue
		item = {'question': qtxt, 'options': opts, 'correct': correct_idx, 'source': 'pdf2'}
		questions.append(item)
		pdf2_rows_all.append(item)
		added += 1
	print(f'PDF2 tables parsed: +{added}, total {len(questions)}')
except Exception as e:
	print('PDF2 table parse error:', e)

# Aggressive table extraction for PDF2
try:
    import pdfplumber
    added = 0
    with pdfplumber.open(str(PDF2_PATH)) as pdf:
        for page in pdf.pages:
            try:
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "min_words_vertical": 1,
                    "min_words_horizontal": 1,
                })
            except Exception:
                tables = []
            for t in tables or []:
                t_norm = [[normalize_cell(c) for c in row] for row in t]
                if not t_norm:
                    continue
                headers = t_norm[0]
                joined = ' | '.join(headers).lower()
                is_header = ('question' in joined) and ('option' in joined) and ('answer' in joined)
                start_idx = 1 if is_header else 0
                for row in t_norm[start_idx:]:
                    cells = [c for c in row if c]
                    if len(cells) < 4:
                        continue
                    ans_cell = row[-1]
                    m = re.search(r'(?:^|\b)option\s*([A-E])\b|^([A-E])$', normalize_cell(ans_cell), flags=re.I)
                    if not m:
                        continue
                    letter = (m.group(1) or m.group(2)).upper()
                    # Attempt to identify question column: choose the longest cell among first half
                    first_half = row[:-1]
                    qtxt = max((normalize_cell(c) for c in first_half), key=len, default='')
                    # Options are remaining non-longest cells (limit 5)
                    opts = [normalize_cell(c) for c in first_half if normalize_cell(c) and normalize_cell(c) != qtxt][:5]
                    if len(opts) < 2:
                        continue
                    correct_idx = ord(letter) - ord('A')
                    if correct_idx >= len(opts):
                        continue
                    item = {'question': qtxt, 'options': opts, 'correct': correct_idx, 'source': 'pdf2'}
                    questions.append(item)
                    pdf2_rows_all.append(item)
                    added += 1
    print(f'PDF2 extra tables parsed: +{added}, total {len(questions)}')
except Exception as e:
    print('PDF2 extra tables error:', e)

# Fallback text parsing for PDF2 segments containing Question/Option/Answer tokens
try:
    import pdfplumber
    added = 0
    with pdfplumber.open(str(PDF2_PATH)) as pdf:
        text = []
        for page in pdf.pages:
            t = page.extract_text() or ''
            text.append(t)
    full = '\n'.join(text)
    # Split by occurrences of 'Question' followed by something and eventually 'Answer'
    blocks = re.split(r'(?=\bQuestion\b)', full, flags=re.I)
    for b in blocks:
        if not b or ('Option' not in b and 'Option A' not in b) or 'Answer' not in b:
            continue
        # Extract question as text before first Option A
        qm = re.split(r'\bOption\s*A\b', b, flags=re.I)
        if len(qm) < 2:
            continue
        qtxt = normalize_cell(qm[0])
        rest = qm[1]
        # Extract options A..E
        opts = []
        for letter in ['A','B','C','D','E']:
            m = re.search(rf'\bOption\s*{letter}\b\s*[:\-]?\s*(.*?)(?=\bOption\s*[A-E]\b|\bAnswer\b|$)', rest, flags=re.I|re.S)
            if m:
                opt = normalize_cell(m.group(1))
                if opt:
                    opts.append(opt)
        if len(opts) < 2:
            continue
        am = re.search(r'\bAnswer\b\s*[:\-]?\s*([A-E])', b, flags=re.I)
        if not am:
            continue
        letter = am.group(1).upper()
        correct_idx = ord(letter) - ord('A')
        if correct_idx >= len(opts):
            continue
        item = {'question': qtxt, 'options': opts, 'correct': correct_idx, 'source': 'pdf2'}
        questions.append(item)
        pdf2_rows_all.append(item)
        added += 1
    print(f'PDF2 text parsed: +{added}, total {len(questions)}')
except Exception as e:
    print('PDF2 text parse error:', e)

# --------------------------
# Deduplicate by normalized question text
# --------------------------
final = []
seen_q = set()
for q in questions:
    qt = re.sub(r'\s+', ' ', q['question']).strip()
    if not qt or not (2 <= len(q.get('options', [])) <= 6):
        continue
    # produce unique (by question text) snapshot
    if qt not in seen_q:
        uq = dict(q)
        uq['question'] = qt
        final.append(uq)
        seen_q.add(qt)

# Also write full (no dedup) with source tagging
OUT_FULL = BASE / 'questions_full.json'
with open(OUT_PATH, 'w', encoding='utf-8') as f:
    json.dump({'questions': final}, f, ensure_ascii=False, indent=2)
with open(OUT_FULL, 'w', encoding='utf-8') as f:
    json.dump({
        'counts': {
            'docx': len(docx_rows),
            'pdf1': len(pdf1_rows_all),
            'pdf2': len(pdf2_rows_all),
            'combined_no_dedup': len(questions),
            'unique_by_text': len(final),
        },
        'questions': questions
    }, f, ensure_ascii=False, indent=2)

print(f'Wrote {len(final)} unique and {len(questions)} total (with duplicates) to questions.json and questions_full.json')
