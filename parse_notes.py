#!/usr/bin/env python3
"""
parse_notes.py  —  Biotech weekly-notes -> normalized JSON (and dashboard HTML)

Reads one or more Word (.docx) files, each a week's biotech news summary written
in the author's nested-bullet style, and produces a normalized data model:

    week -> category (disease | modality | theme)
         -> area   (e.g. Cancer, Cell therapy, Pricing)
         -> group  (e.g. GI, In vivo CAR-T, Ovarian)   [may nest]
         -> story  (a single news item)
         -> details (the supporting sub-bullets)

Usage:
    python parse_notes.py Oct_2025_W2.docx Oct_2025_W3.docx ... --json data.json
    python parse_notes.py *.docx --html-template template.html --html-out dashboard.html

Adding future weeks: just include the new .docx files on the command line and
re-run. Filenames like  Mon_YYYY_WN.docx  (e.g. Oct_2025_W4.docx) are parsed for
the week label/date automatically; anything else falls back to the file name.
"""
import sys, re, json, subprocess, argparse, os, calendar, html as _html

# --------------------------------------------------------------------------
# Known section names.  These make parsing robust to the (inconsistent) way
# areas are sometimes bullets, sometimes plain paragraphs, sometimes under an
# explicit bucket header and sometimes not.  Extend these lists as new areas
# appear in future weeks -- unknown top-level plain-text headings are still
# captured as areas via a fallback, but listing them here fixes their category.
# --------------------------------------------------------------------------
AREA_CATEGORY = {}
def _reg(cat, names):
    for n in names:
        AREA_CATEGORY[n.lower()] = cat

_reg('modality', [
    'cell therapy', 'gene therapy', 'gene editing', 'drug delivery', 'techbio',
    'platforms', 'rna', 'radiopharma', 'radiopharmaceuticals', 'protein degradation',
    'ai discovery', 'ai drug discovery', 'ai', 'devices', 'antibody', 'antibodies',
    'degraders', 'omics', 'peptides',
])
_reg('disease', [
    'cancer', 'oncology', 'cardio', 'cardiovascular', 'neuro', 'neurology',
    'immuno', 'immunology', 'metabolic', 'metabolism', 'infectious', 'eye',
    'ophthalmology', 'respiratory', 'lung', 'psych', 'psychiatry', 'pain', 'muscle',
    'kidney', 'renal', 'derm', 'dermatology', 'hearing', "women's health",
    "women\u2019s health", 'skin', 'gastro', 'hepatology', 'liver', 'rare disease',
    'bone', 'blood', 'hematology', 'aging',
])
_reg('theme', [
    'big pharma', 'china', 'commercialization', 'eu', 'markets', 'preclinical',
    'pricing', 'regulatory', 'restructuring', 'funds', 'healthcare', 'trends',
    'policy', 'ipo', 'm&a', 'payors', 'payers', 'healthcare delivery',
    'consumer health', 'research funding', 'clinical trials', 'distribution',
    'prv', 'policy & regulatory',
])

# Pure "bucket" markers: they set the current category but are NOT areas.
BUCKET_MARKERS = {'platforms': 'modality', 'diseases': 'disease', 'trends': 'theme'}

# Regex signalling a bullet is a *story* (news item) rather than a short grouping label.
STORY_MARKERS = re.compile(
    r'\$|\d+\s?%|\bPh\s?[1-3]\b|\bPhase\b|\bFDA\b|\bEMA\b|\bCHMP\b|\bSeries\s+[A-E]\b|'
    r'acqui|approv|\bIPO\b|\bPIPE\b|biobuck|milestone|\bIND\b|\bCRL\b|\bPDUFA\b|\bBTD\b|'
    r'\bBLA\b|licens|raise|\bFPO\b|\bATM\b|\bORR\b|\bPFS\b|\bOS\b|terminat|partnership',
    re.I,
)

MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
MONTHS.update({m.lower(): i for i, m in enumerate(calendar.month_name) if m})

# ==========================================================================
# DISEASE CANONICALIZATION  (therapeutic-area sections only)
# --------------------------------------------------------------------------
# Every therapeutic area is rendered strictly as  TA -> disease -> stories,
# and the SAME disease keeps the SAME name across weeks so you can track it over
# time. Because the notes group stories inconsistently (sometimes under an organ
# system like "GI", sometimes flat with no grouping), the canonical disease is
# INFERRED FROM EACH STORY'S TEXT, not from where it happened to sit.
#
# THREE EDITABLE TABLES:
#   * TA_CANON         - merge therapeutic-area name variants / fix typos.
#   * DISEASE_PATTERNS - per-TA list of (canonical disease, identifying terms);
#                        order matters (specific first, first match wins).
#   * DISEASE_ALIASES  - tidy fallback labels taken from the author's own headers.
# Stories that match nothing land under "Misc." for that area.
# ==========================================================================
TA_CANON = {
    'oncology': 'Cancer', 'cancer': 'Cancer',
    'cardiovascular': 'Cardio', 'cardio': 'Cardio',
    'immunology': 'Immuno', 'immuno': 'Immuno', 'inflammation': 'Immuno',
    'neurology': 'Neuro', 'neuro': 'Neuro', 'neuroscience': 'Neuro',
    'metabolism': 'Metabolic', 'metabolic': 'Metabolic',
    'infectious': 'Infectious', 'infectious disease': 'Infectious',
    'ophthalmology': 'Eye', 'eye': 'Eye',
    'respiratory': 'Respiratory', 'respiraotry': 'Respiratory', 'pulmonary': 'Respiratory',
    'psychiatry': 'Psych', 'psych': 'Psych',
    'dermatology': 'Derm', 'derm': 'Derm',
    'renal': 'Kidney', 'kidney': 'Kidney', 'nephrology': 'Kidney',
    'hearing': 'Hearing', "women's health": "Women's health",
    "women\u2019s health": "Women's health",
    'muscle': 'Muscle', 'pain': 'Pain', 'bone': 'Bone', 'aging': 'Aging',
    'rare disease': 'Rare disease', 'hepatology': 'Liver', 'liver': 'Liver',
    'hematology': 'Blood', 'blood': 'Blood', 'gastro': 'GI',
}
def canon_ta(name):
    return TA_CANON.get(name.lower().strip().rstrip(':'), name.strip())

# Organ-system / catch-all headers that are NOT diseases: bypassed so the
# specific disease (inferred from the story) is used. If a story under one of
# these can't be resolved, it goes to "Misc." rather than to the organ label.
CONTAINERS = {
    'gi', 'gu', 'genitourinary', 'gastrointestinal', 'gyn', 'gynecologic',
    'gynaecologic', 'chest', 'thoracic', 'hematologic', 'haematologic', 'heme',
    'blood', 'solid tumors', 'solid tumor', 'solid tumours', 'solid tumour',
    'other', 'others', 'misc', 'miscellaneous', 'platforms', 'platform', 'general',
}

# Headers recognized as grouping levels (so a bare organ header with children is
# treated as structure, not as a story). Extend if new grouping headers appear.
KNOWN_GROUP_LABELS = {
    'gi', 'gu', 'gyn', 'gynecologic', 'hematologic', 'blood', 'melanoma',
    'complications', 'diagnostics', 'chest', 'thoracic', 'supportive care',
    'other', 'platforms', 'bladder', 'prostate', 'breast', 'endometrial',
    'ovarian', 'nsclc', 'sarcoma', 'allergy', 'arthritis', 'gut', 'kidney',
    'neuro', 'skin', 'respiratory', 'atherosclerosis', 'blood pressure',
    'cardiomyopathy', 'hf', 'thrombosis', 'diabetes', 'obesity', 't2d', 'mash',
    'enzyme deficiency', 'triglycerides', 'antibiotic', 'antiviral', 'vaccines',
    'non-therapeutics',
}

DISEASE_ALIASES = {
    'hf': 'Heart failure', 'blood pressure': 'Hypertension',
    'triglycerides': 'Dyslipidemia', 'atherosclerosis': 'Atherosclerosis / lipids',
    'gut': 'IBD', 'skin': 'Skin (immune)', 'neuro': 'Neuro-immune',
    'kidney': 'Kidney (immune)', 'arthritis': 'Arthritis', 'allergy': 'Allergy',
    'complications': 'Supportive care', 'respiratory': 'Respiratory (immune)',
    't2d': 'Diabetes', 'mash': 'MASH/NASH', 'antibiotic': 'Bacterial infection',
    'antiviral': 'Viral infection', 'vaccines': 'Vaccines',
    'non-therapeutics': 'Policy & surveillance', 'enzyme deficiency': 'Enzyme deficiency',
}
def canon_disease_label(lab):
    low = lab.lower().strip().rstrip(':')
    return DISEASE_ALIASES.get(low, lab.strip())

# Per-TA disease patterns. Compiled case-SENSITIVELY: use (?i:word) for
# case-insensitive words and bare UPPERCASE for acronyms, so e.g. the acronym
# "MS" doesn't match the word "mass".
_DP_RAW = {
 'Cancer': [
    ('NSCLC', r'NSCLC|(?i:non.?small.?cell)'),
    ('SCLC', r'SCLC|(?i:small.?cell lung)'),
    ('Mesothelioma', r'(?i:mesothelioma)'),
    ('Head & neck cancer', r'HNSCC|(?i:head and neck)'),
    ('Breast cancer', r'TNBC|(?i:breast)'),
    ('Ovarian cancer', r'PROC|(?i:ovarian|platinum.?resistant|folate receptor)|FR[a\u03b1]'),
    ('Endometrial cancer', r'(?i:endometrial|uterine)'),
    ('Cervical cancer', r'(?i:cervical)'),
    ('Prostate cancer', r'mCRPC|mHSPC|PSMA|(?i:prostate)'),
    ('Bladder cancer', r'MIBC|NMIBC|(?i:bladder|urothelial)'),
    ('Renal cell carcinoma', r'ccRCC|RCC|(?i:renal cell)'),
    ('Colorectal cancer', r'CRC|(?i:colorectal)'),
    ('Gastric/GEJ cancer', r'GEJ|(?i:gastric|gastro.?esoph)'),
    ('Neuroendocrine tumors', r'GEP.?NET|NETs?\b|(?i:neuroendocrine)'),
    ('Pancreatic cancer', r'PDAC|(?i:pancrea)'),
    ('Liver cancer (HCC)', r'HCC|(?i:hepatocellular)'),
    ('Uveal melanoma', r'(?i:uveal)'),
    ('Melanoma', r'(?i:melanoma)'),
    ('Multiple myeloma', r'RRMM|MM\b|(?i:myeloma)'),
    ('Myelofibrosis', r'\bMF\b|(?i:myelofibrosis)'),
    ('Polycythemia vera', r'\bPV\b|(?i:polycythemia)'),
    ('AML', r'AML|(?i:acute myeloid)'),
    ('CML', r'CML|(?i:chronic myeloid)'),
    ('Lymphoma', r'DLBCL|LBCL|NHL|CTCL|(?i:lymphoma|s\u00e9zary|sezary)'),
    ('Leukemia (other)', r'CLL|MDS|(?i:leukemia|leukaemia|waldenstr)'),
    ('Skin cancer (CSCC/BCC)', r'CSCC|BCC|(?i:squamous cell carcinoma|basal cell)'),
    ('Sarcoma', r'(?i:sarcoma)'),
    ('Cancer cachexia', r'(?i:cachexia)'),
    ('Multi-cancer detection', r'MCED|(?i:multi.?cancer|early detection|surveillance)'),
    ('Solid tumors (various)', r'(?i:solid tumou?rs?)'),
 ],
 'Cardio': [
    ('Atherosclerosis / lipids', r'LDL|CETP|PCSK9|Lp\(a\)|ASCVD|(?i:atherosclerosis|cholesterol|bempedoic|lipoprotein|statin)'),
    ('Stroke', r'(?i:stroke)'),
    ('Hypertension', r'(?i:hypertension|blood pressure|aldosterone)'),
    ('Heart failure', r'HFpEF|HFrEF|\bHF\b|(?i:heart failure)'),
    ('Cardiomyopathy', r'ATTR|DCM|HCM|BAG3|(?i:cardiomyopathy|myosin)'),
    ('Thrombosis', r'FXIa|(?i:thrombosis|thrombectomy|embolism|anticoag)'),
    ('Valve disease', r'(?i:tricuspid|mitral|aortic|valve)'),
    ('Angina / ischemia', r'(?i:angina|ischemi)'),
 ],
 'Metabolic': [
    ('Obesity', r'(?i:obesity|weight loss|weight-loss|overweight)'),
    ('Diabetes', r'T2D|T1D|HbA1c|(?i:diabet|\binsulin)'),
    ('MASH/NASH', r'MASH|NASH|FGF21|(?i:steatohepat)'),
    ('Dyslipidemia', r'APOC3|FCS|(?i:triglyceride|dyslipidemia)'),
    ('Prader-Willi', r'PWS|(?i:prader|hyperphagia)'),
    ('Pompe disease', r'(?i:pompe)'),
    ('Fabry disease', r'(?i:fabry)'),
    ('Primary hyperoxaluria', r'PH1|HAO1|(?i:hyperoxaluria)'),
    ('Wilson disease', r'(?i:wilson)'),
    ('Lysosomal / MPS', r'NAGLU|MPS|GSDIa|(?i:sanfilippo|mucopolysacchar|glycogen storage)'),
    ('ENPP1 deficiency', r'ENPP1'),
    ('Urea cycle disorders', r'UCD|(?i:urea cycle)'),
 ],
 'Immuno': [
    ('Atopic dermatitis', r'(?i:atopic|eczema)'),
    ('Hidradenitis (HS)', r'\bHS\b|(?i:hidradenitis)'),
    ('Prurigo nodularis', r'(?i:prurigo|pruritus)'),
    ('Asthma', r'(?i:asthma)'),
    ('COPD', r'COPD'),
    ('Bronchiectasis', r'(?i:bronchiectasis)'),
    ('CRSwNP (nasal polyps)', r'CRSwNP|(?i:nasal polyp)'),
    ('Chronic urticaria (CSU)', r'CSU|(?i:urticaria)'),
    ('Food allergy', r'(?i:peanut|food allergy)'),
    ('Gout', r'URAT1|(?i:gout)'),
    ('Rheumatoid arthritis', r'\bRA\b|(?i:rheumatoid)'),
    ('Myositis / IBM', r'\bIBM\b|IMNM|(?i:myositis|inclusion body|dermatomyositis)'),
    ('Vasculitis (ANCA)', r'ANCA|\bAAV\b|(?i:vasculitis)'),
    ('Thyroid eye disease', r'\bTED\b|(?i:thyroid eye)'),
    ('Lupus / lupus nephritis', r'SLE|\bLN\b|(?i:lupus)'),
    ('IgA nephropathy', r'IgAN|(?i:iga nephropathy)'),
    ('Multiple sclerosis', r'\bMS\b|(?i:multiple sclerosis)'),
    ('Psoriasis', r'PsO|(?i:psoriasis|psoriatic)'),
    ('IBD / Crohn\u2019s', r'IBD|(?i:crohn|ulcerative colitis|inflammatory bowel)'),
    ('Scleroderma / SSc', r'SSc|(?i:scleroderma|systemic sclerosis)'),
    ('Pemphigus', r'(?i:pemphigus)'),
    ('Hereditary angioedema', r'HAE'),
    ('PNH', r'PNH'),
    ('GVHD', r'GVHD|(?i:graft.?versus.?host)'),
    ('ITP', r'ITP'),
    ('HLH', r'HLH|(?i:hemophagocytic)'),
    ('Allergy (other)', r'(?i:allergic|allergy)'),
 ],
 'Neuro': [
    ("Parkinson's", r'LRRK2|(?i:parkinson|alpha.?synuclein)'),
    ('Essential tremor', r'(?i:essential tremor|tremor)'),
    ("Alzheimer's", r'pTau|(?i:alzheimer|amyloid|\btau\b)'),
    ('Lewy body dementia', r'(?i:lewy body)'),
    ('Frontotemporal dementia', r'FTD|(?i:frontotemporal)'),
    ('ALS', r'ALS|(?i:motor neuron|amyotrophic)'),
    ('Huntington', r'(?i:huntington)'),
    ('Tardive dyskinesia', r'(?i:tardive|dyskinesia)'),
    ('Epilepsy', r'(?i:epilep|seizure|dravet|lennox)'),
    ('Narcolepsy', r'(?i:narcolepsy|cataplexy)'),
    ('Gaucher disease', r'GBA1|(?i:gaucher)'),
    ('Spinocerebellar ataxia', r'SCA\b|(?i:ataxia)'),
    ('MPS II (Hunter)', r'MPS II|(?i:hunter syndrome)'),
    ('Rett syndrome', r'MECP2|(?i:rett)'),
    ('Multiple system atrophy', r'MSA'),
 ],
 'Infectious': [
    ('HIV', r'HIV|PrEP'),
    ('RSV', r'RSV'),
    ('Influenza', r'(?i:influenza|\bflu\b)'),
    ('COVID-19', r'COVID|SARS-CoV'),
    ('Ebola', r'(?i:ebola|bundibugyo)'),
    ('Measles', r'MMR|(?i:measles)'),
    ('Hepatitis', r'HBV|HCV|HDV|(?i:hepatitis)'),
    ('Tuberculosis', r'\bTB\b|(?i:tuberculosis)'),
    ('Prosthetic joint infection', r'(?i:prosthetic joint|periprosthetic|2-stage exchange)'),
    ('C. difficile', r'(?i:difficile|clostridi)'),
 ],
 'Eye': [
    ('Geographic atrophy / AMD', r'\bGA\b|AMD|(?i:geographic atrophy|macular degeneration|wet AMD)'),
    ('Diabetic eye disease', r'DME|(?i:macular edema|diabetic retinopath)'),
    ('Uveitis', r'(?i:uveitis|non-infectious)'),
    ('Stargardt disease', r'(?i:stargardt)'),
    ('Glaucoma', r'(?i:glaucoma)'),
    ('Dry eye', r'(?i:dry eye)'),
 ],
 'Muscle': [
    ('Duchenne (DMD)', r'DMD|(?i:duchenne)'),
    ('FSHD', r'FSHD|DUX4'),
    ('Spinal muscular atrophy', r'\bSMA\b|(?i:spinal muscular)'),
    ('Myasthenia gravis', r'(?i:myasthenia)'),
    ('Myotonic dystrophy', r'DM1|(?i:myotonic)'),
 ],
 'Respiratory': [
    ('Asthma', r'(?i:asthma)'),
    ('COPD', r'COPD|AATD'),
    ('Cystic fibrosis', r'CFTR|\bCF\b|(?i:cystic fibrosis)'),
    ('Sarcoidosis', r'(?i:sarcoidosis)'),
    ('IPF / fibrosis', r'IPF|(?i:pulmonary fibrosis|fibrosis)'),
 ],
 'Psych': [
    ('Schizophrenia', r'PANSS|(?i:schizophren)'),
    ('Depression / TRD', r'\bTRD\b|\bMDD\b|(?i:depression)'),
    ('Anxiety / GAD', r'\bGAD\b|\bSAD\b|(?i:anxiety|social anxiety)'),
    ('Bipolar', r'(?i:bipolar)'),
    ('ADHD', r'ADHD'),
    ('Substance use', r'\bOUD\b|(?i:opioid use|addiction|psilocybin)'),
 ],
 'Pain': [
    ('Migraine', r'CGRP|(?i:migraine)'),
    ('Acute / post-surgical pain', r'Nav1\.8|(?i:post-surgical|acute pain)'),
    ('CRPS', r'CRPS'),
 ],
 'Derm': [
    ('Vitiligo', r'(?i:vitiligo)'),
    ('Alopecia', r'(?i:alopecia)'),
    ('Hidradenitis', r'\bHS\b|(?i:hidradenitis)'),
    ('Prurigo nodularis', r'(?i:prurigo)'),
 ],
 'Bone': [
    ('Osteoarthritis', r'\bOA\b|(?i:osteoarthritis)'),
    ('FOP', r'FOP|(?i:fibrodysplasia)'),
    ('Osteoporosis', r'(?i:osteoporosis)'),
 ],
 'Blood': [
    ('Hemophilia / bleeding', r'vWD|(?i:hemophilia|coagulation|glanzmann|bleeding)'),
    ('HHT', r'HHT|(?i:telangiectasia)'),
    ('Sickle cell', r'\bSCD\b|(?i:sickle cell)'),
    ('Prion disease', r'(?i:prion)'),
 ],
 "Women's health": [
    ('Menopause / VMS', r'VMS|(?i:menopaus|vasomotor)'),
    ('Endometriosis', r'(?i:endometriosis)'),
    ('Contraception', r'(?i:contracepti)'),
 ],
}
DISEASE_PATTERNS = {ta: [(name, re.compile(p)) for name, p in rows]
                    for ta, rows in _DP_RAW.items()}

def infer_disease(ta, text):
    for name, rx in DISEASE_PATTERNS.get(ta, []):
        if rx.search(text):
            return name
    return None

def is_group_label(label, nkids):
    """True if this header is a grouping level (organ/disease bucket) rather than
    a story headline."""
    low = label.lower().strip().rstrip(':')
    if STORY_MARKERS.search(label):
        return False
    if low in CONTAINERS or low in KNOWN_GROUP_LABELS:
        return True
    if word_count(low) <= 2 and nkids >= 2 and not re.search(r'[\d$:]', label):
        return True
    return False

def resolve_disease(ta, title, detail_text, chain):
    """Canonical disease for one story: infer from title, then details, then fall
    back to the nearest non-container grouping header, else 'Misc.'."""
    d = infer_disease(ta, title)
    if d: return d
    d = infer_disease(ta, detail_text)
    if d: return d
    for lab in reversed(chain):
        if lab.lower().strip().rstrip(':') in CONTAINERS:
            continue
        return canon_disease_label(lab)
    return 'Misc.'

# --------------------------------------------------------------------------
# Low-level text helpers
# --------------------------------------------------------------------------
def clean_inline(text):
    """Unescape pandoc escapes; keep markdown links as (text,url) untouched for later."""
    # pandoc escapes: \$ \> \< \' \" \~ \[ \] etc.
    text = re.sub(r'\\([\\`*_{}\[\]()#+\-.!$><~\'"^&|%@=;:])', r'\1', text)
    return text.strip()

def strip_emphasis(text):
    """Remove **bold**, *italic*, and [x]{.class} wrappers (.underline, .mark, ...)
    to get bare label text."""
    t = re.sub(r'\[([^\]]*)\]\{\.[a-zA-Z][\w-]*\}', r'\1', text)
    t = re.sub(r'\*\*([^*]*)\*\*', r'\1', t)
    t = re.sub(r'\*([^*]*)\*', r'\1', t)
    return t.strip()

def preprocess(text):
    """The author uses '*' or a lone '\\' as a multiplication sign in market-size
    math (pandoc-escaped). Normalize both to '×' so they aren't mistaken for
    markdown emphasis and read cleanly."""
    text = text.replace('\\*', '\u00d7')
    text = re.sub(r'(?<=\s)\\(?=\s)', '\u00d7', text)  # " \ " between numbers -> ×
    return text

def label_of(text):
    """The plain label used for classification / display headers (links stripped to text)."""
    t = preprocess(text)
    t = re.sub(r'\[([^\]]*)\]\([^)]*\)', r'\1', t)   # [t](u) -> t
    t = strip_emphasis(t)
    return clean_inline(t)

def is_underlined_header(raw):
    return bool(re.search(r'\{\.underline\}', raw)) or (raw.startswith('**') and raw.endswith('**'))

def is_italic_subheader(raw):
    """Italic-underline header used as a WITHIN-area sub-group divider, e.g.
    `*[Cancer GI]{.underline}*` (single asterisks, not the bold `**...**` bucket
    header). In newer notes these appear as bare paragraphs between the bucket
    header and the story bullets."""
    r = raw.strip()
    if r.startswith('**'):
        return False
    return bool(re.match(r'\*\[[^\]]*\]\{\.underline\}\*\s*$', r))

def to_html_inline(text):
    """Escape HTML then re-insert markdown links as anchors."""
    text = preprocess(text)
    # protect links
    links = []
    def _stash(m):
        links.append((m.group(1), m.group(2)))
        return f'\x00{len(links)-1}\x00'
    text = re.sub(r'\[([^\]]*)\]\(([^)]*)\)', _stash, text)
    text = clean_inline(text)
    text = strip_emphasis(text)
    text = _html.escape(text)
    def _restore(m):
        i = int(m.group(1))
        t, u = links[i]
        t = _html.escape(clean_inline(t)); u = _html.escape(u, quote=True)
        return f'<a href="{u}" target="_blank" rel="noopener">{t}</a>'
    text = re.sub(r'\x00(\d+)\x00', _restore, text)
    return text

def word_count(s):
    return len(s.split())

def looks_like_group(label, has_children):
    if not has_children:
        return False
    if STORY_MARKERS.search(label):
        return False
    lab = label.rstrip(':').strip()
    return word_count(lab) <= 5 and len(lab) <= 46

# --------------------------------------------------------------------------
# Badge extraction (small chips shown on each story)
# --------------------------------------------------------------------------
# Market size -> (value in $millions, display label like "$4.4B"), or (None, None).
# Used to show a green market-size badge at the end of big-market headlines.
_UNIT_MUSD = {'M': 1, 'B': 1000, 'T': 1000000}
def market_size(blob):
    m = re.search(r'=\s*\$?\s?(\d[\d.,]*)\s?([MBT])\s*TAM', blob)
    if not m:
        m = re.search(r'(?i:market size)[^=]*=\s*\u00d7?\s*\$?\s?(\d[\d.,]*)\s?([MBT])\b', blob)
    if not m:
        return (None, None)
    numraw = m.group(1).rstrip('.,')
    unit = m.group(2).upper()
    try:
        val = float(numraw.replace(',', ''))
    except ValueError:
        return (None, None)
    return (val * _UNIT_MUSD[unit], '$' + numraw + unit)

# Negative-news detection (colors the headline red). Edit to taste.
NEGATIVE_RX = re.compile(
    r'(?i:'
    r'\bfail(?:s|ed|ure|ing)?\b(?!\s+(?:detection|mode|analysis))|discontinu(?:e|es|ed|ing)\b(?!\s*\d{4})|terminat(?:e|es|ed|ion)|'
    r'\bCRL\b(?!\s+(?:collaborat|partner))|complete response letter|clinical hold|\bhalt(?:s|ed|ing)?\b|'
    r'misse[ds]?\s+(?:the\s+|its\s+|all\s+)?(?:primary|key|co-primary|secondary|endpoint)|'
    r'did\s?n[o\u2019\']?t\s+meet|does not meet|not meet(?:ing)?\s+(?:the\s+)?(?:primary|endpoint)|'
    r'negative\s+(?:opinion|topline|result|data|read)|'
    r'(?:FDA|EMA|CHMP|regulators?)\s+reject|reject(?:s|ed)\s+(?:the\s+)?(?:application|filing|BLA|NDA|approval)|'
    r'withdraw(?:s|n|al|ing)?|delist|'
    r'lay(?:s|ing)?\s?off|laid off|layoffs?|cut\s+\d+%?\s+(?:of\s+)?staff|cuts?\s+staff|'
    r'winding down|wind down|shut(?:ting)?\s+down|bankrupt|insolven|'
    r'warning letter|boxed warning|black box|patient death|'
    r'treatment-related death|tx-related death|voluntary recall|product recall|recall of|'
    r'\bsuspend(?:s|ed)?\b|\bpaus(?:e|ed|ing)\b|scrap(?:s|ped)?|abandon(?:s|ed)?|'
    r'disappoint|underperform|cut\s+(?:fy\d+\s+)?(?:revenue\s+)?guidance|'
    r'refus(?:e|al) to file|\bRTF\b'
    r')')
def is_negative(blob):
    return bool(NEGATIVE_RX.search(blob))

# Leading company / sponsor name of a headline (for the highlight on big,
# non-negative stories). Handles single names, camelCase (eGenesis), acronyms
# (GSK, AZ), slash / & / + partnerships (Merck/Moderna, Bausch + Lomb, J&J),
# and drops a trailing possessive 's (highlight "Novartis", not "Novartis's").
_CO_TOKEN = r"[A-Za-z][\w.\-]*"
_CO_RX = re.compile(r"^\s*(" + _CO_TOKEN + r"(?:\s*[&+/]\s*" + _CO_TOKEN + r")*)")
def extract_company(text):
    m = _CO_RX.match(text or '')
    if not m:
        return None
    co = m.group(1).strip(" .,&+/-")
    if not co or len(co) > 45 or not re.search(r'[A-Za-z]', co):
        return None
    return co

# --------------------------------------------------------------------------
# docx -> list of (level, raw_text) items, plus header lines
# --------------------------------------------------------------------------
def read_lines(path):
    md = subprocess.run(['pandoc', '-t', 'markdown', '--wrap=none', path],
                        capture_output=True, text=True, check=True).stdout
    lines = []
    for ln in md.splitlines():
        if not ln.strip():
            continue
        if ln.strip() in ('```{=html}', '```', '<!-- -->', '```{=html}\n<!-- -->'):
            continue
        if ln.strip().startswith('<!--'):
            continue
        m = re.match(r'^( *)-\s+(.*)$', ln)
        if m:
            level = len(m.group(1)) // 4
            lines.append(('bullet', level, m.group(2)))
        else:
            # non-bullet, non-blank line
            indent = len(ln) - len(ln.lstrip(' '))
            lines.append(('para', indent, ln.strip()))
    return lines

# --------------------------------------------------------------------------
# Build one week's tree
# --------------------------------------------------------------------------
def parse_week(path, meta):
    lines = read_lines(path)
    categories = {'disease': [], 'modality': [], 'theme': []}
    cat_area_index = {}   # (category, area_name) -> area dict

    cur_cat = 'disease'          # default before any marker
    cur_area = None
    # buffer of (rel_level, node) items for the current area to build sub-tree
    raw_nodes = []

    def flush_area():
        nonlocal raw_nodes, cur_area
        if cur_area is None:
            raw_nodes = []
            return
        # Build nested tree from raw_nodes using rel_level
        roots = []
        stack = []  # list of (rel_level, node)
        for rel, node in raw_nodes:
            while stack and stack[-1][0] >= rel:
                stack.pop()
            if stack:
                stack[-1][1]['kids'].append(node)
            else:
                roots.append(node)
            stack.append((rel, node))
        # disease areas -> strict TA -> disease -> stories; others keep nesting
        if cur_area.get('_cat') == 'disease':
            cur_area['children'] = regroup_disease(cur_area['area'], roots)
        else:
            cur_area['children'] = [classify(n) for n in roots]
        raw_nodes = []

    def raw_detail_text(node):
        parts = []
        def rec(n):
            for k in n['kids']:
                parts.append(k['label']); rec(k)
        rec(node)
        return ' '.join(parts)

    def make_story(node):
        dtext = raw_detail_text(node)
        blob = node['full_text'] + ' ' + dtext
        mk, mklabel = market_size(blob)         # market size can live in a sub-bullet
        return {
            'type': 'story',
            'title': node['title_html'],
            'label': node['label'],
            'mkt': mk,
            'mktLabel': mklabel,
            'neg': is_negative(node['full_text']),  # sentiment: headline only
            'company': extract_company(node['full_text']),
            'details': [to_detail(k) for k in node['kids']],
        }

    def regroup_disease(ta_name, roots):
        buckets, order = {}, []
        def add(dis, story):
            if dis not in buckets:
                buckets[dis] = []; order.append(dis)
            buckets[dis].append(story)
        def walk(nodes, chain):
            for n in nodes:
                if n['kids'] and is_group_label(n['label'], len(n['kids'])):
                    walk(n['kids'], chain + [n['label']])
                else:
                    add(resolve_disease(ta_name, n['label'], raw_detail_text(n), chain),
                        make_story(n))
        walk(roots, [])
        ordered = sorted(order, key=lambda d: (1 if d == 'Misc.' else 0,
                                               -len(buckets[d]), d.lower()))
        return [{'type': 'group', 'label': d, 'children': buckets[d]} for d in ordered]

    def classify(node):
        """Return a display node: group {type:'group',...} or story {type:'story',...}."""
        label = node['label']
        kids = node['kids']
        if looks_like_group(label, bool(kids)):
            return {
                'type': 'group',
                'label': label.rstrip(':').strip(),
                'children': [classify(k) for k in kids],
            }
        # It's a story; all descendants become nested details
        dtext = raw_detail_text(node)
        blob = node['full_text'] + ' ' + dtext
        mk, mklabel = market_size(blob)         # market size can live in a sub-bullet
        return {
            'type': 'story',
            'title': node['title_html'],
            'label': label,
            'mkt': mk,
            'mktLabel': mklabel,
            'neg': is_negative(node['full_text']),  # sentiment: headline only
            'company': extract_company(node['full_text']),
            'details': [to_detail(k) for k in kids],
        }

    def to_detail(node):
        return {'html': node['title_html'], 'children': [to_detail(k) for k in node['kids']]}

    def get_area(cat, name):
        disp = canon_ta(name) if cat == 'disease' else name.strip()
        key = (cat, disp.lower())
        if key not in cat_area_index:
            area = {'area': disp, 'children': [], '_cat': cat}
            cat_area_index[key] = area
            categories[cat].append(area)
        return cat_area_index[key]

    def make_node(raw):
        lab = label_of(raw)
        return {'label': lab, 'title_html': to_html_inline(raw),
                'full_text': label_of(raw), 'kids': []}

    i = 0
    while i < len(lines):
        kind, a, raw = lines[i]
        if kind == 'para':
            plain = label_of(raw)
            low = plain.lower().strip().rstrip(':')
            if is_underlined_header(raw) and low in BUCKET_MARKERS:
                flush_area(); cur_area = None
                cur_cat = BUCKET_MARKERS[low]
                i += 1; continue
            if low in BUCKET_MARKERS:
                flush_area(); cur_area = None
                cur_cat = BUCKET_MARKERS[low]
                i += 1; continue
            # Italic-underline sub-divider (e.g. "Cancer GI") inside an existing
            # area: NOT a new area. Stories keep flowing into the current area and
            # are regrouped by the disease inferred from each story's own text, so
            # the organ label is redundant. (Only promote it to an area if there's
            # no current area to attach to.)
            if is_italic_subheader(raw) and cur_area is not None:
                i += 1; continue
            # area heading (known -> fixes category; unknown plain heading -> new area in cur_cat)
            if low in AREA_CATEGORY:
                flush_area()
                cur_cat = AREA_CATEGORY[low]
                cur_area = get_area(cur_cat, plain)
                cur_area['_base'] = -1
            else:
                # short-ish plain heading with no digits -> treat as a new area.
                # Unknown named buckets are (in these notes) business/theme items,
                # so default them to theme rather than the previous category.
                if word_count(plain) <= 5 and not re.search(r'\d', plain):
                    flush_area()
                    cur_cat = 'theme'
                    cur_area = get_area(cur_cat, plain)
                    cur_area['_base'] = -1
                # else ignore stray paragraph
            i += 1; continue

        # kind == 'bullet'
        level = a
        plain = label_of(raw)
        low = plain.lower().strip().rstrip(':')
        # A level-0 bullet becomes a NEW top-level AREA only in "bullet-style area
        # list" context: right after a bucket marker (cur_area is None) or when the
        # current area itself came from a bullet (base==0, i.e. a sibling area).
        # If the current area came from a plain-text heading (base==-1), level-0
        # bullets are its CHILDREN -- even if their name matches a known area
        # (e.g. Kidney/Neuro/Skin sub-groups under Immuno).
        base = cur_area.get('_base') if cur_area else None
        if level == 0 and (cur_area is None or base == 0):
            new_cat = AREA_CATEGORY.get(low, cur_cat)
            # only treat as an area if it looks like an area label (short, group-ish)
            if low in AREA_CATEGORY or (word_count(plain) <= 5 and not STORY_MARKERS.search(plain)):
                flush_area()
                cur_cat = new_cat
                cur_area = get_area(cur_cat, plain)
                cur_area['_base'] = 0
                i += 1; continue

        if cur_area is None:
            # bullets before any area -> stash under an implicit area for the category
            cur_area = get_area(cur_cat, _default_area(cur_cat))
            cur_area['_base'] = 0
        rel = level - (cur_area.get('_base', -1) + 1)
        if rel < 0:
            rel = 0
        raw_nodes.append((rel, make_node(raw)))
        i += 1

    flush_area()
    # drop empty areas & clean helper keys
    for cat in categories:
        cleaned = []
        for area in categories[cat]:
            area.pop('_base', None)
            area.pop('_cat', None)
            if area['children']:
                cleaned.append(area)
        categories[cat] = cleaned
    return {'meta': meta, 'categories': categories}

def _default_area(cat):
    return {'disease': 'General', 'modality': 'Platforms', 'theme': 'General'}[cat]

# --------------------------------------------------------------------------
# Week metadata from filename  (Mon_YYYY_WN.docx)
# --------------------------------------------------------------------------
def week_meta(path):
    base = os.path.splitext(os.path.basename(path))[0]
    m = re.match(r'([A-Za-z]+)[_\- ](\d{4})[_\- ]W(\d+)', base)
    if m:
        mon, year, wk = m.group(1), int(m.group(2)), int(m.group(3))
        mnum = MONTHS.get(mon.lower(), 1)
        # approximate a date: week N ~ day (N-1)*7 + 4
        day = min((wk - 1) * 7 + 4, calendar.monthrange(year, mnum)[1])
        sort = year * 10000 + mnum * 100 + day
        return {
            'id': base,
            'label': f'{calendar.month_abbr[mnum]} {year} · W{wk}',
            'short': f'{calendar.month_abbr[mnum]} W{wk}',
            'year': year, 'month': mnum, 'week': wk,
            'date_iso': f'{year:04d}-{mnum:02d}-{day:02d}',
            'sort': sort,
        }
    return {'id': base, 'label': base, 'short': base, 'sort': 0, 'date_iso': ''}

# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def build(paths):
    weeks = [parse_week(p, week_meta(p)) for p in paths]
    weeks.sort(key=lambda w: w['meta']['sort'])
    return {'weeks': weeks, 'generated': True}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('files', nargs='+')
    ap.add_argument('--json', help='write normalized data JSON here')
    ap.add_argument('--html-template', help='HTML template containing /*__DATA__*/')
    ap.add_argument('--html-out', help='write final self-contained dashboard here')
    args = ap.parse_args()

    # Accept folders and wildcards, not just explicit file paths. This lets the
    # updater scripts point at a whole "notes" folder regardless of OS/shell.
    import glob as _glob
    resolved = []
    for a in args.files:
        if os.path.isdir(a):
            resolved += sorted(_glob.glob(os.path.join(a, '*.docx')))
        elif any(c in a for c in '*?['):
            resolved += sorted(_glob.glob(a))
        else:
            resolved.append(a)
    # de-dupe, drop temp lock files (~$...docx), keep order
    seen, files = set(), []
    for f in resolved:
        b = os.path.basename(f)
        if b.startswith('~$') or f in seen:
            continue
        seen.add(f); files.append(f)
    if not files:
        sys.exit('No .docx files found. Put your notes (e.g. Oct_2025_W4.docx) '
                 'in the notes folder and try again.')

    data = build(files)
    if args.json:
        with open(args.json, 'w') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print('wrote', args.json)
    if args.html_template and args.html_out:
        tpl = open(args.html_template).read()
        payload = json.dumps(data, ensure_ascii=False)
        out = tpl.replace('/*__DATA__*/', 'window.BIOTECH_DATA = ' + payload + ';')
        with open(args.html_out, 'w') as f:
            f.write(out)
        print('wrote', args.html_out)
    if not args.json and not args.html_out:
        print(json.dumps(data, indent=2, ensure_ascii=False))

if __name__ == '__main__':
    main()
