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

# --------------------------------------------------------------------------
# Indication subgroups: a logical mid-layer between a therapeutic area (TA) and
# its indications, so a crowded TA like Cancer reads as GI / Thoracic / Heme etc.
# SUBGROUP_ORDER fixes the subgroup display order per TA (same in every week, so
# columns stay aligned). SUBGROUP_MEMBERS maps a subgroup -> its indications.
# Any indication in a mapped TA that isn't listed (incl. 'Misc.') falls into the
# TA's last subgroup ('Other'). A TA with no entry here stays flat (no subgroups).
# --------------------------------------------------------------------------
SUBGROUP_DEFS = {
    'Cancer': [
        ('Thoracic',        ['NSCLC', 'SCLC', 'Mesothelioma', 'Lung cancer']),
        ('GI',              ['Pancreatic cancer', 'Colorectal cancer', 'Esophageal cancer',
                             'Gastric cancer', 'Gastric/GEJ cancer', 'Biliary tract cancer',
                             'Hepatocellular carcinoma', 'Liver cancer (HCC)',
                             'Gastroesophageal cancer', 'Anal cancer', 'GI cancer (other)']),
        ('GU',              ['Prostate cancer', 'Bladder cancer', 'Renal cell carcinoma',
                             'Urothelial cancer', 'Testicular cancer']),
        ('Breast & gyn',    ['Breast cancer', 'Ovarian cancer', 'Cervical cancer',
                             'Endometrial cancer', 'Uterine cancer']),
        ('Skin',            ['Melanoma', 'Uveal melanoma', 'Merkel cell carcinoma',
                             'Cutaneous squamous cell carcinoma', 'Basal cell carcinoma',
                             'Skin cancer (CSCC/BCC)']),
        ('Heme malignancy', ['Multiple myeloma', 'Lymphoma', 'AML', 'ALL', 'CLL', 'CML',
                             'Leukemia', 'Leukemia (other)', 'Polycythemia vera', 'Myelofibrosis',
                             'Myelodysplastic syndrome', 'MDS']),
        ('Head & neck',     ['Head & neck cancer']),
        ('Neuro-oncology',  ['Glioblastoma', 'GBM', 'Brain cancer']),
        ('Other solid',     ['Solid tumors (various)', 'Neuroendocrine tumors', 'Sarcoma',
                             'Thyroid cancer', 'Cancer cachexia']),
        ('Screening & Dx',  ['Multi-cancer detection']),
        ('Other', []),
    ],
    'Cardio': [
        ('Lipids & athero', ['Atherosclerosis / lipids', 'Dyslipidemia', 'Lp(a)']),
        ('Heart muscle',    ['Cardiomyopathy', 'Heart failure']),
        ('Vascular',        ['Hypertension', 'Stroke', 'Pulmonary hypertension', 'Thrombosis']),
        ('Structural',      ['Valve disease']),
        ('Other', []),
    ],
    'Immuno': [
        ('Skin & allergy',  ['Atopic dermatitis', 'Hidradenitis (HS)', 'Chronic urticaria',
                             'Prurigo nodularis', 'Vitiligo', 'Alopecia']),
        ('Rheumatology',    ['Lupus / lupus nephritis', 'Myositis / IBM', 'Rheumatoid arthritis',
                             'Vasculitis (ANCA)', 'Psoriatic arthritis', 'Scleroderma',
                             'Sjogren’s', "Sjogren's"]),
        ('Respiratory',     ['Asthma', 'Bronchiectasis', 'COPD', 'Respiratory (immune)']),
        ('GI',              ['IBD / Crohn’s', "IBD / Crohn's", 'Ulcerative colitis', 'Celiac']),
        ('Renal',           ['IgA nephropathy', 'Membranous nephropathy']),
        ('Neuro-immune',    ['Multiple sclerosis', 'Myasthenia gravis', 'NMOSD']),
        ('Other', []),
    ],
    'Infectious': [
        ('Viral',           ['HIV', 'Hepatitis', 'COVID-19', 'Influenza', 'Measles', 'Ebola',
                             'RSV', 'Dengue', 'Mpox', 'CMV']),
        ('Bacterial & fungal', ['C. difficile', 'Prosthetic joint infection', 'Tuberculosis',
                             'UTI', 'Gonorrhea', 'Fungal']),
        ('Vaccines & policy', ['Vaccines', 'Policy & surveillance', 'RFK Jr.']),
        ('Other', []),
    ],
    'Metabolic': [
        ('Cardiometabolic', ['Obesity', 'Diabetes', 'MASH/NASH', 'Dyslipidemia', 'NAFLD']),
        ('Rare metabolic',  ['ENPP1 deficiency', 'Fabry disease', 'Pompe disease',
                             'Lysosomal / MPS', 'Prader-Willi', 'Gaucher', 'Wilson disease',
                             'Homocystinuria', 'PKU']),
        ('Other', []),
    ],
    'Neuro': [
        ('Neurodegenerative', ["Alzheimer's", "Parkinson's", 'ALS', 'Huntington',
                             'Frontotemporal dementia', 'Multiple system atrophy']),
        ('Seizure & sleep', ['Epilepsy', 'Narcolepsy', 'Insomnia']),
        ('Other', []),
    ],
    'Psych': [
        ('Mood & anxiety',  ['Depression / TRD', 'Bipolar', 'Anxiety / GAD', 'PTSD']),
        ('Psychotic',       ['Schizophrenia']),
        ('Other', []),  # ADHD, substance use, Misc.
    ],
    'Muscle': [
        ('Dystrophies',     ['Duchenne (DMD)', 'FSHD', 'Myotonic dystrophy', 'Becker',
                             'Limb-girdle']),
        ('Motor neuron',    ['Spinal muscular atrophy']),
        ('Other', []),
    ],
}
# flatten to lookup tables
SUBGROUP_ORDER = {ta: [sg for sg, _ in defs] for ta, defs in SUBGROUP_DEFS.items()}
SUBGROUP_OF = {}
for _ta, _defs in SUBGROUP_DEFS.items():
    m = {}
    for _sg, _members in _defs:
        for _ind in _members:
            m[_ind.lower()] = _sg
    SUBGROUP_OF[_ta] = m

def subgroup_for(ta, indication):
    """Return the subgroup label for an indication within a TA, or None if the TA
    has no subgroup scheme defined."""
    if ta not in SUBGROUP_OF:
        return None
    return SUBGROUP_OF[ta].get(indication.strip().lower(), 'Other')

def apply_subgroups(area):
    """Wrap a disease area's indication groups under subgroup nodes, in place.
    No-op for TAs without a subgroup scheme, or areas with <2 subgroups populated."""
    ta = area.get('area')
    if ta not in SUBGROUP_ORDER:
        return
    groups = [g for g in area['children'] if g.get('type') == 'group']
    if not groups:
        return
    order = SUBGROUP_ORDER[ta]
    buckets = {}
    for g in groups:
        sg = subgroup_for(ta, g['label']) or 'Other'
        buckets.setdefault(sg, []).append(g)
    # if everything landed in one subgroup, keep the area flat
    if len([sg for sg in buckets if buckets[sg]]) < 2:
        return
    out = []
    for idx, sg in enumerate(order):
        if buckets.get(sg):
            out.append({'type': 'subgroup', 'label': sg, 'order': idx,
                        'children': buckets[sg]})
    # any subgroup produced but not in the fixed order (shouldn't happen) -> append
    for sg in buckets:
        if sg not in order and buckets[sg]:
            out.append({'type': 'subgroup', 'label': sg, 'order': len(order),
                        'children': buckets[sg]})
    area['children'] = out

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
    ('Biliary tract cancer', r'BTC|(?i:biliary|cholangio|gallbladder)'),
    ('Esophageal cancer', r'ESCC|(?i:esophageal|oesophageal)'),
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
    # generic organ-group buckets (kept AFTER the specific indications so a specific
    # match always wins; used mainly for an explicit "for <group>" headline clause)
    ('GI cancer (other)', r'(?i:\bGI (?:cancers?|tumou?rs?|malignanc)|gastrointestinal (?:cancers?|tumou?rs?))'),
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

def infer_disease(ta, text, exclude=None):
    """The disease a piece of text is about. When several are named (e.g. a
    'for melanoma/lung/esophageal' pan-tumor headline), prefer the one mentioned
    EARLIEST, which is almost always the primary indication; pattern-list order
    only breaks ties. Scanning by text position (not list order) keeps a drug from
    being filed under whichever indication happens to sit higher in the table."""
    best = None  # (start_pos, list_index, name)
    for idx, (name, rx) in enumerate(DISEASE_PATTERNS.get(ta, [])):
        if exclude and name in exclude:
            continue
        m = rx.search(text)
        if m and (best is None or m.start() < best[0]):
            best = (m.start(), idx, name)
    return best[2] if best else None

# Generic "group" buckets that should be refined to a specific indication from the
# details when possible -> maps the bucket to its subgroup so we only pull a MORE
# specific disease from the SAME organ system (never a comparator from elsewhere).
GENERIC_BUCKETS = {
    'GI cancer (other)': ('Cancer', 'GI'),
    'Solid tumors (various)': None,   # genuinely non-specific; don't refine
}

# The indication a headline explicitly states, from its "... for <X> ..." clause.
# This is what the drug is FOR, and it must win over disease names that appear
# elsewhere in the story (comparators, prior trials, other assets in the note).
_FOR_RX = re.compile(r'\bfor\s+(?:the\s+treatment\s+of\s+|patients?\s+with\s+)?(.{2,70}?)(?=\s*[:;(]|$)', re.I)
def indication_phrase(title):
    m = _FOR_RX.search(title or '')
    if not m:
        return None
    phrase = m.group(1).strip()
    # drop obvious non-indication "for" clauses (deals, money, timing)
    if re.search(r'(?i)\$|\bupfront\b|\bmilestone|\bbiobucks|\bstake\b|\bseries [A-E]\b|\bIPO\b', phrase):
        return None
    return phrase

def refine_generic(ta, generic_name, detail_text):
    """When the stated indication is a generic group (e.g. 'GI cancers'), pull the
    specific disease from the details, but ONLY from within the same organ-system
    subgroup so a comparator (e.g. an NSCLC trial) can never hijack it."""
    spec = GENERIC_BUCKETS.get(generic_name)
    if not spec:
        return None
    gta, subgroup = spec
    for name, rx in DISEASE_PATTERNS.get(gta, []):
        if name == generic_name:
            continue
        if subgroup_for(gta, name) != subgroup:
            continue
        if rx.search(detail_text or ''):
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
    """Canonical disease for one story. Priority, highest first:
      1. the indication the headline explicitly states ('... for <X> ...'),
         refined to a specific disease from the details when it's a generic group;
      2. any disease named in the rest of the headline;
      3. a disease named in the details (comparators live here, so this is last
         among text sources);
      4. the nearest non-container grouping header; else 'Misc.'.
    Anchoring on the stated indication stops a comparator drug's disease (e.g. an
    NSCLC trial cited inside a GI-cancer story) from hijacking the classification."""
    # 1) the headline's explicit "for <indication>" clause
    phrase = indication_phrase(title)
    if phrase:
        d = infer_disease(ta, phrase)
        if d:
            return refine_generic(ta, d, detail_text) or d
    # 2) anywhere else in the headline
    d = infer_disease(ta, title)
    if d: return d
    # 3) the details (last text source: comparators & prior trials live here)
    d = infer_disease(ta, detail_text)
    if d: return d
    # 4) fall back to the nearest real grouping header, else Misc.
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
# A drug is "interesting" at or above this peak-sales figure ($millions).
INTERESTING_MUSD = 750
_UNIT_MUSD = {'M': 1, 'B': 1000, 'T': 1000000, 'K': 0.001}
def _fmt_musd(musd):
    """Human label from a $millions value: 1800 -> '$1.8B', 450 -> '$450M'."""
    if musd >= 1000:
        b = musd / 1000.0
        return '$' + (('%.1f' % b).rstrip('0').rstrip('.')) + 'B'
    return '$' + (('%.0f' % musd)) + 'M'

def _market_explicit(blob):
    """The author's written '= $X' peak-sales / TAM figure, if present."""
    m = re.search(r'=\s*\$?\s?(\d[\d.,]*)\s?([MBT])\s*TAM', blob)
    if not m:
        m = re.search(r'(?i:market size|peak sales)[^=]*=\s*\u00d7?\s*\$?\s?(\d[\d.,]*)\s?([MBT])\b', blob)
    if not m:
        m = re.search(r'=\s*\$\s?(\d[\d.,]*)\s?([MBT])\s*(?:peak|sales|TAM|opportunity)', blob)
    if not m:
        return None
    try:
        val = float(m.group(1).rstrip('.,').replace(',', ''))
    except ValueError:
        return None
    return val * _UNIT_MUSD[m.group(2).upper()]

def _market_computed(blob):
    """Compute peak sales from the component chain on a 'Market size:' line when
    the author didn't write an explicit '= $X'. Pattern:
      'Market size: 100K pts * 15% * 8% * $2.4M net = ...'
    => count * product(percentages) * net_price. Returns $millions or None."""
    mm = re.search(r'(?i)market size\s*:?\s*(.+?)(?:\n|$)', blob)
    if not mm:
        return None
    seg = mm.group(1)
    seg = seg.split('=')[0]  # if an '=' total existed, _market_explicit handled it
    # population count (first number with K/M/B magnitude)
    cm = re.search(r'(\d[\d.,]*)\s*([KMB])\b', seg)
    if not cm:
        return None
    try:
        count = float(cm.group(1).replace(',', '')) * {'K':1e3,'M':1e6,'B':1e9}[cm.group(2)]
    except ValueError:
        return None
    # net price per patient ($XK or $XM)
    pm = re.search(r'\$\s?(\d[\d.,]*)\s*([KM])\b(?:[^%]*?net)?', seg)
    if not pm:
        return None
    try:
        price = float(pm.group(1).replace(',', '')) * {'K':1e3,'M':1e6}[pm.group(2)]
    except ValueError:
        return None
    # multiply by every percentage in the chain
    frac = 1.0
    for p in re.findall(r'(\d[\d.]*)\s*%', seg):
        try: frac *= float(p) / 100.0
        except ValueError: pass
    musd = count * frac * price / 1e6
    if musd <= 0 or musd > 500000:   # sanity bound
        return None
    return musd

def market_size(blob):
    """(value in $millions, display label) or (None, None). Prefers the author's
    explicit '= $X'; falls back to computing it from the component chain."""
    musd = _market_explicit(blob)
    computed = False
    if musd is None:
        musd = _market_computed(blob)
        computed = True
    if musd is None:
        return (None, None)
    lab = _fmt_musd(musd)
    if computed:
        lab = '~' + lab      # tilde marks a Claude-computed (not author-written) figure
    return (musd, lab)

# --------------------------------------------------------------------------
# Standardized market-size ESTIMATE for assets the notes left unquantified.
# Same shape as the authors' own math (addressable patients x peak share x net
# annual price), but driven by a built-in epidemiology/pricing reference so every
# drug gets a comparable figure. Flagged separately ('est') so it's never
# confused with a doc-provided or doc-derived number.
# Values are rough, defensible US peak-sales order-of-magnitude anchors:
#   (addressable drug-treated US patients / yr, net annual price $)
# --------------------------------------------------------------------------
MARKET_REF = {
    # ---- oncology (incident, drug-treated) ----
    'nsclc': (85000, 160000), 'sclc': (20000, 130000), 'mesothelioma': (3000, 150000),
    'breast cancer': (90000, 130000), 'prostate cancer': (90000, 110000),
    'colorectal cancer': (90000, 130000), 'pancreatic cancer': (50000, 150000),
    'esophageal cancer': (30000, 140000), 'gastric/gej cancer': (27000, 150000),
    'gastric cancer': (27000, 150000), 'biliary tract cancer': (12000, 160000),
    'hepatocellular carcinoma': (30000, 150000), 'melanoma': (60000, 180000),
    'uveal melanoma': (2500, 180000), 'bladder cancer': (60000, 140000),
    'renal cell carcinoma': (45000, 170000), 'head & neck cancer': (40000, 140000),
    'ovarian cancer': (20000, 150000), 'cervical cancer': (13000, 130000),
    'endometrial cancer': (30000, 130000), 'glioblastoma': (13000, 150000),
    'multiple myeloma': (35000, 200000), 'lymphoma': (60000, 180000),
    'aml': (20000, 180000), 'leukemia': (40000, 160000), 'leukemia (other)': (40000, 160000),
    'cll': (25000, 150000), 'polycythemia vera': (60000, 60000),
    'myelofibrosis': (18000, 150000), 'neuroendocrine tumors': (20000, 150000),
    'sarcoma': (15000, 150000), 'skin cancer (cscc/bcc)': (120000, 90000),
    'solid tumors (various)': (60000, 150000),
    # ---- cardiometabolic ----
    'obesity': (8000000, 5500), 'diabetes': (6000000, 4000), 'mash/nash': (1500000, 15000),
    'dyslipidemia': (6000000, 5000), 'atherosclerosis / lipids': (6000000, 5000),
    'hypertension': (10000000, 1800), 'heart failure': (3000000, 6000),
    'cardiomyopathy': (150000, 90000), 'stroke': (500000, 8000),
    'pulmonary hypertension': (40000, 90000),
    # ---- immunology / inflammation ----
    'atopic dermatitis': (900000, 35000), 'asthma': (1000000, 25000),
    'copd': (1200000, 20000), 'bronchiectasis': (300000, 30000),
    'rheumatoid arthritis': (1000000, 40000), 'lupus / lupus nephritis': (200000, 45000),
    'myositis / ibm': (40000, 60000), 'vasculitis (anca)': (40000, 60000),
    'ibd / crohn’s': (700000, 45000), 'iga nephropathy': (120000, 60000),
    'hidradenitis (hs)': (150000, 35000), 'multiple sclerosis': (300000, 70000),
    'thyroid eye disease': (60000, 200000),
    # ---- neuro / psych ----
    'alzheimer’s': (2000000, 28000), "alzheimer's": (2000000, 28000),
    'parkinson’s': (800000, 20000), "parkinson's": (800000, 20000),
    'als': (25000, 150000), 'huntington': (30000, 90000), 'epilepsy': (1200000, 12000),
    'narcolepsy': (165000, 18000), 'migraine': (4000000, 8000),
    'depression / trd': (2000000, 10000), 'schizophrenia': (1500000, 15000),
    'bipolar': (1500000, 10000), 'anxiety / gad': (3000000, 6000),
    # ---- rare / muscle / blood ----
    'duchenne (dmd)': (12000, 400000), 'spinal muscular atrophy': (10000, 300000),
    'myotonic dystrophy': (40000, 90000), 'fshd': (16000, 90000),
    'hemophilia / bleeding': (20000, 500000), 'sickle cell': (70000, 150000),
    'fabry disease': (5000, 350000), 'pompe disease': (3500, 400000),
    'lysosomal / mps': (5000, 400000), 'enpp1 deficiency': (600, 400000),
    'prader-willi': (18000, 60000),
    # ---- infectious ----
    'hiv': (1000000, 25000), 'hepatitis': (2500000, 25000), 'covid-19': (3000000, 1500),
    'influenza': (5000000, 800), 'rsv': (3000000, 300),
}
# per-therapeutic-area fallback when a specific indication isn't in the table
AREA_REF = {
    'Cancer': (40000, 150000), 'Immuno': (400000, 45000), 'Metabolic': (400000, 10000),
    'Cardio': (2000000, 4000), 'Neuro': (500000, 30000), 'Psych': (1500000, 12000),
    'Infectious': (800000, 6000), 'Muscle': (12000, 300000), 'Blood': (30000, 300000),
    'Eye': (300000, 15000), 'Kidney': (150000, 60000), 'Liver': (800000, 20000),
    'Lung': (800000, 25000), 'Skin': (500000, 20000), 'Bone': (300000, 25000),
    'Pain': (3000000, 5000), 'Respiratory': (800000, 25000), "Women's health": (2000000, 4000),
}
DEFAULT_REF = (60000, 60000)
# net-price multiplier by modality (premium one-time / cell-gene vs chronic small mol)
PRICE_MULT = {
    'Gene therapy': 1.6, 'Base editing': 1.6, 'CRISPR/editing': 1.5, 'CAR-T': 1.4,
    'TCR-T': 1.4, 'Cell therapy': 1.4, 'Enzyme/ERT': 1.4, 'Radiopharma': 1.2,
    'ADC': 1.2, 'Bispecific Ab': 1.15, 'T-cell engager': 1.2, 'siRNA/RNAi': 1.1,
    'ASO/antisense': 1.1, 'Small molecule': 0.9,
}
PEAK_SHARE = 0.18   # a credible single-asset peak share of its addressable pool
# indications where simple epi x price doesn't model the economics -> no estimate
_NO_EST_DISEASE = {'multi-cancer detection', 'misc.', 'policy & surveillance',
                   'vaccines', 'rfk jr.', 'solid tumors (various)',
                   'ebola', 'measles'}   # outbreak/stockpile, not a commercial market
# modalities that aren't priced as a chronic per-patient therapeutic -> no estimate
# (diagnostics, devices, and vaccines have very different economics; a pure
# discovery/platform story has no modality and is left as an unsized "concept")
_NO_EST_MODALITY = {'Diagnostic', 'BCI/device', 'Vaccine'}

def estimate_market(ta, disease, modality):
    """A standardized peak-sales estimate (value_$M, label, derivation) for a
    concrete therapeutic the notes didn't size, or None when it can't/shouldn't be
    modeled (unspecified indication, platform/discovery story, or non-drug economics)."""
    if not disease or not modality or modality in _NO_EST_MODALITY:
        return None
    dl = disease.strip().lower()
    if dl in _NO_EST_DISEASE:
        return None
    ref = MARKET_REF.get(dl) or AREA_REF.get(ta) or DEFAULT_REF
    pts, price = ref
    mult = PRICE_MULT.get(modality, 1.0)
    musd = pts * PEAK_SHARE * price * mult / 1e6
    if musd <= 0:
        return None
    # round to 2 significant figures so estimates read as estimates, not false precision
    import math
    digits = int(math.floor(math.log10(musd)))
    musd = round(musd, -(digits - 1)) if digits >= 1 else round(musd, 1)
    label = '~' + _fmt_musd(musd) + ' est'
    ppl = (('%.0f' % (pts / 1000.0)) + 'K') if pts < 1e6 else (('%.1f' % (pts / 1e6)).rstrip('0').rstrip('.') + 'M')
    pr = ('$%.1fM' % (price / 1e6)).replace('.0M', 'M') if price >= 1e6 else ('$%dK' % round(price / 1e3))
    deriv = ('Model estimate: ' + ppl + ' addressable US pts × ' + ('%d' % (PEAK_SHARE * 100)) +
             '% peak share × ' + pr + '/yr net' +
             (' × ' + ('%.2g' % mult) + ' (' + modality + ')' if mult != 1.0 and modality else '') +
             ' ≈ ' + _fmt_musd(musd))
    return (musd, label, deriv)

# -------- asset / mechanism / development-stage extraction --------
# Modality (drug class). First match wins, so order specific -> generic.
_MODALITY = [
    ('Radiopharma',        r'\bRLT\b|radioligand|radiopharm|Lu-?177|\bAc-?225\b|actinium|\b177Lu\b'),
    ('CAR-T',              r'CAR-?T\b|chimeric antigen receptor'),
    ('TCR-T',              r'\bTCR-?T\b|TCR therapy|TCR-engineered'),
    ('T-cell engager',     r'\bTCE\b|\bBiTE\b|T-?cell engager|CD3 bispecific|\bTCB\b'),
    ('ADC',                r'\bADC\b|antibody-?drug conjugate'),
    ('Base editing',       r'base edit|base-?edited'),
    ('CRISPR/editing',     r'CRISPR|prime edit|gene-?edit'),
    ('Gene therapy',       r'gene therapy|\bAAV\b|lentiviral|adeno-?associated'),
    ('Cell therapy',       r'cell therapy|\bHSC\b|stem-?cell|allogeneic|autologous|\bNK cell'),
    ('siRNA/RNAi',         r'siRNA|RNAi|\bsiRNA\b'),
    ('ASO/antisense',      r'\bASO\b|antisense'),
    ('mRNA',               r'\bmRNA\b|\bsaRNA\b|\btaRNA\b|self-amplifying'),
    ('Protein degrader',   r'degrader|PROTAC|molecular glue|\bRIPTAC\b'),
    ('Bispecific Ab',      r'bispecific|biparatopic|trispecific'),
    ('Peptide',            r'\bpeptide\b|peptide-'),
    ('Monoclonal Ab',      r'\bantibod(?:y|ies)\b|\bmAb\b|-mab\b|nanobody'),
    ('Enzyme/ERT',         r'\bERT\b|enzyme replacement'),
    ('Vaccine',            r'\bvaccine\b|immunization'),
    ('BCI/device',         r'\bBCI\b|brain-?computer|neurostimulat|endovascular|\bimplant\b|wearable|\bsensor\b'),
    ('Diagnostic',         r'\bdiagnostic\b|companion diagnostic|\bassay\b|liquid biopsy|screening (?:test|platform)|\bMCED\b'),
    ('Small molecule',     r'small molecule|\binhibitor\b|\bagonist\b|\bantagonist\b|\bactivator\b|\bblocker\b|-ib\b|-stat\b|-rasib\b|kinase'),
]
_MODALITY = [(n, re.compile(p)) for n, p in _MODALITY]
def extract_modality(blob):
    for name, rx in _MODALITY:
        if rx.search(blob):
            return name
    return None

# Drug target / pathway (for the mechanism landscape and the breakthrough brief).
# Curated high-frequency biotech targets; first match wins. Case-sensitive so
# acronyms don't match lowercase words. Edit/extend as coverage grows.
_TARGETS = [
    r'PD-?L1', r'PD-?1', r'CTLA-?4', r'TIGIT', r'LAG-?3', r'CD3', r'CD19', r'CD20',
    r'CD22', r'CD38', r'CD28', r'CD47', r'CD123', r'CD228', r'CD79[ab]?', r'BCMA',
    r'GPRC5D', r'CLDN18\.2', r'CLDN6', r'DLL3', r'HER2', r'HER3', r'EGFR', r'cMET|c-MET|\bMET\b',
    r'KRAS', r'NRAS', r'\bRAS\b', r'BRAF', r'PI3K', r'AKT', r'mTOR', r'CDK4/6|CDK2|CDK7',
    r'PARP', r'BTK', r'TYK2', r'JAK1|JAK2|JAK', r'SYK', r'Nectin-?4', r'TROP2|TROP-?2',
    r'FR[aα]|folate receptor', r'PSMA', r'SSTR2', r'FAP', r'B7-?H3', r'5T4',
    r'TL1A', r'IL-?31', r'IL-?17[A-F]?', r'IL-?23', r'IL-?13', r'IL-?4R', r'IL-?5',
    r'IL-?6', r'IL-?15', r'IL-?2', r'TSLP', r'TNF', r'OX40L?|OX40', r'BAFF', r'APRIL',
    r'FcRn', r'C5\b|complement', r'IgE', r'CCR8', r'CCR5', r'CSF1R', r'KIT',
    r'GLP-?1', r'GIP', r'glucagon', r'amylin', r'UCN2', r'MC4R', r'MC1R', r'leptin',
    r'FGF21', r'APOC3', r'ANGPTL3', r'PCSK9', r'Lp\(a\)|LPA\b', r'CETP', r'FXIa|FXI\b',
    r'factor XI', r'TTR', r'myosin', r'aldosterone', r'ENaC', r'sGC',
    r'CGRP', r'Nav1\.[78]', r'KCNT1|Kv7', r'LRRK2', r'GBA1', r'alpha-?synuclein|a-?synuclein',
    r'amyloid|Aβ', r'tau\b', r'TDP-?43', r'SOD1', r'HTT|huntingtin', r'MECP2',
    r'D2\b|dopamine', r'5-?HT2A|serotonin', r'orexin', r'NMDA', r'GABA',
    r'CFTR', r'DMD|dystrophin', r'SMN', r'DUX4', r'SERPINA1|AATD',
    r'BCL11A', r'HBV', r'TMPRSS6|hepcidin', r'JAK2 V617F', r'FLT3', r'IDH1|IDH2',
    r'menin', r'BCL-?2', r'NLRP3', r'STING', r'cGAS', r'URAT1', r'ENPP1', r'IGF-?1R',
    r'TSHR', r'RXFP1|relaxin', r'ActRII|activin', r'GDF15', r'SIRP[aα]', r'WT1',
    r'VEGF', r'VEGFR2?', r'ALK', r'ROS1', r'RET', r'FGFR', r'NTRK',
]
_TARGET_RX = [re.compile(r'\b(?:' + p + r')\b') for p in _TARGETS]
_TARGET_NAME = [re.split(r'[\\|]', p)[0].replace('(','').replace(')','').replace('?','')
                for p in _TARGETS]
def extract_target(blob):
    for i, rx in enumerate(_TARGET_RX):
        if rx.search(blob):
            return _TARGET_NAME[i]
    return None

# --- patient-segment extraction: line of therapy, histology, biomarker selection ---
# Powers the frontier coverage map, the whitespace callouts and the ideation brief.
_SEG_LINE = [
    ('early/adjuvant', r'(?i:neoadjuvant|adjuvant|perioperative|resected|resectable|localized|localised|early-?stage|non-metastatic|\bnmCRPC\b|\bnmCSPC\b)'),
    ('1L',            r'(?i:\b1L\b|first-?line|frontline|front-?line|treatment-?na[iï]ve|newly diagnosed|untreated|hormone-?sensitive|\bmHSPC\b|\bmCSPC\b)'),
    ('2L+/later',     r'(?i:\b2L\b|\b3L\b|2L\+|2L/3L|later-?line|previously treated|pre-?treated|post-(?:TKI|Tagrisso|osimertinib|PD-?1|PD-?L1|IO|platinum|chemo|SoC|surgery|radiation)|relapsed|refractory|\bR/R\b|\bRRMM\b|castration-?resistant|\bmCRPC\b)'),
]
_SEG_HISTO = [
    ('squamous',     r'(?i:squamous|\bSCC\b|\bLUSC\b|\bHNSCC\b|\bcSCC\b|\bESCC\b)'),
    ('non-squamous', r'(?i:non-?squamous|non-?SCC|\bLUAD\b|adenocarcinoma)'),
]
# biomarker / molecularly-selected population markers (selection, not just drug target)
_SEG_BIO_RX = re.compile(
    r'(EGFR(?:m| exon ?\d+| ex\d+)?|HER2(?:m|-low|\+|-)?|HER3|KRAS(?: ?G12[CDV])?|\bALK\b|ROS1|BRAF|'
    r'\bMET(?:-OE| exon ?14)?|\bRET\b|NTRK|PIK3CA|HR\+|ER\+|PD-?L1(?:[ -]?(?:high|low|negative|neg|\u2264?\s?\d+%|TC ?<?\d+%))?|'
    r'MSI-?H|dMMR|BRCA|PSMA\+?|TROP2|5T4|CDH6|Claudin ?18\.?2?|CLDN18\.2|BCMA|GPRC5D|FR[a\u03b1]|DLL3|exon ?20|exon20-ins)')
def extract_segments(blob):
    line = [name for name, pat in _SEG_LINE if re.search(pat, blob)]
    histo = [name for name, pat in _SEG_HISTO if re.search(pat, blob)]
    raw = []
    for m in _SEG_BIO_RX.findall(blob):
        tok = re.sub(r'\s+', ' ', m).strip()
        if tok and tok not in raw:
            raw.append(tok)
    # drop a token that is a substring of a more specific one already present
    bio = [t for t in raw if not any(t != o and t.lower() in o.lower() for o in raw)]
    return {'line': line, 'histo': histo, 'bio': bio[:3]}

# Best efficacy read in the text — the benchmark a new drug must beat.
# Returns a short string like "PFS HR 0.40; OS HR 0.48; ORR 74%" or
# "met PFS endpoint", or None.
def extract_efficacy(blob):
    out = []
    for m in re.finditer(r'(?i)(PFS|OS|rPFS|DFS|EFS|DoR)\s*HR\s*(?:of\s*)?(\d?\.\d+)', blob):
        out.append(m.group(1).upper() + ' HR ' + m.group(2))
    for m in re.finditer(r'(\d?\.\d+)\s*(?:PFS|OS)?\s*HR\b', blob):
        tag = 'HR ' + m.group(1)
        if not any('HR ' + m.group(1) in o for o in out):
            out.append(tag)
    m = (re.search(r'(?i)ORR\s*(?:of\s*)?\(?(\d{1,3})%', blob)
         or re.search(r'(?i)\(?(\d{1,3})%\s*(?:vs\s*\d{1,3}%\s*)?ORR', blob)
         or re.search(r'(?i)met ORR \((\d{1,3})%', blob)
         or re.search(r'(?i)response rate of (\d{1,3})%', blob))
    if m: out.append('ORR ' + m.group(1) + '%')
    m = re.search(r'(\d{1,3})%\s*(?:vs\s*\d{1,3}%\s*)?MRD[- ]?(?:neg\w*|CR)?', blob)
    if m: out.append(m.group(1) + '% MRD-CR')
    m = re.search(r'(?i)(\d{1,3})%\s*(?:PFS/OS\s*)?(?:risk\s*)?reduction', blob)
    if m: out.append(m.group(1) + '% risk reduction')
    m = re.search(r'(?i)(\d+\.?\d*)\s*mos?\s*vs\s*(\d+\.?\d*)\s*mos?\s*(OS|PFS)?', blob)
    if m: out.append((m.group(3) or 'OS').upper() + ' ' + m.group(1) + ' vs ' + m.group(2) + 'mo')
    for m in re.finditer(r'(?i)median\s*(PFS|OS)\s*(?:of\s*)?(\d+\.?\d*)\s*(?:mo|months)', blob):
        out.append('m' + m.group(1).upper() + ' ' + m.group(2) + 'mo')
    seen = []
    for x in out:
        if x not in seen:
            seen.append(x)
    if seen:
        return '; '.join(seen[:3])
    # qualitative fallback
    mq = re.search(r'(?i)met (?:the )?(?:co-)?primary\s*(PFS|OS|ORR|endpoint)', blob)
    if mq:
        g = mq.group(1).upper()
        return 'met ' + ('primary endpoint' if g == 'ENDPOINT' else g + ' endpoint')
    return None

# Nature of the news (for the "concepts rising vs failing" summaries).
def news_kind(blob, neg, pos, stage_rank):
    b = blob
    if re.search(r'(?i:discontinu|terminat|shelv|wind(?:ing)? down|deprioriti)', b):
        return 'discontinued'
    if neg and re.search(r'(?i:fail|missed? (?:the )?(?:primary|endpoint)|did ?n[o\u2019\']?t meet|does not meet|negative (?:data|topline|result))', b):
        return 'negative data'
    if neg and re.search(r'(?i:\bCRL\b|not[- ]approvable|clinical hold|withdraw|rejected)', b):
        return 'regulatory setback'
    if neg and re.search(r'(?i:lay ?off|cut \d|cut staff|restructur|going concern)', b):
        return 'restructuring'
    if neg:
        return 'setback'
    if re.search(r'(?i:FDA approv|\bapproved\b|accelerated approval|full approval|CHMP positive|positive opinion|cleared)', b):
        return 'approval'
    if pos or re.search(r'(?i:met (?:the )?(?:primary|co-primary|key)|positive (?:ph|phase|topline|data|readout)|statistically significant|superiority|improved (?:OS|PFS|survival))', b):
        return 'positive data'
    if re.search(r'(?i:\bBLA\b|\bNDA\b|\bMAA\b|PDUFA|priority review|filed|submission|rolling)', b):
        return 'regulatory filing'
    if re.search(r'(?i:\bIPO\b|Series [A-E]\b|raise[sd]?|\bseed\b|\bPIPE\b|financing|\$\d[\d.]*[MB] (?:round|Series))', b):
        return 'financing'
    if re.search(r'(?i:acquir|acquisition|merger|licen[sc]|partnership|option (?:deal|to)|collaborat)', b):
        return 'deal'
    if stage_rank >= 2:
        return 'in development'
    return 'update'

# Development stage, ranked (higher = more advanced). Used for the "frontier".
def extract_stage(blob):
    b = blob
    # a rejection at filing (Not Approvable / CRL) is NOT "approved"
    if re.search(r'(?i:not[- ]approvable|complete response letter|\bCRL\b)', b):
        return ('Filed', 5)
    if re.search(r'\bnet sales\b|\bnet revenue\b|\bsales (?:of )?\$|launch(?:ed|es)?\b|commercial\b|\$\d[\d.]*[MB]\s*\d?[QH]\d', b) \
       and re.search(r'\bFDA|approv|market', b):
        return ('Marketed', 7)
    if re.search(r'FDA approv|approved|approval|cleared|\bsBLA\b|accelerated approval|full approval|CHMP positive|positive opinion', b):
        return ('Approved', 6)
    if re.search(r'\bBLA\b|\bNDA\b|\bMAA\b|PDUFA|priority review|filing|filed|submitt|rolling (?:review|submission)|De Novo', b):
        return ('Filed', 5)
    if re.search(r'Ph\s?3|phase\s?3|Ph\s?2/3|pivotal', b):
        return ('Phase 3', 4)
    if re.search(r'Ph\s?2|phase\s?2|Ph\s?1/2', b):
        return ('Phase 2', 3)
    if re.search(r'Ph\s?1|phase\s?1|first-in-human|\bFIH\b|\bIND\b|healthy volunteer', b):
        return ('Phase 1', 2)
    if re.search(r'preclinical|\bIND-enabling\b|discovery|in vivo|in vitro|mouse|murine|primate', b):
        return ('Preclinical', 1)
    return ('Platform/other', 0)

# Positive-data signal (for status). Negative handled by NEGATIVE_RX.
_POSITIVE_RX = re.compile(
    r'(?i:met (?:the )?(?:primary|co-primary|key)|hit (?:the )?primary|'
    r'positive (?:ph|phase|topline|data|result|readout|opinion)|'
    r'statistically significant|stat-?sig|superiority|superior to|'
    r'FDA approv|approved|accelerated approval|breakthrough (?:therapy|designation)|'
    r'durable response|improved (?:OS|PFS|survival|RFS)|'
    r'\bORR\b.*\d\d%|response rate of \d)')

# Drug-asset fingerprints: company drug-codes (AMG 513, SION-719, HG302) and
# INN drug-name suffixes (-mab, -tinib, -sen, -gene, -cel, -rasib, ...).
_DRUGCODE_RX = re.compile(r'\b[A-Z]{2,5}[- ]?\d{2,6}\b')
_DRUGNAME_RX = re.compile(
    r'(?i:\w(?:mab|tinib|ciclib|parib|rasib|degib|lisib|sertib|metinib|'
    r'anib|afenib|sen|gene|cel|vir|limab|lizumab|ngogene|ersen|nersen|'
    r'kacept|cept|stat|prazan|glutide|tide)\b)')
# Clear business / policy / legal / epidemiology news that is NOT an R&D asset
# story even if it name-drops a drug or modality.
_HARDNEWS_RX = re.compile(
    r'(?i:court|lawsuit|\bFOIA\b|executive order|CDC director|HHS Secretary|'
    r'\bRFK\b|Senate confirm|exemption|outbreak|re-?establish|elimination status|'
    r'screening guideline|guidelines? (?:now )?(?:urge|recommend)|'
    r'guidance (?:from|cut)|cut (?:FY|fy)?\s?\d|revenue (?:guidance|\+|of \$| of €)|'
    r'net sales|pro forma|benefit(?:s)? redesign|strategic plan|party-line|'
    r'denies responsib|CNN interview|chooses .* (?:provider|AWS)|preferred .* provider|'
    r'doubles down|no doubt|school year|supply disruption|patent (?:settle|dispute)|'
    r'executive hub|employer-benefit|cost exposure|'
    r'study (?:of|finds|shows|confirms|suggests)|researchers (?:map|found|show|identif)|'
    r'\batlas\b|protein interactions|research (?:hub|institute)|single-cell|'
    r'launches .* (?:hub|institute)|[1-4][QH]\d\d? (?:revenue|sales)|revenue \(\+|'
    r'\bpivot\b|real-world (?:adoption|data|psilocybin))')
# Positive R&D-asset vocabulary (mechanism / class / trial), independent of stage.
_ASSET_VOCAB_RX = re.compile(
    r'(?i:\bPh\s?[1-3]\b|phase\s?[1-3]|preclinical|\bIND\b|IND-enabling|'
    r'first-in-human|inhibitor|agonist|antagonist|antibod|\bADC\b|CAR-?T\b|'
    r'\bTCE\b|siRNA|\bASO\b|\bmRNA\b|gene therapy|cell therapy|degrader|'
    r'\bpeptide\b|\bRLT\b|base edit|CRISPR|small molecule|bispecific|biparatopic|'
    r'\bERT\b|enzyme replacement|oncolytic|superagonist|\bTKI\b)')
def is_asset_story(blob, modality=None, stage_rank=0):
    """True if the story is a drug / asset / platform worth watching as an R&D
    opportunity (vs policy, legal, epidemiology, M&A or pure commercial news).
    Precision-biased: a clear business/news signal vetoes it."""
    if _HARDNEWS_RX.search(blob):
        return False
    # An R&D asset shows at least one of: a known modality, a clinical stage,
    # explicit asset vocabulary, a drug-code, or an INN drug-name suffix.
    if modality or stage_rank >= 1:
        return True
    if _ASSET_VOCAB_RX.search(blob):
        return True
    if _DRUGCODE_RX.search(blob) or _DRUGNAME_RX.search(blob):
        return True
    return False

# Negative-news detection (colors the headline red). Edit to taste.
NEGATIVE_RX = re.compile(
    r'(?i:'
    r'\bfail(?:s|ed|ure|ing)?\b(?!\s+(?:detection|mode|analysis))|discontinu(?:e|es|ed|ing)\b(?!\s*\d{4})|terminat(?:e|es|ed|ion)|'
    r'\bCRL\b(?!\s+(?:collaborat|partner))|complete response letter|not[- ]approvable|clinical hold|\bhalt(?:s|ed|ing)?\b|'
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
    r'refus(?:e|al) to file|\bRTF\b|deprioriti[sz]|shelv(?:e|ed|ing)|going concern'
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
    # True while the current top-level story was opened by a PARAGRAPH headline
    # (as in the "reorganized" notes). In that mode the bullets that follow are the
    # story's details and must nest one level deeper, instead of each becoming its
    # own sibling story.
    para_anchor = False

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

    def _enrich(node, label):
        """Shared story fields: market size, sentiment, dev stage, modality,
        asset/concept classification, and the 'interesting' flag."""
        dtext = raw_detail_text(node)
        blob = node['full_text'] + ' ' + dtext
        mk, mklabel = market_size(blob)         # market size can live in a sub-bullet
        neg = is_negative(node['full_text'])    # sentiment: headline only
        stage, stage_rank = extract_stage(blob)
        modality = extract_modality(blob)
        asset = is_asset_story(blob, modality, stage_rank)
        positive = bool(_POSITIVE_RX.search(blob)) and not neg
        # market status: quantified / unquantified drug-or-concept / not-applicable
        if mk is not None:
            mstatus = 'quant'
        elif asset:
            mstatus = 'unquant'      # a real asset/platform that lacks a $ figure
        else:
            mstatus = 'na'
        # INTERESTING = big market (>=$750M) OR a concept/platform worth watching
        # that simply has no peak-sales number yet.
        interesting = (mk is not None and mk >= INTERESTING_MUSD) or \
                      (mk is None and asset)
        return {
            'mkt': mk, 'mktLabel': mklabel, 'mstatus': mstatus,
            'neg': neg, 'pos': positive,
            'company': extract_company(node['full_text']),
            'modality': modality, 'target': extract_target(blob),
            'stage': stage, 'stageRank': stage_rank,
            'newsKind': news_kind(blob, neg, positive, stage_rank),
            'efficacy': extract_efficacy(blob),
            'seg': extract_segments(blob),
            'asset': asset, 'interesting': interesting,
            'details': [to_detail(k) for k in node['kids']],
        }

    def make_story(node, ta=None, disease=None):
        f = _enrich(node, node['label'])
        f.update({'type': 'story', 'title': node['title_html'], 'label': node['label']})
        # Standardized estimate for a real asset the notes left unsized, so every
        # drug carries a comparable market figure (clearly flagged as an estimate).
        if f['mkt'] is None and f['asset']:
            est = estimate_market(ta, disease, f['modality'])
            if est:
                val, lab, deriv = est
                f['mkt'] = val; f['mktLabel'] = lab; f['mstatus'] = 'est'; f['mktEst'] = deriv
                f['interesting'] = (val >= INTERESTING_MUSD) or f['interesting']
        return f

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
                    dis = resolve_disease(ta_name, n['label'], raw_detail_text(n), chain)
                    add(dis, make_story(n, ta_name, dis))
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
        f = _enrich(node, label)
        f.update({'type': 'story', 'title': node['title_html'], 'label': label})
        return f

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
                flush_area(); cur_area = None; para_anchor = False
                cur_cat = BUCKET_MARKERS[low]
                i += 1; continue
            if low in BUCKET_MARKERS:
                flush_area(); cur_area = None; para_anchor = False
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
                flush_area(); para_anchor = False
                cur_cat = AREA_CATEGORY[low]
                cur_area = get_area(cur_cat, plain)
                cur_area['_base'] = -1
            elif word_count(plain) <= 5 and not re.search(r'\d', plain):
                # short-ish plain heading with no digits -> treat as a new area.
                # Unknown named buckets are (in these notes) business/theme items,
                # so default them to theme rather than the previous category.
                flush_area(); para_anchor = False
                cur_cat = 'theme'
                cur_area = get_area(cur_cat, plain)
                cur_area['_base'] = -1
            elif cur_area is not None:
                # A longer paragraph inside an area is a STORY HEADLINE (the
                # "reorganized" note style). Push it as a top-level story; the
                # bullets that follow become its details (para_anchor mode).
                raw_nodes.append((0, make_node(raw)))
                para_anchor = True
            # else: stray paragraph before any area -> ignore
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
                flush_area(); para_anchor = False
                cur_cat = new_cat
                cur_area = get_area(cur_cat, plain)
                cur_area['_base'] = 0
                i += 1; continue

        if cur_area is None:
            # bullets before any area -> stash under an implicit area for the category
            cur_area = get_area(cur_cat, _default_area(cur_cat))
            cur_area['_base'] = 0
        if para_anchor:
            # bullets after a paragraph headline are that story's details: one
            # level deeper than the headline (rel 0), so level-0 -> rel 1, etc.
            rel = level + 1
        else:
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
def validate(weeks):
    """Non-fatal sanity checks that run on every build, so a new upload surfaces
    classification problems instead of silently shipping them. Prints warnings to
    stderr; never raises."""
    warns = []
    def plain(h):
        return _html.unescape(re.sub(r'<[^>]+>', '', h or ''))
    def story_text(n):
        def det(ds):
            s = ''
            for d in ds or []:
                s += ' ' + plain(d.get('html', '')) + det(d.get('children'))
            return s
        return plain(n.get('title', '')) + det(n.get('details'))
    for w in weeks:
        wl = w['meta']['label']
        # map subgroup -> indication grouping & check every indication is placed
        for area in w['categories'].get('disease', []):
            ta = area['area']
            for node in area['children']:
                groups = node['children'] if node.get('type') == 'subgroup' else [node]
                sg = node.get('label') if node.get('type') == 'subgroup' else None
                for g in groups:
                    if g.get('type') != 'group':
                        continue
                    dis = g['label']
                    # a) subgroup placement must match the registered subgroup
                    if ta in SUBGROUP_ORDER and dis != 'Misc.':
                        want = subgroup_for(ta, dis)
                        if sg is not None and want and sg != want and not (want == 'Other' and sg == 'Other'):
                            warns.append(f'[{wl}] {ta}: "{dis}" placed in subgroup "{sg}" but maps to "{want}"')
                    # b) the resolved disease should actually appear in the story text
                    for s in g['children']:
                        if s.get('type') != 'story' or dis in ('Misc.',):
                            continue
                        txt = story_text(s)
                        # re-derive from the headline's stated indication; flag a mismatch
                        title = plain(s.get('title', ''))
                        ph = indication_phrase(title)
                        if ph:
                            stated = infer_disease(ta, ph)
                            if stated:
                                stated = refine_generic(ta, stated, txt) or stated
                            if stated and stated != dis:
                                warns.append(f'[{wl}] {ta}/{dis}: headline says "{ph}" -> {stated}; '
                                             f'filed under {dis}: {title[:60]}')
    if warns:
        sys.stderr.write('\n[validate] %d classification warning(s):\n' % len(warns))
        for ln in warns[:40]:
            sys.stderr.write('  ' + ln + '\n')
        if len(warns) > 40:
            sys.stderr.write('  ... and %d more\n' % (len(warns) - 40))
    else:
        sys.stderr.write('[validate] no classification warnings\n')
    return warns

def build(paths):
    weeks = [parse_week(p, week_meta(p)) for p in paths]
    weeks.sort(key=lambda w: w['meta']['sort'])
    # cluster each disease area's indications into logical subgroups (GI, Thoracic…)
    for w in weeks:
        for area in w['categories'].get('disease', []):
            apply_subgroups(area)
    validate(weeks)
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
