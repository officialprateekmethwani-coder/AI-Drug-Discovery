#!/usr/bin/env python3
"""
Generate PowerPoint presentations for AI for Drug Discovery course.
Restructured: 4 teaching days + exam day.
  Day 1 (15.04.2026): 6 units (270 min) - Introduction, pipeline, AI overview, projects, RDKit practical
  Day 2 (22.04.2026): 8 units (360 min) - Molecular ML: SMILES, fingerprints, QSAR, evaluation, SHAP
  Day 3 (05.05.2026): 8 units (360 min) - Deep learning, GNNs, generative AI, protein targets/AlphaFold
  Day 4 (19.05.2026): 4 units (180 min) - Ethics, physiology as drug readout, project finalization
  Exam  (27.05.2026): Presentations / Posters

Run: python generate_slides.py
Requirements: pip install python-pptx
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

NAVY = RGBColor(23, 55, 94)
TEAL = RGBColor(31, 119, 180)
DARK_SLATE = RGBColor(44, 62, 80)
WHITE = RGBColor(255, 255, 255)
LIGHT_GRAY = RGBColor(240, 240, 240)
BODY_TEXT = RGBColor(50, 50, 50)
ACCENT_GREEN = RGBColor(39, 174, 96)
ACCENT_RED = RGBColor(231, 76, 60)
ACCENT_ORANGE = RGBColor(243, 156, 18)
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

def new_prs():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    return prs

def _add_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def _add_textbox(slide, left, top, width, height, text, font_size=18, bold=False, color=BODY_TEXT, alignment=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Calibri"
    p.alignment = alignment
    return tf

def _add_slide_number(slide, num):
    _add_textbox(slide, Inches(12.3), Inches(7.0), Inches(0.8), Inches(0.4), str(num), font_size=10, color=RGBColor(150,150,150), alignment=PP_ALIGN.RIGHT)

def _add_bullets(tf, bullets, font_size=18, color=BODY_TEXT):
    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = bullet
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.space_after = Pt(8)

def _set_notes(slide, notes_text):
    slide.notes_slide.notes_text_frame.text = notes_text

def make_title_slide(prs, title, subtitle, day_num, course_name, notes, date_str=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, NAVY)
    label = f"DAY {day_num}" + (f" — {date_str}" if date_str else "")
    _add_textbox(slide, Inches(0.8), Inches(0.6), Inches(5), Inches(0.6), label, font_size=16, bold=True, color=TEAL)
    _add_textbox(slide, Inches(0.8), Inches(1.8), Inches(11), Inches(2), title, font_size=44, bold=True, color=WHITE)
    _add_textbox(slide, Inches(0.8), Inches(4.0), Inches(11), Inches(1), subtitle, font_size=22, color=LIGHT_GRAY)
    _add_textbox(slide, Inches(0.8), Inches(6.2), Inches(5), Inches(0.5), course_name, font_size=14, color=TEAL)
    _set_notes(slide, notes)

def make_section_divider(prs, section_title, slide_num, notes, subtitle=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, TEAL)
    _add_textbox(slide, Inches(1), Inches(2.5), Inches(11), Inches(1.5), section_title, font_size=40, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
    if subtitle:
        _add_textbox(slide, Inches(1), Inches(4.2), Inches(11), Inches(1), subtitle, font_size=20, color=WHITE, alignment=PP_ALIGN.CENTER)
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)

def make_content_slide(prs, title, bullets, slide_num, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, Inches(1.1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    _add_textbox(slide, Inches(0.6), Inches(0.15), Inches(12), Inches(0.8), title, font_size=28, bold=True, color=WHITE)
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.5), Inches(5.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    _add_bullets(tf, bullets, font_size=20, color=BODY_TEXT)
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)

def make_two_column_slide(prs, title, left_title, left_bullets, right_title, right_bullets, slide_num, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, Inches(1.1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    _add_textbox(slide, Inches(0.6), Inches(0.15), Inches(12), Inches(0.8), title, font_size=28, bold=True, color=WHITE)
    _add_textbox(slide, Inches(0.8), Inches(1.3), Inches(5.5), Inches(0.5), left_title, font_size=22, bold=True, color=TEAL)
    txL = slide.shapes.add_textbox(Inches(0.8), Inches(1.9), Inches(5.5), Inches(5))
    tfL = txL.text_frame; tfL.word_wrap = True
    _add_bullets(tfL, left_bullets, font_size=18)
    _add_textbox(slide, Inches(7.0), Inches(1.3), Inches(5.5), Inches(0.5), right_title, font_size=22, bold=True, color=TEAL)
    txR = slide.shapes.add_textbox(Inches(7.0), Inches(1.9), Inches(5.5), Inches(5))
    tfR = txR.text_frame; tfR.word_wrap = True
    _add_bullets(tfR, right_bullets, font_size=18)
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)

def make_discussion_slide(prs, question, context_bullets, slide_num, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, DARK_SLATE)
    _add_textbox(slide, Inches(1), Inches(0.5), Inches(11), Inches(0.6), "DISCUSSION QUESTION", font_size=18, bold=True, color=TEAL)
    _add_textbox(slide, Inches(1), Inches(1.5), Inches(11), Inches(2), question, font_size=30, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
    if context_bullets:
        txBox = slide.shapes.add_textbox(Inches(1.5), Inches(4.0), Inches(10), Inches(3))
        tf = txBox.text_frame; tf.word_wrap = True
        _add_bullets(tf, context_bullets, font_size=18, color=LIGHT_GRAY)
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)

def make_takeaway_slide(prs, takeaways, day_num, slide_num, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, NAVY)
    _add_textbox(slide, Inches(1), Inches(0.5), Inches(11), Inches(0.8), f"KEY TAKEAWAYS — DAY {day_num}", font_size=28, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
    txBox = slide.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10), Inches(5))
    tf = txBox.text_frame; tf.word_wrap = True
    _add_bullets(tf, takeaways, font_size=22, color=WHITE)
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)

def make_story_slide(prs, title, story_text, slide_num, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, DARK_SLATE)
    _add_textbox(slide, Inches(1), Inches(0.5), Inches(11), Inches(0.7), title, font_size=26, bold=True, color=ACCENT_GREEN)
    _add_textbox(slide, Inches(1.2), Inches(1.6), Inches(10.5), Inches(5.2), story_text, font_size=20, color=WHITE)
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)


def make_quiz_question_slide(prs, q_number, question, options, slide_num, notes):
    """Quiz question slide: question + 4 options labeled A-D."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, NAVY)
    _add_textbox(slide, Inches(0.8), Inches(0.4), Inches(3), Inches(0.5),
                 f"QUIZ — Question {q_number}/10", font_size=16, bold=True, color=TEAL)
    _add_textbox(slide, Inches(0.8), Inches(1.1), Inches(11.5), Inches(1.4),
                 question, font_size=28, bold=True, color=WHITE)
    labels = ["A", "B", "C", "D"]
    y_start = 3.1
    for i, opt in enumerate(options):
        # option box
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(1.5), Inches(y_start + i * 1.05),
                                      Inches(10), Inches(0.85))
        box.fill.solid()
        box.fill.fore_color.rgb = DARK_SLATE
        box.line.color.rgb = TEAL
        box.line.width = Pt(2)
        tf = box.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.LEFT
        p = tf.paragraphs[0]
        p.text = f"  {labels[i]})  {opt}"
        p.font.size = Pt(22)
        p.font.color.rgb = WHITE
        p.font.name = "Calibri"
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)


def make_quiz_answer_slide(prs, q_number, question, options, correct_idx, explanation, slide_num, notes):
    """Quiz answer reveal: correct answer highlighted green, others gray."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, NAVY)
    _add_textbox(slide, Inches(0.8), Inches(0.3), Inches(3), Inches(0.4),
                 f"ANSWER — Question {q_number}/10", font_size=16, bold=True, color=ACCENT_GREEN)
    _add_textbox(slide, Inches(0.8), Inches(0.8), Inches(11.5), Inches(0.9),
                 question, font_size=22, bold=True, color=LIGHT_GRAY)
    labels = ["A", "B", "C", "D"]
    y_start = 2.0
    for i, opt in enumerate(options):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(1.5), Inches(y_start + i * 0.75),
                                      Inches(10), Inches(0.65))
        if i == correct_idx:
            box.fill.solid()
            box.fill.fore_color.rgb = ACCENT_GREEN
            box.line.color.rgb = ACCENT_GREEN
            txt_color = WHITE
        else:
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(60, 60, 70)
            box.line.color.rgb = RGBColor(80, 80, 90)
            txt_color = RGBColor(140, 140, 140)
        box.line.width = Pt(2)
        tf = box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"  {labels[i]})  {opt}"
        p.font.size = Pt(18)
        p.font.color.rgb = txt_color
        p.font.name = "Calibri"
    # Explanation box
    _add_textbox(slide, Inches(1.0), Inches(5.15), Inches(11), Inches(2.0),
                 explanation, font_size=16, color=WHITE)
    _add_slide_number(slide, slide_num)
    _set_notes(slide, notes)


# =====================================================================
# DAY 1 — 15 April 2026 — 6 units (270 min)
# Introduction, Drug Discovery Pipeline, AI Overview, Project Proposals
# =====================================================================
def generate_day1():
    prs = new_prs()
    n = 0

    # --- Slide 1: Title ---
    n += 1
    make_title_slide(prs,
        "AI for Drug Discovery",
        "From molecules to medicine — how AI is rewriting the rules",
        1, "AI for Drug Discovery", date_str="15 April 2026",
        notes="TIMING: 0-15 min (Unit 1 start)\n\nWelcome everyone! Start with a round of introductions — name, background, what interests you about AI and/or drug discovery. This is essential since we only have 4 teaching days together.\n\n=== ABOUT YOUR INSTRUCTOR ===\nDr. Étienne Serbe-Kamp is a neuroscientist whose research focuses on physiology and function of neural circuits in the Drosophila (fruit fly) visual system. Key publications:\n- Ammer, Serbe-Kamp et al. (2023) 'Multilevel visual motion opponency in Drosophila.' Nature Neuroscience 26:1894-1905 — showed that GluClα-mediated inhibition underlies direction-opponent responses at all levels of the motion circuit\n- Serbe-Kamp et al. (2023) 'Voltage to Calcium Transformation Enhances Direction Selectivity in Drosophila T4 Neurons.' J. Neurosci. 43:2497-2514 — demonstrated a nonlinear voltage-to-calcium transformation that enhances direction selectivity\n- Serbe et al. (2016) 'Comprehensive Characterization of the Major Presynaptic Elements to the Drosophila OFF Motion Detector.' Neuron 89:829-841 — this is a PHYSIOLOGY paper: used calcium imaging (GCaMP), whole-cell patch-clamp electrophysiology, and optogenetics (CsChrimson, GtACR1) to functionally characterize the neurons providing input to T5 motion detectors\n- Maisak, Haag, Ammer, Serbe et al. (2013) 'A Directional Tuning Map of Drosophila Elementary Motion Detectors.' Nature 500:212-216\n\nIMPORTANT: The 2016 Neuron paper is about PHYSIOLOGY of cells — measuring their functional responses using calcium imaging, whole-cell patch clamp, and optogenetics. Dr. Serbe-Kamp is currently working on connectomics with the MESH repository (separate project).\n\nHe is also Co-Director of Summer Fellowships at Backyard Brains, developing open-source neuroscience tools (SpikerBox/SpikerBot). Recent work:\n- Madariaga, [...], Serbe-Kamp, Marzullo (2024) 'A library of electrophysiological responses in plants.' Plant Signaling & Behavior 19(1):2310977\n- Serbe-Kamp et al. (2023) 'Open Citizen Science: fostering open knowledge with participation.' Research Ideas and Outcomes 9:e96476\n\nWhy a neuroscientist teaching drug discovery? The tools are the same! The pipeline (signal → features → model → prediction) is identical whether you study neural circuits, molecular properties, or physiological drug effects. And many important drug targets are ion channels and receptors that Dr. Serbe-Kamp studies.")

    # --- Slide 2: Course Overview & Schedule ---
    n += 1
    make_content_slide(prs,
        "Course Overview — 4 Days + Exam",
        ["Day 1 (TODAY, 15 Apr): Introduction, Drug Discovery Pipeline, AI, Projects — 6 units",
         "Day 2 (22 Apr): Molecular Representation, QSAR, ML, Evaluation, SHAP — 8 units",
         "Day 3 (5 May): Deep Learning, GNNs, Generative AI, AlphaFold — 8 units",
         "Day 4 (19 May): Ethics, Physiology as Drug Readout, Project Workshop — 4 units",
         "EXAM (27 May): Group Presentations / Posters",
         "",
         "Every day has: lecture + hands-on coding + discussion",
         "Assessment: Implementation 40% | Scientific reasoning 25% | Evaluation 20% | Presentation 15%",
         "",
         "Project sign-up TODAY — groups of 3-4 students"],
        n,
        notes="TIMING: 15-25 min\n\nWalk through the schedule. Emphasize: this is a compact, intensive format. Every day has a practical coding component. Students must form project groups TODAY and choose or propose a project topic. The exam on 27 May will be presentations (15-20 min per group) or posters (if collecting own data). Code will be in Python/Jupyter on Google Colab. No prior chemistry needed.")

    # --- Slide 3: The Brutal Reality ---
    n += 1
    make_content_slide(prs,
        "The Brutal Reality of Drug Discovery",
        ["Average time to market: 12-15 years",
         "Average cost: $2.6 billion per approved drug (DiMasi et al., 2016)",
         "Success rate from Phase I to approval: ~7.9%",
         "Over 90% of drug candidates fail in clinical trials",
         "",
         "But here's the key insight for this course:",
         "  The LAST readout of a drug is always PHYSIOLOGY",
         "  → Does it change cell function? Organ function? Patient health?",
         "  → Electrophysiology, calcium imaging, behavioral assays",
         "  → The same techniques used in neuroscience research on ion channels!",
         "",
         "\"The most expensive experiment in science is a failed clinical trial.\""],
        n,
        notes="TIMING: 25-35 min\n\nCITATION CHECK: DiMasi, Grabowski & Hansen (2016) 'Innovation in the pharmaceutical industry: New estimates of R&D costs.' J. Health Economics 47:20-33. The $2.6B figure is the commonly rounded version of their $2.558B out-of-pocket estimate ($2.87B capitalized, in 2013 dollars). The ~7.9% Phase I-to-approval success rate comes from BIO/Informa Pharma Intelligence/QLS Advisors (2021) 'Clinical Development Success Rates 2011-2020.'\n\nLet the numbers sink in. Ask: does anyone know how much a Phase III trial costs? ($50-100M+). Point out the key theme: no matter how good your AI model is at predicting molecular properties, the FINAL validation is always physiological — does the drug actually change cell/tissue/organ function?\n\nThis is why physiology matters for drug design: electrophysiology (patch-clamp, extracellular recording) is the gold standard for ion channel drugs. Calcium imaging measures receptor activation. These are EXACTLY the techniques in Dr. Serbe-Kamp's 2016 Neuron paper (Serbe et al., Neuron 89:829-841) and 2023 J.Neurosci. paper (Serbe-Kamp et al., J. Neurosci. 43:2497-2514).\n\nThe Backyard Brains SpikerBox/SpikerBot offer a DIY version of these physiological readouts — making the final step of drug testing accessible. SpikerBox ref: Marzullo & Gage (2012) PLoS ONE 7(3):e30837.")

    # --- Slide 4: Traditional Pipeline ---
    n += 1
    make_content_slide(prs,
        "The Drug Discovery Pipeline",
        ["1. TARGET IDENTIFICATION — Find a biological target (protein, ion channel, receptor)",
         "2. TARGET VALIDATION — Confirm relevance to disease",
         "3. HIT DISCOVERY — Screen millions of compounds (HTS or virtual screening)",
         "4. LEAD OPTIMIZATION — Improve potency, selectivity, ADMET properties",
         "5. PRECLINICAL — Animal studies for safety and efficacy",
         "6. PHASE I — Safety in healthy volunteers (20-100 people)",
         "7. PHASE II — Efficacy in patients (100-300 people)",
         "8. PHASE III — Large-scale trials (1,000-3,000+ people)",
         "9. FDA/EMA REVIEW & APPROVAL",
         "10. POST-MARKET SURVEILLANCE (Phase IV)",
         "",
         "At EVERY stage: physiology is the readout (cell assays → animal → human)"],
        n,
        notes="TIMING: 35-45 min\n\nWalk through each step in 5 min. Emphasize: HTS screens 1-2 million compounds physically — expensive. Lead optimization is where AI has huge potential (predict ADMET = Absorption, Distribution, Metabolism, Excretion, Toxicity). Chemical space is ~10^60 drug-like molecules.\n\nKEY POINT: At every step, the validation is physiological. Target validation uses cell assays (electrophysiology, calcium imaging). Hit discovery uses biochemical/cell assays. Preclinical uses whole-animal physiology. Clinical trials measure patient physiology. The SpikerBox/SpikerBot let you do step 5's physiology on a tabletop.")

    # --- Slide 5: Where AI Fits ---
    n += 1
    make_content_slide(prs,
        "Where AI Fits in the Pipeline",
        ["Target Identification — NLP on literature, network biology, multi-omics",
         "Virtual Screening — Score millions of compounds in silico (vs. months in HTS)",
         "Lead Optimization — Predict ADMET, suggest modifications",
         "De novo Drug Design — Generate entirely new molecules with desired properties",
         "Clinical Trial Optimization — Patient stratification, endpoint prediction",
         "Repurposing — Find new uses for existing drugs",
         "",
         "AI can compress timelines by years and reduce costs significantly",
         "  McKinsey/MGI (2020): 'The Bio Revolution' report + 2022 analyses",
         "",
         "Case study: Insilico Medicine ISM001-055 (Rentosertib)",
         "  First AI-designed drug to reach Phase IIa (enrolled Apr 2023)",
         "  Target to Phase II in ~30 months (vs. 4-5 years typical)"],
        n,
        notes="TIMING: 45-55 min (end of Unit 1)\n\nCITATION: The McKinsey figures come from McKinsey Global Institute 'The Bio Revolution' (May 2020) and the follow-up 'How AI could revolutionize drug discovery' (McKinsey, 2022). The exact '2-4 years' and '30-50%' numbers are synthesized from these analyses — they are estimates, not precise measurements. Be honest about this with students.\n\nCITATION: Insilico Medicine ISM001-055 (Rentosertib). Published as: Ren et al. (2024) 'A small-molecule TNIK inhibitor targets fibrosis in preclinical and clinical models.' Nature Biotechnology. doi:10.1038/s41587-024-02143-0. The drug is a first-in-class TNIK inhibitor for idiopathic pulmonary fibrosis (IPF). Project started ~late 2020, reached Phase IIa enrollment in April 2023 (~30 months). Phase IIa results (71 patients, NCT05938920) were positive, announced September 2024. NOTE: Published in Nature Biotechnology, NOT Chemical Science.\n\nGo through each bullet with examples. The Insilico case is landmark: AI identified a novel target (TNIK) AND designed the molecule. End-to-end AI-driven discovery.\n\nNeuroscience connection: AI on the Drosophila connectome identifies new neural targets. Virtual screening for BBB-penetrant CNS drugs. AI predicts ion channel selectivity. BenevolentAI identified baricitinib for COVID-19.\n\nTransition: Let's look at the most famous AI success story in biology...")

    # --- Slide 6: AlphaFold ---
    n += 1
    make_content_slide(prs,
        "AlphaFold: The Protein Folding Revolution",
        ["Protein structure prediction — a 50-year grand challenge",
         "CASP14 (2020): AlphaFold2 achieved experimental-level accuracy",
         "Jumper et al. (2021), Nature — 2024 Nobel Prize in Chemistry",
         "",
         "Impact on drug discovery & neuroscience:",
         "  200+ million protein structures predicted (AlphaFold DB)",
         "  Ion channels, GPCRs, neurotransmitter receptors now have structures",
         "  Structure-based drug design for CNS targets (hERG, GABA-A, GluCl)",
         "",
         "AlphaFold3 (2024): predicts protein-ligand complexes",
         "",
         "Student project option: reproduce AlphaFold figures for a drug target!"],
        n,
        notes="TIMING: 55-65 min (Unit 2 start)\n\nCITATION: Jumper et al. (2021) 'Highly accurate protein structure prediction with AlphaFold.' Nature 596:583-589. doi:10.1038/s41586-021-03819-2. Won the 2024 Nobel Prize in Chemistry for Demis Hassabis and John Jumper (shared with David Baker for computational protein design).\n\nCITATION: AlphaFold3: Abramson et al. (2024) 'Accurate structure prediction of biomolecular interactions with AlphaFold 3.' Nature 630:493-500. doi:10.1038/s41586-024-07487-w. Extends to protein-ligand, protein-DNA, protein-RNA complexes.\n\nNeuroscience connection: GluCl (from Dr. Serbe-Kamp's Drosophila research, Cys-loop receptor superfamily) can be modeled. GABA-A receptors, GPCRs (serotonin 5-HT2A, dopamine D2 receptors), all benefit. About 34% of all FDA-approved drugs target GPCRs (Saikia et al. 2019, Curr. Drug Targets 20:522-539).\n\nStudent project: pick a neuroscience drug target, retrieve its AlphaFold structure, analyze binding sites, and compare to experimental structures. They can reproduce key figures from the Jumper et al. paper or apply it to novel targets.")

    # --- Slide 7: David Willson mRNA story ---
    n += 1
    make_story_slide(prs,
        "The DIY Medicine Story: David Willson's mRNA Vaccine for His Dog",
        "In 2024, David Willson's dog was diagnosed with cancer.\n\nHe worked with his veterinarian to design a personalized mRNA vaccine targeting his dog's specific tumor antigens — the same technology behind COVID-19 vaccines.\n\nThe approach: sequence the tumor → identify neoantigens → design mRNA → train the immune system.\n\nThe story went viral. Key question: is this the future of personalized medicine?\n\nNOTE: This is a widely reported news story, not a peer-reviewed publication. But the underlying science (mRNA cancer vaccines, neoantigen prediction) IS well-published.\n\nPossible student project: present the science behind personalized mRNA cancer vaccines and the AI pipeline that powers neoantigen prediction.",
        n,
        notes="TIMING: 65-75 min\n\nSpend 5 min. This is a hook story. The David Willson mRNA vaccine for his dog went viral in 2024 (covered by Wired, BBC, etc.).\n\nIMPORTANT: This is NOT a peer-reviewed publication — it's a news story / case report. The underlying science of mRNA cancer vaccines IS published (BioNTech/Moderna clinical trials, neoantigen prediction papers). Students could present this as a TOPIC — explaining the AI pipeline behind neoantigen prediction, mRNA optimization, and immune response modeling — but they would NOT be reproducing specific published data.\n\nAs a student project option: present the science behind personalized mRNA vaccines, explain the AI components, discuss ethics and regulation. Best suited for a PRESENTATION rather than a poster with own data.\n\nThe broader point: the tools for personalized medicine are becoming accessible. AI accelerates every step.")

    # --- Slide 8: Molecular Data Types ---
    n += 1
    make_section_divider(prs, "Molecular Data Types", n,
        notes="TIMING: 75-80 min\n\nTransition: Now let's get technical. How do we represent molecules for computers?",
        subtitle="How do we represent molecules for computers?")

    # --- Slide 9: SMILES ---
    n += 1
    make_content_slide(prs,
        "SMILES: Molecular-Input Line-Entry System",
        ["A way to write molecular structures as text strings",
         "Weininger (1988), J. Chem. Inf. Comput. Sci. 28:31-36",
         "",
         "Examples (including neuroscience drugs):",
         "  Water: O  |  Ethanol: CCO  |  Aspirin: CC(=O)Oc1ccccc1C(=O)O",
         "  Serotonin: NCCc1c[nH]c2ccc(O)cc12  (neurotransmitter)",
         "  Dopamine: NCCc1ccc(O)c(O)c1  (reward pathway)",
         "  GABA: NCCCC(=O)O  (main inhibitory neurotransmitter)",
         "  Diazepam (Valium): ClC1=CC2=C(C=C1)N(C)C(=O)CN2  (GABA-A modulator)",
         "  Caffeine: Cn1c(=O)c2c(ncn2C)n(C)c1=O  (adenosine antagonist)",
         "",
         "Why SMILES matters: compact, human-readable, machine-parseable"],
        n,
        notes="TIMING: 80-90 min\n\nSpend 5-7 min on SMILES. Cover basic syntax: atoms as letters, implicit single bonds, = for double, lowercase for aromatic, parentheses for branches, digits for rings.\n\nNeuroscience drug SMILES:\n1. SEROTONIN: indolamine, 14 receptor subtypes, targeted by SSRIs\n2. DOPAMINE: catecholamine, reward/movement, targeted by L-DOPA, antipsychotics\n3. GABA: simplest structure, main inhibitory NT, GABA-A is same superfamily as GluCl!\n4. DIAZEPAM: benzodiazepine, positive allosteric modulator of GABA-A\n5. CAFFEINE: xanthine, adenosine receptor antagonist\n\nAsk: can you identify the hydroxyl group (O) and amine (N) in the serotonin SMILES?")

    # --- Slide 10: Bioactivity ---
    n += 1
    make_content_slide(prs,
        "Bioactivity & Physiology: Measuring Drug Effects",
        ["How well does a molecule interact with its target?",
         "",
         "Key measures:",
         "  IC50 — concentration inhibiting 50% of target  |  pIC50 = -log10(IC50)",
         "  Ki — binding affinity  |  EC50 — half-maximal effective concentration",
         "",
         "THE PHYSIOLOGICAL READOUT (the final test):",
         "  Patch-clamp electrophysiology → IC50 for ion channel block (hERG, Nav, GluCl)",
         "  Calcium imaging → EC50 for receptor activation (GPCRs, GluCl)",
         "  Extracellular recording → drug effects on neural firing patterns",
         "  SpikerBox/SpikerBot → accessible version of these same measurements!",
         "",
         "Databases: ChEMBL (~2.5M compounds), PubChem (>110M), BindingDB",
         "This is where physiology expertise meets drug discovery"],
        n,
        notes="TIMING: 90-100 min\n\nSpend 5-7 min. IC50 analogy: how much drug to shut down 50% of the target? Lower = more potent.\n\nCITATION: ChEMBL: Gaulton et al. (2017) 'The ChEMBL database in 2017.' Nucleic Acids Res. 45(D1):D945-D954. Latest release (ChEMBL 35, Dec 2024) contains ~2.5M compounds.\n\n=== PHYSIOLOGY AS DRUG READOUT — KEY THEME ===\nThis is the central connection between Dr. Serbe-Kamp's research and drug discovery:\n\n1. PATCH-CLAMP ELECTROPHYSIOLOGY: Gold standard for ion channel drugs. Measure current through channels, add drug, measure IC50. This is exactly what Dr. Serbe-Kamp's 2016 Neuron paper used — whole-cell recordings + optogenetics to characterize T5 neuron inputs (Serbe et al. 2016, Neuron 89:829-841). For hERG cardiac safety testing: patch-clamp HEK293 cells, measure block IC50.\n\n2. CALCIUM IMAGING: Dr. Serbe-Kamp's 2023 J.Neurosci. paper (43:2497-2514) used two-photon calcium imaging in Drosophila T4 neurons. Same technique used in drug screening (FLIPR assays for GPCR screens).\n\n3. EXTRACELLULAR RECORDING: The SpikerBox records extracellular spikes from neurons/muscles. Apply a drug (e.g., lidocaine = Nav blocker) and watch the action potentials disappear. This IS drug effect measurement! Ref: Marzullo & Gage (2012) PLoS ONE 7(3):e30837.\n\nThe SpikerBox/SpikerBot democratize the final step of drug testing: measuring physiological effects.")

    # --- Slide 11: Physiology as Drug Readout ---
    n += 1
    make_content_slide(prs,
        "Physiology: The Final Readout of Drug Action",
        ["No matter how good your AI model, the drug must change PHYSIOLOGY:",
         "",
         "The hierarchy of drug readout:",
         "  1. In silico: AI predicts binding affinity (QSAR, docking, GNN)",
         "  2. In vitro: biochemical assay confirms binding (Ki, IC50)",
         "  3. Cellular: electrophysiology / calcium imaging shows functional effect",
         "  4. Tissue/organ: physiological response (cardiac QT, neural firing)",
         "  5. Organism: behavioral / clinical outcome",
         "",
         "Serbe et al. (2016) Neuron 89:829-841: steps 3-4 in Drosophila",
         "  → Calcium imaging + electrophysiology + optogenetics",
         "  → Characterizing physiological responses of neurons to stimuli",
         "",
         "DIY drug readout: BYB SpikerBox measures steps 3-4 on your desk!",
         "This perspective connects AI predictions to biological reality"],
        n,
        notes="TIMING: 100-110 min\n\n=== KEY SLIDE — PHYSIOLOGY AS DRUG DESIGN APPROACH ===\n\nThis is a critical conceptual slide that frames the entire course. Spend 5-7 min.\n\nThe drug discovery pipeline generates predictions (AI models, virtual screening). But predictions are HYPOTHESES. The test is always physiological.\n\nDr. Serbe-Kamp's 2016 Neuron paper ('Comprehensive Characterization of the Major Presynaptic Elements to the Drosophila OFF Motion Detector', Serbe et al., 2016, Neuron 89:829-841) is a physiology paper:\n- Used calcium imaging (GCaMP) to measure visual responses\n- Used whole-cell electrophysiology to record from neurons\n- Used optogenetics (CsChrimson, GtACR) to activate/silence specific neurons\n- This is NOT a connectomics paper — it's about FUNCTIONAL characterization of neural circuits\n- Dr. Serbe-Kamp is currently working on connectomics with the MESH repository (separate project)\n\nThe BYB gear connection:\n- SpikerBox = extracellular recording (like steps 3-4 in the hierarchy)\n- You can measure drug effects: apply lidocaine to a cockroach leg preparation, watch action potentials disappear\n- This IS the physiological readout that validates all the AI predictions\n- Making this accessible = democratizing the final step of drug testing\n\nFor drug design approach: start from the desired physiological endpoint (e.g., block a specific ion channel, change neural firing pattern) and work backwards. AI helps predict which molecules will achieve that endpoint. But the physiological measurement is what proves it works.")

    # --- Slide 12: Section divider ---
    n += 1
    make_section_divider(prs, "Biophysics & The SpikerBot", n,
        notes="TIMING: 110-112 min\n\nTransition: Let's see how neuroscience tools connect to drug discovery.",
        subtitle="Bridging neuroscience tools and drug discovery")

    # --- Slide 13: SpikerBot & Biophysics ---
    n += 1
    make_content_slide(prs,
        "Biophysical Modelling & The SpikerBot",
        ["The SpikerBot (Backyard Brains) — record AND stimulate real neurons",
         "  BYB Summer Fellowships — citizen neuroscience at Backyard Brains",
         "",
         "Ion channels are important drug targets (Santos et al. 2017: ~6% of efficacy targets):",
         "  Nav (antiepileptics), GABA-A (anxiolytics), nAChR (addiction), hERG (safety)",
         "  GluCl channels (Cys-loop receptor family, studied in Drosophila) = ivermectin target!",
         "",
         "Hodgkin-Huxley model (1952; Nobel Prize in Physiology/Medicine 1963):",
         "  Simulate how drugs alter ion channel gating → predict physiological effects",
         "  We simulate this in the Day 1 notebook!",
         "",
         "The computational SpikerBox: simulate neural signals computationally,",
         "  then validate with real recordings — bridging AI and physiology"],
        n,
        notes="TIMING: 112-125 min (Unit 3 start)\n\nSpend 7-8 min. Cover the SpikerBot hardware, ion channel pharmacology, and the Hodgkin-Huxley model.\n\n=== DR. SERBE-KAMP'S RESEARCH ===\nKey publications (PHYSIOLOGY focus):\n1. Maisak, Haag, Ammer, Serbe et al. (2013, Nature 500:212-216) — directional tuning map of T4/T5 neurons. The first paper to show that T4 and T5 cells are the elementary motion detectors of the fly visual system.\n2. Serbe et al. (2016, Neuron 89:829-841) — PHYSIOLOGY paper: calcium imaging (GCaMP6f, two-photon), whole-cell patch-clamp electrophysiology, and optogenetics (CsChrimson, GtACR1) to characterize T5 motion detector inputs. This is NOT connectomics.\n3. Ammer, Serbe-Kamp et al. (2023, Nature Neurosci. 26:1894-1905) — GluClα mediates motion-opponent inhibition at all levels of the Drosophila motion pathway\n4. Serbe-Kamp et al. (2023, J. Neurosci. 43:2497-2514) — demonstrated that nonlinear voltage-to-calcium transformation in T4 neurons sharpens directional tuning\n\nThe 2016 Neuron paper is about PHYSIOLOGY: functional characterization using calcium imaging, whole-cell recordings, and optogenetics. NOT connectomics. Dr. Serbe-Kamp currently works on connectomics with MESH.\n\n=== GluCl / IVERMECTIN CONNECTION ===\nGluCl is a glutamate-gated chloride channel in the Cys-loop receptor superfamily (same family as GABA-A, glycine receptors, nAChR). Ivermectin locks GluCl open → Cl⁻ influx → hyperpolarization → paralysis in parasites. 2015 Nobel Prize in Physiology or Medicine (Satoshi Ōmura & William C. Campbell) for the discovery of avermectin/ivermectin. Dr. Serbe-Kamp's 2023 Nat.Neurosci. paper (Ammer, Serbe-Kamp et al., 26:1894-1905) found GluClα mediates inhibition in Drosophila motion circuits — same channel family, different organism, different function.\n\n=== HODGKIN-HUXLEY MODEL ===\nCITATION: Hodgkin & Huxley (1952) 'A quantitative description of membrane current and its application to conduction and excitation in nerve.' J. Physiol. 117:500-544. Won the 1963 Nobel Prize in Physiology or Medicine (shared with John Eccles).\nEquation: C_m * dV/dt = -g_Na*m³h*(V-E_Na) - g_K*n⁴*(V-E_K) - g_L*(V-E_L) + I_ext\nDrug effects = modifying conductance parameters (g_Na, g_K). For example:\n- Lidocaine (Nav blocker): reduce g_Na → fewer/no action potentials\n- 4-aminopyridine (Kv blocker): reduce g_K → broader action potentials\n- Tetrodotoxin (TTX, pufferfish toxin): completely blocks Nav → no firing\nWe simulate this in the Day 1 practical!\n\n=== SPIKERBOX ===\nCITATION: Marzullo & Gage (2012) 'The SpikerBox: A Low Cost, Open-Source BioAmplifier for Increasing Public Participation in Neuroscience Inquiry.' PLoS ONE 7(3):e30837. Note: published in PLoS ONE, not Advances in Physiology Education.")

    # --- Slide 14: Introduction to RDKit ---
    n += 1
    make_content_slide(prs,
        "RDKit: The Chemistry Toolkit",
        ["RDKit — open-source cheminformatics library (C++ core, Python bindings)",
         "Landrum, G. (2006-2024). https://www.rdkit.org",
         "",
         "What RDKit does:",
         "  Parse SMILES → molecular objects",
         "  Calculate properties (MW, LogP, HBD, HBA, TPSA)",
         "  Generate molecular fingerprints (Morgan/ECFP, MACCS)",
         "  Substructure searching and molecular similarity",
         "  2D and 3D visualization",
         "",
         "Installation: pip install rdkit",
         "We use RDKit throughout this entire course",
         "",
         "→ Hands-on practical coming right after the break!"],
        n,
        notes="TIMING: 125-130 min\n\n2 min quick overview. RDKit is the NumPy of chemistry. We'll use it immediately in the practical.")

    # --- Slide 15: BREAK ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="TIMING: 130-145 min\n\nBreak time! Students should stretch, get coffee, and start thinking about project ideas. Encourage informal conversations about which project they might want to do.",
        subtitle="Come back ready for the practical + project sign-up!")

    # --- Slide 16: Project Proposals ---
    n += 1
    make_two_column_slide(prs,
        "Student Projects — Choose or Propose Your Own!",
        "Hardware / Physiology Projects (Poster format)",
        ["1. Computational SpikerBox — simulate neural signals,",
         "   classify spike patterns with ML, model drug effects",
         "   on ion channels computationally",
         "2. SpikerBot — program stimulation patterns, record",
         "   physiological responses, model dose-response",
         "3. BYB Human Signals — record your own EMG/ECG/EEG,",
         "   apply ML classification, relate to pharmacology",
         "4. Eye-Tracking — visual processing + CNS drug effects",
         "   on attention/saccades, ML on gaze data",
         "5. VR Setup — VR-based neuro experiments, behavioral",
         "   data + ML, connect to drug effects on perception"],
        "Computational Projects (Presentation format)",
        ["6. Classical QSAR — build QSAR model for a neuro",
         "   drug target (GABA-A, 5-HT, dopamine D2, hERG)",
         "   from ChEMBL data (Day 2-3 methods)",
         "7. AlphaFold Structure — predict a neuro target,",
         "   analyze binding sites, reproduce key figures",
         "",
         "OWN IDEA? Propose it today!",
         "",
         "Format: Poster (own data) or Presentation",
         "  (reproduce/explain published figures)",
         "Groups of 3-4 | Exam: 27 May"],
        n,
        notes="TIMING: 145-170 min (25 min for this slide + discussion + sign-up)\n\n=== PROJECT DETAILS ===\n\nHARDWARE / PHYSIOLOGY PROJECTS (Poster format — collect your own data):\n\n1. COMPUTATIONAL SPIKERBOX:\n   - Use Hodgkin-Huxley simulation to generate synthetic neural signals\n   - Apply ML classification to spike patterns\n   - Model how ion channel drugs (Nav blockers, K+ modulators) change firing patterns\n   - Justify: directly connects to Day 1 content (ion channels, drug targets, physiology as readout)\n   - If real SpikerBox available: validate simulations against real recordings\n\n2. SPIKERBOT:\n   - Program stimulation patterns (tonic, burst, Poisson)\n   - Record muscle/nerve responses to different stimulation\n   - Model computationally: build dose-response curves from stimulation parameters\n   - Justify: connects to biophysical modeling, drug-target interactions, dose-response (Day 2)\n\n3. BYB HUMAN SIGNALS:\n   - Record EMG (muscle), ECG (heart), or EEG (brain) with BYB devices\n   - Apply ML classification (e.g., classify hand gestures from EMG, detect arrhythmia from ECG)\n   - Pharmacological connection: caffeine effects on EMG/ECG, stress hormones on heart rate\n   - Justify: connects to physiological readout theme, ML classification (Day 2), signal processing\n\n4. EYE-TRACKING:\n   - Record eye movements during visual tasks\n   - ML on gaze data: classify attention states, predict task performance\n   - CNS drug connection: stimulants affect saccade latency, sedatives reduce fixation accuracy\n   - Justify: connects to Drosophila visual system (Dr. Serbe-Kamp's research), CNS pharmacology, ML\n\n5. VR SETUP:\n   - VR-based neuroscience experiment (reaction time, spatial navigation, perception)\n   - Collect behavioral data, apply ML analysis\n   - Drug connection: how CNS drugs alter perception, reaction time, spatial cognition\n   - Justify: connects to neural processing, computational modeling, behavioral pharmacology\n\nCOMPUTATIONAL PROJECTS (Presentation format — reproduce or explain published work):\n\n6. CLASSICAL QSAR:\n   - Pick a neuroscience drug target from ChEMBL (e.g., GABA-A, serotonin 5-HT2A, dopamine D2, hERG)\n   - Build QSAR model using methods from Day 2-3\n   - Evaluate with scaffold splits, interpret with SHAP\n   - Justify: core course content, directly applicable\n\n7. ALPHAFOLD STRUCTURE:\n   - Pick a neuroscience drug target (ion channel, GPCR, transporter)\n   - Retrieve AlphaFold prediction, analyze binding sites\n   - Compare to experimental structure if available\n   - Reproduce key figures from Jumper et al. 2021 or related\n   - Justify: connects to protein structure, drug design, Day 3 content\n\n=== ON THE DAVID WILLSON mRNA STORY ===\nStudents COULD present on personalized mRNA cancer vaccines as a topic, explaining the AI pipeline. This would be a PRESENTATION (not poster). It's based on news reporting, not a single reproducible paper. Students would need to:\n- Explain the neoantigen prediction pipeline\n- Discuss mRNA vaccine technology (well-published from COVID-19 work)\n- Address ethics and regulation\n- This is feasible but they should focus on the published science, not the specific dog case\n\n=== SIGN-UP PROCESS ===\nHave a sign-up sheet ready. Let students discuss for 10 min, then form groups and sign up before the practical. They can change until Day 2.")

    # --- Slide 17: Project Details ---
    n += 1
    make_content_slide(prs,
        "What You'll Present on 27 May",
        ["POSTER (Projects 1-5: own data collection):",
         "  → Collect physiological/behavioral data using BYB gear / eye-tracker / VR",
         "  → Apply ML analysis from the course",
         "  → Present results as a scientific poster",
         "  → I will ask questions to assess understanding depth",
         "",
         "PRESENTATION (Projects 6-7 or publication-based):",
         "  → Reproduce figures from a published paper OR",
         "  → Explain how figures were generated (e.g., AlphaFold methods)",
         "  → 15-20 minutes per group + Q&A",
         "",
         "OWN IDEAS welcome! Propose by Day 2 (22 Apr)",
         "All publications will be available in the course repository",
         "",
         "Sign up NOW → groups of 3-4"],
        n,
        notes="TIMING: 170-175 min\n\nQuick 3-min clarification of the exam format. The key points:\n1. Both formats (poster and presentation) will involve Q&A where you assess understanding\n2. For posters: original data collection + analysis shows initiative\n3. For presentations: reproducing or explaining published figures shows scientific depth\n4. Own ideas are welcome but must be approved by Day 2\n5. Publications will be uploaded to the publications/ folder in the repository")

    # --- Slide 18: Practical intro ---
    n += 1
    make_content_slide(prs,
        "Practical: Explore Molecules with RDKit + Hodgkin-Huxley",
        ["Open: day1_introduction/Day1_Practical.ipynb in Google Colab",
         "",
         "Part 1 — RDKit Basics (~30 min):",
         "  1. Install and import RDKit",
         "  2. Load EGFR inhibitor dataset from ChEMBL",
         "  3. Parse SMILES and visualize molecules",
         "  4. Calculate molecular properties (MW, LogP, rings)",
         "  5. Plot property distributions",
         "",
         "Part 2 — Hodgkin-Huxley Simulation (~25 min):",
         "  6. Simulate a neuron's action potential",
         "  7. Model drug effects: change Nav/K channel conductance",
         "  8. Visualize how drugs alter firing patterns",
         "",
         "Time budget: ~55 min including code execution | Work in pairs"],
        n,
        notes="TIMING: 175-180 min (start practical)\n\nSpend 3 min explaining, then let students code. Walk around and help. The notebook is designed for Google Colab — they just upload and run.\n\nPart 1 uses EGFR inhibitors as the example dataset. Students may know gefitinib (Iressa), erlotinib (Tarceva).\n\nPart 2 is the physiology connection: Hodgkin-Huxley model simulates action potentials. By changing conductance parameters, students model drug effects (e.g., halving g_Na mimics a partial Nav blocker like lidocaine). This directly connects to the SpikerBox project options.\n\nBudget 55 min including code execution time. Code cells should run in <30 sec each on Colab.")

    # --- Slide 19: Practical session (hands-on) ---
    n += 1
    make_section_divider(prs, "🔬 HANDS-ON PRACTICAL", n,
        notes="TIMING: 180-235 min (Units 4-5)\n\n55 minutes of hands-on coding. Walk around the room. Help students who are stuck. Encourage discussions between pairs.",
        subtitle="Open your notebooks — let's code!")

    # --- Slide 20: Discussion ---
    n += 1
    make_discussion_slide(prs,
        "If you could measure ONE physiological effect\nof a drug, which technique would you choose and why?",
        ["Consider: electrophysiology (ion channels), calcium imaging (signaling),",
         "  EEG/EMG (whole-system), behavioral assays, imaging...",
         "Think about: sensitivity, specificity, throughput, cost",
         "Connect to: which of the 7 projects interests you most?"],
        n,
        notes="TIMING: 235-245 min (Unit 6 start)\n\n7-min discussion connecting the physiology theme to project choice. Guide toward: different techniques answer different questions. Electrophysiology gives molecular-level resolution (single ion channels), calcium imaging gives cellular resolution, EEG/EMG gives systems-level data. The SpikerBox bridges the gap between molecular and systems levels.")

    # --- Slide 21: Key Takeaways ---
    n += 1
    make_takeaway_slide(prs,
        ["Drug discovery: 12-15 years, $2.6B, >90% failure rate",
         "AI accelerates every stage — Insilico Medicine, AlphaFold",
         "Physiology is the LAST and MOST IMPORTANT readout of drugs",
         "  → Electrophysiology, calcium imaging = gold standard for drug validation",
         "  → SpikerBox/SpikerBot = accessible DIY drug readout",
         "SMILES converts molecules to text; bioactivity (IC50/Ki) comes from physiological assays",
         "GluCl channels: from Drosophila motion vision to ivermectin drug target",
         "7 project options proposed — sign up TODAY!",
         "Next (Day 2, 22 Apr): molecular fingerprints, QSAR, train your first ML model!"],
        1, n,
        notes="TIMING: 245-255 min\n\nQuick 5-min recap. Emphasize the three themes: 1) Drug discovery has a problem (slow/expensive/risky), 2) AI is the solution, 3) Physiology is the validation. The GluCl example perfectly bridges neuroscience and drug discovery.")

    # --- Slide 22: Next Day preview ---
    n += 1
    make_content_slide(prs,
        "Coming Up: Day 2 — 22 April 2026",
        ["Morning (4 units):",
         "  Molecular fingerprints (Morgan/ECFP) deep dive",
         "  QSAR: Quantitative Structure-Activity Relationships",
         "  Train Random Forest & XGBoost models",
         "  Practical: predict aqueous solubility",
         "",
         "Afternoon (4 units):",
         "  Model evaluation: scaffold splits, metrics, overfitting",
         "  SHAP: explain your model's predictions",
         "  Practical: re-evaluate and interpret your QSAR model",
         "",
         "Preparation:",
         "  Finish today's practical notebook",
         "  Confirm your project group and topic"],
        n,
        notes="TIMING: 255-260 min\n\nQuick 3-min preview. Day 2 is the most content-heavy day (8 units). Students will go from molecular representation to trained, evaluated, and interpreted ML models in one day. Remind them to finish the Day 1 notebook and confirm project groups.")

    # --- Slide 23: Thank You ---
    n += 1
    make_section_divider(prs, "Thank You! Questions?", n,
        notes="TIMING: 260-270 min (end)\n\nOpen floor for questions. Remind about project sign-up. See you on 22 April!",
        subtitle="Project sign-up sheet is open — form your groups!")

    os.makedirs("day1_introduction", exist_ok=True)
    prs.save("day1_introduction/slides.pptx")
    print(f"  Day 1: {n} slides → day1_introduction/slides.pptx")


# =====================================================================
# DAY 2 — 22 April 2026 — 8 units (360 min)
# Molecular Representation, QSAR, ML Baselines, Evaluation, SHAP
# =====================================================================
def generate_day2():
    prs = new_prs()
    n = 0

    # --- Title ---
    n += 1
    make_title_slide(prs,
        "Molecular ML: From Fingerprints to Predictions",
        "Represent molecules, train models, evaluate properly, and explain results",
        2, "AI for Drug Discovery", date_str="22 April 2026",
        notes="TIMING: 0-10 min (Unit 1 start)\n\nWelcome back! Quick check-in: how did the Day 1 practical go? Who finished the Hodgkin-Huxley simulation? Today is a BIG day — 8 units. By the end, you will have trained, evaluated, and interpreted ML models on real molecular data.\n\nAgenda:\nMorning (Units 1-4, ~180 min): SMILES deep dive, fingerprints, QSAR, RF/XGBoost, Practical 1\nAfternoon (Units 5-8, ~180 min): Evaluation, scaffold splits, SHAP, Practical 2")


    # =========================================================
    # PHYSIOLOGY INTRO (added for Day 2 — CS student audience)
    # =========================================================

    # --- Slide 2: Homework Reminder from Last Week ---
    n += 1
    make_content_slide(prs,
        "Homework from Last Week — Did You…?",
        ["📱 Download SpikeRecorder (backyardbrains.com/SpikeRecorder)",
         "  → Works on iOS, Android, and desktop",
         "  → We will use it to visualize and simulate neural signals today",
         "",
         "📄 Read the three papers shared after Day 1:",
         "  → If not — try to read before Day 3 (5 May)",
         "  → Key insight in each paper connects to today's ML methods",
         "",
         "🐍 RDKit exercise: plot all neurotransmitters as molecular structures",
         "  → import rdkit; Draw.MolsToGridImage([...]) for all NTs",
         "  → Neurotransmitters: serotonin, dopamine, GABA, glutamate, acetylcholine",
         "  →                    norepinephrine, glycine, histamine, adenosine",
         "",
         "💡 Have you started thinking about your project idea?",
         "  → Groups of 3-4, sign-up deadline: Day 2 end"],
        n,
        notes="TIMING: 5-10 min (before the physiology intro)\n\n=== HOMEWORK REVIEW ===\n\n1. SPIKERECORDER APP:\nFree app from Backyard Brains. Download at: backyardbrains.com/SpikeRecorder\nWorks standalone (as a simulator) even without SpikerBox hardware.\nStudents who downloaded: great! Students who didn't: do it now on your phone/laptop while we talk.\n\n2. THE THREE PAPERS:\nThe instructor shared three papers after Day 1. These likely connect to the course themes (Drosophila physiology, ion channels, and/or drug discovery with AI). If students haven't read them yet, that's okay — but they should try before Day 3 when the papers become more relevant (Deep Learning and GNNs for molecular biology).\n\n3. RDKIT NEUROTRANSMITTER EXERCISE:\nThis is a good warm-up for today. The exercise: use RDKit to draw the 9 main neurotransmitters as molecular structures in a grid.\nCode skeleton (put on board or share screen):\n\nfrom rdkit import Chem\nfrom rdkit.Chem import Draw, AllChem\n\nneurotransmitters = {\n    \'Serotonin\': \'NCCc1c[nH]c2ccc(O)cc12\',\n    \'Dopamine\': \'NCCc1ccc(O)c(O)c1\',\n    \'GABA\': \'NCCCC(=O)O\',\n    \'Glutamate\': \'N[C@@H](CCC(=O)O)C(=O)O\',\n    \'Acetylcholine\': \'CC(=O)OCC[N+](C)(C)C\',\n    \'Norepinephrine\': \'NCC(O)c1ccc(O)c(O)c1\',\n    \'Glycine\': \'NCC(=O)O\',\n    \'Histamine\': \'NCCc1c[nH]cn1\',\n    \'Adenosine\': \'Nc1ncnc2c1ncn2[C@@H]1O[C@H](CO)[C@@H](O)[C@H]1O\'\n}\nmols = [Chem.MolFromSmiles(s) for s in neurotransmitters.values()]\nimg = Draw.MolsToGridImage(mols, molsPerRow=3, legends=list(neurotransmitters.keys()))\nimg.save(\'neurotransmitters.png\')\n\n4. PROJECT IDEAS:\nStudents should be forming groups of 3-4. If they haven\'t settled on a topic: remind them of the 7 project options from Day 1. The sign-up deadline is today (end of Day 2). Encourage them to think about which option excites them most and which connects to their CS skills.")

    # --- Slide 3: Day 2 Opening — Physiology Overview ---
    n += 1
    make_section_divider(prs, "\U0001f9e0 Physiology Intro — Setting the Stage", n,
        notes="TIMING: 10-15 min\n\nBefore diving into the ML content, do a short physiology primer (~30 min total). The students are CS specialists — they understand algorithms but may have never studied biology or medicine. This section gives them the biological vocabulary to make sense of WHY we are doing molecular ML. Cover the topics listed on the next few slides. Each slide has detailed notes explaining WHAT you are saying and WHY it matters.",
        subtitle="Biology crash-course for Computer Scientists")

    # --- Slide 4: What is Physiology? ---
    n += 1
    make_content_slide(prs,
        "What is Physiology? Why Does it Matter for Drug Discovery?",
        ["Physiology = how living systems FUNCTION (not just their structure)",
         "  → Organs, tissues, cells, proteins, ions — dynamic interaction",
         "  → Example: the heart BEATS because ion channels in cardiac cells open/close rhythmically",
         "",
         "Drugs change physiology — that is their entire purpose:",
         "  AI predicts molecular binding → physiology VALIDATES whether it works",
         "  No matter how good our QSAR model, a drug only works if it changes FUNCTION",
         "",
         "Today\'s physiology topics (brief intro, ~30 min total):",
         "  \u26a1 Action Potentials — how neurons/muscles fire electrically",
         "  \U0001f52c SpikerBot & Computational SpikerBox — record and simulate signals",
         "  \U0001f99f Fly Motion Vision — a model neural circuit",
         "  \U0001f52d Electron Microscopy (EM) — mapping brain wiring (connectomics)",
         "  \U0001f441 Eye-Tracking — measuring drug effects on behavior"],
        n,
        notes="TIMING: 15-20 min\n\n=== WHAT IS PHYSIOLOGY? (explain to CS students) ===\n\nPhysiology is the branch of biology that asks HOW, not WHAT. Anatomy tells you what structures exist (this is the heart, these are its valves). Physiology tells you how those structures work together (the heart muscle contracts because calcium floods into cardiomyocytes, triggering actin-myosin interaction, driven by action potentials from the SA node...).\n\nFor a CS analogy: anatomy is reading the source code, physiology is running the program and profiling it. You need both.\n\nWHY DOES PHYSIOLOGY MATTER FOR DRUG DISCOVERY?\nEvery drug must CHANGE a physiological process. There are only a few ways drugs can do this:\n1. BIND to a receptor → change its signaling (agonist = activates, antagonist = blocks)\n2. INHIBIT an enzyme → reduce production of a molecule\n3. BLOCK an ion channel → change electrical activity\n4. Modify gene expression (some drugs do this)\n5. Physically (antacids neutralize stomach acid)\n\nThe ML pipeline (QSAR, GNNs) predicts how strongly a molecule binds (step 1). Physiology tells you what HAPPENS after binding (steps 2-5). Both are needed.\n\nTODAY\'S AGENDA is a 30-minute tour through the physiology concepts that come up repeatedly in this course. Think of it as vocabulary building — by the end of Day 2, you will have heard terms like \'action potential\', \'ion channel\', \'patch clamp\', \'optogenetics\', \'connectomics\', \'eye-tracking\', and you will know what they mean and why they matter for AI drug discovery.")

    # --- Slide 5: Action Potentials ---
    n += 1
    make_content_slide(prs,
        "Action Potentials: The Electrical Language of Neurons",
        ["A neuron at rest: membrane voltage ~-70 mV (inside more negative than outside)",
         "  → Maintained by Na+/K+ ATPase pump (uses ATP to move ions)",
         "  → Na+ mostly OUTSIDE; K+ mostly INSIDE",
         "",
         "How a spike fires (all-or-nothing, ~1 ms):",
         "  1. Stimulus → depolarization past threshold (~-55 mV)",
         "  2. Voltage-gated Na+ channels OPEN → Na+ rushes in → peak at +40 mV",
         "  3. Na+ channels inactivate; K+ channels open → K+ out → repolarization",
         "  4. Brief hyperpolarization → refractory period → back to rest",
         "",
         "Drug targets at every step:",
         "  Nav blockers: lidocaine (anesthesia), carbamazepine (epilepsy)",
         "  hERG (Kv11.1) blockers: DANGEROUS cardiac side effect of many drugs",
         "  GABA-A modulators: diazepam (Valium), alcohol, propofol (anesthesia)",
         "Hodgkin & Huxley (1952, Nobel 1963): first computational neuroscience model"],
        n,
        notes="TIMING: 20-27 min\n\n=== ACTION POTENTIALS — FULL EXPLANATION FOR CS STUDENTS ===\n\nImagine a neuron as a thin tube (the axon can be meters long in a human!) filled with salty water (cytoplasm), wrapped in a lipid membrane (like a soap bubble wall, ~5 nm thick), immersed in salty fluid (extracellular space).\n\nWHY IS THERE A VOLTAGE ACROSS THE MEMBRANE?\nBecause ions are unevenly distributed. At rest:\n- Na+ (sodium): 145 mM outside, 12 mM inside (12× more outside)\n- K+ (potassium): 4 mM outside, 155 mM inside (39× more inside)\n- Cl- (chloride): 123 mM outside, 4 mM inside (more outside)\nThe cell membrane is much more permeable to K+ than Na+ at rest. K+ wants to flow OUT (concentration gradient). As K+ flows out, it leaves negative charge behind → membrane becomes negative inside. This builds until electric force pulling K+ back equals concentration force pushing it out → equilibrium at ~-70 mV.\n\nThe Na+/K+ ATPase (sodium-potassium pump) uses 1 ATP to move 3 Na+ out and 2 K+ in. This maintains the gradient. About 25% of all ATP in the brain powers these pumps!\n\nTHE SPIKE (step by step):\n1. Input arrives (e.g., from another neuron): briefly opens some channels → membrane depolarizes slightly\n2. If it reaches -55 mV (threshold), VOLTAGE-GATED Na+ channels activate. These channels have a sensor that feels the voltage and snaps open. Na+ floods in (145 mM → 12 mM: 12× concentration gradient PLUS electrical gradient both pulling Na+ in). Voltage shoots to +40 mV in <1 ms.\n3. Na+ channels INACTIVATE (they have a second gate that closes automatically ~1 ms after opening). Voltage-gated K+ channels open (these are slower). K+ floods out. Voltage falls back.\n4. Brief undershoot below -70 mV because K+ channels take a moment to close. Then everything resets.\n\nTHE ALL-OR-NOTHING PRINCIPLE: Either the threshold is crossed and a FULL spike fires (same shape, same size every time), OR nothing happens. This is like a digital bit — ON or OFF. Information is encoded in the TIMING and FREQUENCY of spikes, not their amplitude.\n\nDRUG IMPLICATIONS:\n- LIDOCAINE (dental anesthetic): binds inside the Na+ channel pore and blocks it. No Na+ can enter. No action potentials. No pain signal to brain. (Wears off as drug diffuses away)\n- CARBAMAZEPINE (anticonvulsant for epilepsy): also a Na+ channel blocker. Neurons can\'t fire as rapidly. Seizures require many neurons firing in synchrony — block Nav → fewer synchronized bursts.\n- hERG (pronounced \'h-ERG\', also called Kv11.1): a K+ channel in cardiac muscle. If blocked, the action potential in heart cells lasts longer → QT interval prolongation on ECG → can trigger fatal arrhythmia. EVERY drug candidate must be tested for hERG block. This is why hERG QSAR models are so important (Day 2 practical context!).\n- GABA-A: an inhibitory ligand-gated Cl- channel. When GABA (the main inhibitory neurotransmitter) binds, Cl- flows in → hyperpolarization → neuron less likely to fire. Diazepam (Valium) potentiates GABA-A (makes it more sensitive to GABA). Alcohol does the same. Propofol (surgery anesthetic) activates GABA-A directly.\n\nHODGKIN-HUXLEY MODEL: Two British scientists (Alan Hodgkin and Andrew Huxley) made recordings from squid giant axons (these are very large, 1 mm diameter, easy to record from). They developed mathematical equations (4 coupled ODEs with voltage-dependent conductances) that precisely describe the action potential. 1952 paper, 1963 Nobel Prize in Physiology or Medicine. This is the Day 1 notebook simulation!")

    # --- Slide 6: SpikerBot & Computational SpikerBox ---
    n += 1
    make_content_slide(prs,
        "SpikerBot & The Computational SpikerBox",
        ["SpikerBox (Backyard Brains): open-source DIY electrophysiology",
         "  Records extracellular spikes from neurons/muscles ($100-250 USD)",
         "  Real drug experiments: apply lidocaine → watch spikes disappear",
         "  Ref: Marzullo & Gage (2012) PLoS ONE 7(3):e30837",
         "",
         "SpikerBot: adds stimulation + locomotion + multi-channel recording",
         "  Record FROM and stimulate a living neural preparation",
         "  Bridges AI prediction and physiological validation",
         "",
         "Computational SpikerBox: Hodgkin-Huxley model in software",
         "  Interactively change ion channel conductances → see the effect",
         "  In silico drug testing before wet-lab experiments",
         "",
         "SpikeRecorder app: FREE — record, visualize, analyze, simulate",
         "  Download: backyardbrains.com/SpikeRecorder",
         "  Works on phone/tablet/laptop, no hardware needed for simulation"],
        n,
        notes="TIMING: 27-32 min\n\n=== THE SPIKERBOX ECOSYSTEM ===\n\nBackyard Brains is a company founded by Greg Gage and Tim Marzullo (University of Michigan neuroscience PhD students). Their mission: democratize neuroscience education. The original SpikerBox ($100) records extracellular neural spikes from cockroach legs, earthworms, and human muscles.\n\nHOW EXTRACELLULAR RECORDING WORKS:\nA metal electrode is placed near (but not inside) a neuron. When the neuron fires, the action potential current flows in a loop: Na+ enters the cell at the spike, must come from outside, creating a tiny local current. The electrode detects this current as a ~100-500 microvolt deflection. The SpikerBox amplifies this ~1000× so you can see it on screen or hear it as a 'pop' through speakers.\n\nThis is LESS invasive than patch clamp (whole-cell recording, which is the gold standard but requires a glass micropipette sealed to the cell membrane). Patch clamp can measure individual ion channel currents. Extracellular recording measures population firing. For drug testing: both are used.\n\nDRUG EXPERIMENT EXAMPLE:\n1. Record baseline action potentials from a cockroach leg preparation (the large spiny neurons in the leg fire spontaneously)\n2. Apply lidocaine solution (a drop of local anesthetic)\n3. Watch: spikes gradually disappear (Na+ channels blocked → no more action potentials)\n4. Wash out: spikes return\n5. This IS a dose-response experiment! Our QSAR models predict exactly this IC50 value.\n\nTHE SPIKERBOT:\nAdded on top of SpikerBox: can stimulate (electrical pulses to drive activity), has more channels, can drive a small robot based on neural signals. Project 2 in our course uses SpikerBot.\n\nCOMPUTATIONAL SPIKERBOX:\nThis is a software interface to the Hodgkin-Huxley model. Students can:\n- Drag sliders to change Na+ channel conductance (simulating a Nav blocker drug)\n- See action potential shape and frequency change in real time\n- \'Virtually test\' whether a drug effect would eliminate firing or just slow it\nThis is the \'in silico\' step of the AI → physiology pipeline.\n\nSPIKERECORDER APP:\nFree app. Download it now if you haven\'t. It can:\n1. Connect to SpikerBox hardware via audio jack or Bluetooth\n2. Display real-time spike recordings (action potentials look like little mountains)\n3. Run the computational Hodgkin-Huxley simulation\n4. Do basic spike sorting (separating signals from different neurons by their shape)\n5. Export data for analysis in Python")

    # --- Slide 7: Fly Motion Vision System ---
    n += 1
    make_content_slide(prs,
        "Fly Motion Vision: A Model Neural Circuit",
        ["Why Drosophila (fruit fly)?",
         "  ~140,000 neurons (human: ~86 billion) — fully mappable!",
         "  Powerful genetic tools: optogenetics, calcium imaging, EM",
         "  Conserved ion channel families (same targets as human drugs)",
         "",
         "The visual motion circuit (simplified):",
         "  Photoreceptors (light) → Lamina → Medulla → Lobula plate",
         "  T4 neurons: detect ON-motion (brightening edge moves)",
         "  T5 neurons: detect OFF-motion (darkening edge moves)",
         "  4 directions each → population of direction-selective cells",
         "",
         "Drug-relevant ion channels in this circuit:",
         "  GluCl: glutamate-gated Cl- channel = ivermectin target",
         "  Ammer, Serbe-Kamp et al. (2023) Nat. Neurosci. 26:1894-1905",
         "  Same Cys-loop receptor superfamily as GABA-A (anxiolytics, anesthetics)"],
        n,
        notes="TIMING: 32-37 min\n\n=== THE DROSOPHILA MODEL SYSTEM — WHY FLIES? ===\n\nDrosophila melanogaster (common fruit fly) is the most powerful model organism in genetics and increasingly in neuroscience. Why?\n\n1. SMALL BRAIN, FULLY MAPPABLE: 140,000 neurons. Compare: a mouse has ~71 million neurons, a human ~86 billion. With modern electron microscopy (next slide), we can now map EVERY synapse in the fly brain. The complete wiring diagram (connectome) of the fly brain was published in 2023 (Dorkenwald et al., Nature 634:124-138, 2024). This would be impossible for a mouse, let alone a human.\n\n2. GENETIC TOOLS:\n- Optogenetics: we can insert a gene for a light-sensitive protein (channelrhodopsin, a tool from algae) into ANY specific cell type using fly genetics. Shine blue light → that specific cell fires. Invaluable for testing: \'if I activate/silence THIS neuron, what happens to behavior/circuit activity?\'\n- Calcium imaging: insert a gene for GCaMP (a protein that glows brighter when calcium enters, i.e., when the neuron fires). Now you can watch WHICH neurons fire when the fly sees a moving pattern. Serbe et al. (2016, Neuron 89:829-841) and Serbe-Kamp et al. (2023, J. Neurosci. 43:2497-2514) used this to characterize the T4/T5 motion circuit.\n- EM (electron microscopy): see next slide.\n\n3. CONSERVED BIOLOGY: Fly ion channels are homologous to human ion channels. GluCl (glutamate-gated chloride channel) in flies is in the Cys-loop receptor superfamily — the SAME superfamily as human GABA-A receptors, glycine receptors, and nAChR. Ivermectin targets GluCl in parasites (worms, insects) because it locks the channel open → muscle paralysis. The drug doesn\'t affect vertebrates much because vertebrates don\'t have GluCl (they have GLRA channels instead which are less sensitive to ivermectin). This specificity makes ivermectin safe for humans while lethal to parasites.\n\nTHE MOTION VISION CIRCUIT:\n- Photoreceptors (R1-R8): detect light. Like the rods and cones in your eye.\n- Lamina and Medulla: processing layers (like retinal ganglion cells and LGN in mammals).\n- T4 neurons: respond to a bright edge moving in a specific direction. T5: dark edge. 4 types of each for 4 motion directions (up, down, left, right).\n- Lobula Plate Tangential Cells (LPTCs): integrate motion from many T4/T5 neurons to get a global optic flow signal (like: the whole world is rotating left → fly banks right).\n\nFOR DRUG DISCOVERY: The T5 circuit requires GluCl-mediated inhibition for direction selectivity (Ammer, Serbe-Kamp et al. 2023). Block GluCl → T5 can no longer tell direction. This is the same mechanism used by ivermectin in parasites. Understanding this circuit gave insight into why ivermectin is so effective AND specific.")

    # --- Slide 8: Electron Microscopy ---
    n += 1
    make_content_slide(prs,
        "Electron Microscopy (EM) & Connectomics",
        ["EM = electron microscopy: resolves structures down to ~1 nm",
         "  Light microscopy limit: ~200 nm (wavelength of light)",
         "  EM uses electrons (wavelength ~0.001 nm) → 200× better resolution",
         "  Can see individual synaptic vesicles, cell membranes, protein complexes",
         "",
         "Serial section EM (ssEM): map an entire brain\'s wiring",
         "  Slice tissue into ~30-50 nm sections with a diamond knife",
         "  Image each section in the electron microscope",
         "  Reconstruct 3D volume by aligning sections (huge AI task!)",
         "",
         "Connectomics = the wiring diagram of the brain",
         "  Fly brain: 139,000 neurons, 54.5 million synapses",
         "  Dorkenwald et al. (2024) Nature 634:124-138 — first complete fly brain",
         "",
         "AI is essential: segmenting neurons from EM images = image segmentation",
         "  Same CNNs used for image classification → trained on EM data"],
        n,
        notes="TIMING: 37-41 min\n\n=== ELECTRON MICROSCOPY — WHAT IT IS AND WHY IT MATTERS ===\n\nFor CS students: think of EM as a camera that can photograph individual molecules. A standard light microscope can see a cell (10-100 micrometers) or individual large organelles like the nucleus or mitochondria. But a synapse (the connection between two neurons) is only ~20-40 nanometers (nm) wide. And a single protein is 5-10 nm. Light can\'t resolve these — the wavelength of visible light is 400-700 nm, and you physically cannot see objects smaller than about half the wavelength (Abbe diffraction limit).\n\nElectrons have a MUCH shorter wavelength (~0.001-0.01 nm at typical energies). An electron microscope focuses a beam of electrons onto the sample and detects how the electrons scatter or are absorbed. Resolution: ~0.1-1 nm in practice.\n\nSERIAL SECTION EM (ssEM):\nTo map a whole brain, you need 3D information. The approach:\n1. Fix and stain a piece of brain tissue (heavy metals like osmium stain membranes, making them electron-dense = dark in EM)\n2. Embed in hard plastic\n3. Use an ultramicrotome (a diamond knife on a machine) to slice into sections 30-50 nm thick\n4. Image each section with the electron microscope (takes days to weeks per cubic mm of tissue)\n5. Upload to a computer and use AI (convolutional neural networks) to\n   a. Find all cell membranes in each section\n   b. Follow each neuron\'s outline from one section to the next (3D reconstruction)\n   c. Identify synapses (connections between neurons)\n6. Result: a graph where nodes are neurons and edges are synaptic connections\n\nCONNECTOMICS:\nThe connectome is the complete wiring diagram of a nervous system. For the fly (published 2024, Dorkenwald et al. Nature 634:124-138 — FlyWire project):\n- 139,000 neurons in the fly brain\n- 54.5 million synapses\n- Complete wiring diagram published and publicly available\n- This took multiple labs, years, and massive AI effort\n\nFor context: the C. elegans worm (302 neurons) had its connectome mapped manually in 1986 (Nobel Prize-winning work). The fly brain is 460× larger. Human brain would be ~600,000× the fly brain. Full human connectome: not achievable yet, but active research area.\n\nAI\'s ROLE IN EM AND CONNECTOMICS:\n1. Image segmentation: which pixels in the EM image are part of the same neuron? This is semantic segmentation (same CNN architecture as used for self-driving cars to segment road/cars/pedestrians).\n2. Synapse detection: identifying where two neurons touch and form a synapse.\n3. Proofreading: AI makes mistakes; humans verify using tools like CATMAID or Neuroglancer.\n4. Circuit analysis: once you have the graph, GNNs (same ones we use for molecules!) can analyze circuit motifs, predict function, etc.\n\nCONNECTION TO DRUG DISCOVERY:\nKnowing the wiring diagram helps understand how a drug that hits one neuron type affects the whole circuit. For psychiatric diseases (depression, schizophrenia, anxiety), the circuit-level changes are critical — single-neuron electrophysiology doesn\'t tell the whole story. Connectomics + AI + drug discovery = circuit pharmacology.")

    # --- Slide 9: Eye-Tracking ---
    n += 1
    make_content_slide(prs,
        "Eye-Tracking: Measuring Drug Effects on Behavior",
        ["Eye movements reveal cognitive and neural function:",
         "  Saccades (fast jumps), smooth pursuit, fixation, microsaccades",
         "  Pupil dilation — controlled by autonomic nervous system",
         "",
         "Why eye-tracking for drug discovery?",
         "  Non-invasive, continuous, quantitative measure of brain state",
         "  Many CNS drugs change eye movements (a physiological biomarker)",
         "  Example: antipsychotics affect smooth pursuit tracking",
         "  Example: stimulants dilate pupils (sympathomimetic effect)",
         "  Example: opioids cause miosis (pin-point pupils)",
         "",
         "Technical setup:",
         "  Infrared camera tracks corneal reflection and pupil position",
         "  Modern trackers: 1000 Hz, sub-degree accuracy",
         "",
         "AI + Eye-tracking:",
         "  Classify drug state from eye movement patterns",
         "  Predict treatment response in clinical trials",
         "  Project option 4 in our course!"],
        n,
        notes="TIMING: 41-45 min\n\n=== EYE-TRACKING — WHAT IT IS AND HOW IT WORKS ===\n\nEyes are the window to the brain. Eye movements are controlled by a network of brain areas (frontal eye fields, superior colliculus, cerebellum, brainstem nuclei) and muscles (6 extraocular muscles per eye). Because this system is complex and connected to many brain areas, eye movements can reveal the state of many different neural systems.\n\nTYPES OF EYE MOVEMENTS:\n- Saccades: rapid, ballistic eye movements (jumps) from one fixation point to another. Fast (200-700 degrees/second!). Cannot be changed once initiated (like a ballistic missile). Saccade latency (~200 ms) reflects decision-making processes.\n- Smooth pursuit: when following a moving object. Requires the object to actually be moving. Controlled by visual cortex + cerebellum. Many psychiatric drugs impair smooth pursuit.\n- Microsaccades: tiny involuntary movements during fixation. Reflect attention and arousal state.\n- Pupil dilation (pupillometry): the pupil is controlled by the autonomic nervous system. Sympathetic = dilates (fight-or-flight). Parasympathetic = constricts. Cognitive load and arousal also affect pupil size.\n\nDRUG EFFECTS ON EYE MOVEMENTS:\n- Antipsychotics (dopamine D2 blockers): smooth pursuit impairment — patients and people on these drugs show characteristic \'staircase\' pursuit (the eye jumps instead of smoothly following)\n- Opioids (morphine, fentanyl): miosis (pupil constriction) because opioids activate the Edinger-Westphal nucleus (parasympathetic). Police use this to test for opioid intoxication.\n- Stimulants (cocaine, amphetamine): mydriasis (pupil dilation) because sympathomimetic effects.\n- Benzodiazepines (diazepam): reduce saccade velocity and increase fixation duration\n- Alcohol: smooth pursuit breaks down, nystagmus (involuntary oscillation) at high doses\n\nHOW EYE-TRACKERS WORK:\nMost modern eye-trackers use a near-infrared (NIR) LED to illuminate the eye. A camera captures the reflection. Two key features are tracked:\n1. Corneal reflection (CR): the reflection of the NIR LED off the curved corneal surface — moves with head movements but not eye movements\n2. Pupil center: tracked by thresholding the dark pupil area\nGaze direction = vector from CR to pupil center. This is head-movement independent.\n\nModern trackers: 1000 Hz sampling rate, <0.5 degree accuracy. Can track during free head movement (using glasses-mounted systems) or on a screen (tower-mounted).\n\nAI + EYE-TRACKING:\n1. Time-series classification: eye movement trace → classify drug state (e.g., \'sober\' vs \'0.08% BAC\'). Can be done with LSTM, transformer, or even RF on features.\n2. Pupillometry: track pupil diameter over time → extract pupil light reflex parameters (latency, amplitude, re-dilation rate) → these are biomarkers for autonomic nervous system state.\n3. Scanpath analysis: sequence of fixations and saccades → reveals what a person is \'reading\' visually (e.g., does a drug impair reading a cluttered visual scene?).\n4. Clinical applications: multiple sclerosis, Parkinson\'s, schizophrenia all have characteristic eye movement signatures. AI can detect these with high accuracy — potentially earlier than clinical symptoms.\n\nCOURSE PROJECT 4 (Eye-tracking): Build a classifier that distinguishes different cognitive states or drug conditions from eye movement features. Uses the skills from Day 2 (feature engineering, RF, evaluation) applied to time-series data instead of molecular data.")


    # --- Slide 11: Morning Agenda ---
    n += 1
    make_content_slide(prs,
        "Morning Agenda (Units 1-4)",
        ["1. SMILES deep dive — syntax rules, practice, limitations",
         "2. Molecular fingerprints — Morgan/ECFP, Tanimoto similarity",
         "3. Feature engineering — descriptors, Lipinski, BBB rules",
         "4. QSAR: Quantitative Structure-Activity Relationships",
         "5. ML baselines: Random Forest & XGBoost",
         "6. PRACTICAL 1: Build a QSAR model for solubility prediction",
         "",
         "☕ BREAK after Unit 2 (after ~90 min)",
         "",
         "CNS drug focus: blood-brain barrier penetration, neuroscience targets"],
        n,
        notes="TIMING: 10-12 min\n\nQuick overview. Emphasize the logical flow: represent molecules (SMILES → fingerprints → descriptors) → framework for prediction (QSAR) → ML tools (RF, XGBoost). Practical uses the Delaney solubility dataset.")

    # --- Slide 3: SMILES Deep Dive ---
    n += 1
    make_section_divider(prs, "SMILES: Deep Dive", n,
        notes="TIMING: 12-15 min\n\nTransition to SMILES syntax. Weininger (1988) 'SMILES, a Chemical Language and Information System. 1. Introduction to Methodology and Encoding Rules.' J. Chem. Inf. Comput. Sci. 28:31-36.",
        subtitle="Weininger (1988), J. Chem. Inf. Comput. Sci. 28:31-36")

    # --- Slide 4: SMILES Syntax ---
    n += 1
    make_content_slide(prs,
        "SMILES Syntax Rules",
        ["Atoms: C, N, O, S, P, F, Cl, Br, I (organic subset, implicit H)",
         "Single bond: implicit (CC = ethane)  |  Double: = (C=O)  |  Triple: # (C#N)",
         "Aromatic: lowercase (c1ccccc1 = benzene)",
         "Branches: parentheses — CC(=O)O = acetic acid",
         "Rings: matching digits — C1CCCCC1 = cyclohexane",
         "Charges: [NH4+], [O-]  |  Stereochemistry: / \\ for E/Z, @ @@ for R/S",
         "",
         "Canonical SMILES: unique representation (RDKit Chem.MolToSmiles())",
         "",
         "Alternatives:",
         "  SELFIES — 100% valid strings (Krenn et al. 2020) — great for generative AI",
         "  InChI — IUPAC standard  |  DeepSMILES — neural network friendly"],
        n,
        notes="TIMING: 15-25 min\n\nSpend 7 min on syntax with board examples. Build from simple (ethanol CCO) to complex (benzene c1ccccc1). Key point about canonical SMILES: OCC and CCO are both valid for ethanol, but canonical is CCO. Mention alternatives briefly — SELFIES will return on Day 3 for generative AI.")

    # --- Slide 5: SMILES Practice ---
    n += 1
    make_content_slide(prs,
        "SMILES Practice: Can You Read These?",
        ["1. NCCc1c[nH]c2ccc(O)cc12  →  ?  (hint: neurotransmitter, mood)",
         "2. NCCc1ccc(O)c(O)c1  →  ?  (hint: reward pathway)",
         "3. NCCCC(=O)O  →  ?  (hint: main inhibitory NT in the brain)",
         "4. CC(=O)Oc1ccccc1C(=O)O  →  ?  (hint: common painkiller)",
         "5. Cn1c(=O)c2c(ncn2C)n(C)c1=O  →  ?  (hint: in your coffee)",
         "",
         "Answers: 1) Serotonin  2) Dopamine  3) GABA  4) Aspirin  5) Caffeine",
         "",
         "Note: GABA is simpler than serotonin, but both are critical NTs",
         "GABA-A receptor is in the same Cys-loop superfamily as GluCl (ivermectin target)"],
        n,
        notes="TIMING: 25-30 min\n\n5 min interactive. Reveal answers one by one. Neuroscience SMILES walkthrough:\n1. SEROTONIN: indole ring + ethylamine + hydroxyl. Targeted by SSRIs.\n2. DOPAMINE: catechol ring + ethylamine. Targeted by L-DOPA, antipsychotics.\n3. GABA: simplest — amine + 4C + carboxyl. GABA-A = same superfamily as GluCl!\n4. ASPIRIN: acetyl + salicylic acid. COX inhibitor.\n5. CAFFEINE: purine xanthine. Adenosine A1/A2A antagonist.\n\nAsk: notice how simple GABA is vs serotonin? Structure ≠ biological importance.")

    # --- Slide 6: Molecular Fingerprints ---
    n += 1
    make_section_divider(prs, "Molecular Fingerprints", n,
        notes="TIMING: 30-32 min\n\nTransition: SMILES = text. For ML, we need numerical vectors. Enter fingerprints.",
        subtitle="From molecular structure to numerical vectors")

    # --- Slide 7: Fingerprint Concept ---
    n += 1
    make_content_slide(prs,
        "Molecular Fingerprints: The Concept",
        ["A fingerprint converts a molecule into a fixed-length bit vector",
         "Each bit indicates presence/absence of a substructure",
         "",
         "Types:",
         "  Structural keys (MACCS, 166 bits) — predefined patterns",
         "  Topological (RDKit FP) — paths through molecular graph",
         "  Circular (Morgan/ECFP) — local atom environments at radius r",
         "",
         "Morgan fingerprints are the most widely used today",
         "  Rogers & Hahn (2010), JCIM 50:742-754",
         "  ECFP4 = radius 2, 2048 bits = default starting point",
         "",
         "In RDKit: AllChem.GetMorganFingerprintAsBitVect(mol, radius=2, nBits=2048)"],
        n,
        notes="TIMING: 32-40 min\n\nSpend 5 min. Barcode analogy. Morgan/ECFP captures local atom neighborhoods at increasing radius. ECFP4 (radius=2, 2048 bits) is the standard. CITATION: Rogers & Hahn (2010) 'Extended-Connectivity Fingerprints.' J. Chem. Inf. Model. 50:742-754. This is the foundational paper that formalized ECFP as Morgan fingerprints.")

    # --- Slide 8: Morgan Algorithm ---
    n += 1
    make_content_slide(prs,
        "Morgan / ECFP: How They Work",
        ["Algorithm (for each atom):",
         "  1. Start: identifier = atom type + properties",
         "  2. Iteration 1: collect identifiers from radius-1 neighbors",
         "  3. Hash combined info → new identifier",
         "  4. Iteration 2: expand to radius-2, hash again",
         "  5. Map all identifiers to fixed-length bit vector",
         "",
         "ECFP4 = radius 2  |  ECFP6 = radius 3",
         "",
         "Similarity: Tanimoto coefficient = |A∩B| / |A∪B|",
         "  Range: 0 (different) to 1 (identical)",
         "  Tanimoto > 0.85 → likely similar activity (but exceptions exist!)",
         "  In RDKit: DataStructs.TanimotoSimilarity(fp1, fp2)"],
        n,
        notes="TIMING: 40-50 min\n\nSpend 7 min with board diagram. Draw a molecule, pick one atom, show radius-1 and radius-2 neighborhoods. The connection to message passing (Day 3) is important — same idea, but hashed instead of learned.\n\nTanimoto: Jaccard index. The 0.85 threshold is heuristic (Similar Property Principle). Exceptions = activity cliffs (covered later today).")

    # --- Slide 9: Feature Engineering ---
    n += 1
    make_content_slide(prs,
        "Feature Engineering & Lipinski's Rule of Five",
        ["Physicochemical descriptors:",
         "  MW, LogP (lipophilicity), HBD, HBA, TPSA, rotatable bonds, aromatic rings",
         "",
         "Lipinski (Pfizer, 1997) — most cited paper in medicinal chemistry:",
         "  Orally bioavailable if: MW ≤ 500, LogP ≤ 5, HBD ≤ 5, HBA ≤ 10",
         "",
         "For CNS drugs — crossing the blood-brain barrier (Pardridge 2005):",
         "  MW < 450, LogP 1-3, HBD < 3, TPSA < 90 Å²",
         "",
         "Example: serotonin (MW=176, LogP=0.2) does NOT cross BBB well",
         "  → That's why we give SSRIs (which cross BBB) not serotonin itself",
         "  → L-DOPA crosses BBB, dopamine doesn't → Parkinson's treatment",
         "",
         "RDKit: 200+ descriptors  |  Mordred: >1800 descriptors"],
        n,
        notes="TIMING: 50-60 min\n\n7 min. CITATION: Lipinski et al. (1997) 'Experimental and computational approaches to estimate solubility and permeability in drug discovery and development settings.' Adv. Drug Deliv. Rev. 23(1-3):3-25. One of the most cited papers in medicinal chemistry. For CNS drugs, BBB adds constraints. Pardridge (2005) 'The blood-brain barrier: Bottleneck in brain drug development.' NeuroRx 2:3-14.\n\nThe serotonin/dopamine BBB example is powerful: explains why Parkinson's patients take L-DOPA, not dopamine. This is a PERFECT ML task: predict BBB permeability from descriptors.")

    # --- Slide 10: QSAR ---
    n += 1
    make_content_slide(prs,
        "QSAR: Quantitative Structure-Activity Relationships",
        ["Core idea: molecular structure determines biological activity",
         "  Activity = f(molecular features)",
         "  Hansch (1964) — founded the field",
         "",
         "Modern QSAR workflow:",
         "  1. Curate dataset from ChEMBL (target + activity data)",
         "  2. Generate features (fingerprints + descriptors)",
         "  3. Split data (scaffold split, NOT random!)",
         "  4. Train model (RF, XGBoost, GNN)",
         "  5. Evaluate rigorously (RMSE, R², AUC-ROC)",
         "  6. Interpret (SHAP, feature importance)",
         "",
         "This is the core pipeline you'll use for your projects!"],
        n,
        notes="TIMING: 60-68 min\n\n5 min overview. QSAR was formalized by Corwin Hansch in 1964 — over 60 years old! The modern version adds ML, but the fundamental idea is unchanged: structure → activity. Emphasize that steps 3-6 are critical and often done poorly.")

    # --- Slide 11: RF & XGBoost ---
    n += 1
    make_two_column_slide(prs,
        "ML Baselines: Random Forest & XGBoost",
        "Random Forest",
        ["Ensemble of decision trees (Breiman 2001)",
         "Each tree trained on bootstrap sample",
         "Final prediction = average (regression) or vote",
         "Handles high-dimensional fingerprints well",
         "Resistant to overfitting",
         "Easy to interpret (feature importance)"],
        "XGBoost",
        ["Sequential trees correct errors (Chen & Guestrin 2016)",
         "Often higher accuracy than RF",
         "Built-in regularization",
         "Handles missing values",
         "Dominant in Kaggle competitions",
         "Sheridan (2016, JCIM 56:2353-2360): XGBoost modestly > RF for QSAR"],
        n,
        notes="TIMING: 68-78 min\n\n7 min. RF: ensemble of independent trees. XGBoost: sequential boosting. CITATION: Chen & Guestrin (2016) 'XGBoost: A Scalable Tree Boosting System.' KDD '16, pp. 785-794. Won Best Paper at KDD 2016. Breiman (2001) 'Random forests.' Machine Learning 45:5-32 defined Random Forest. For tabular data (which molecular features are), XGBoost is still king.\n\nKey message: ALWAYS try RF and XGBoost as baselines before jumping to deep learning. They often match or beat GNNs on small datasets.")

    # --- Slide 12: Break ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="TIMING: 78-93 min (after Unit 2)\n\nBreak! Students should stretch. Morning practical is coming up next.",
        subtitle="Practical 1 starts right after the break!")

    # --- Slide 13: Practical 1 ---
    n += 1
    make_content_slide(prs,
        "Practical 1: Build a QSAR Model",
        ["Open: day2_molecular_ml/Day2_Practical_QSAR.ipynb",
         "",
         "Dataset: Delaney solubility (1,128 molecules, logS values)",
         "  Delaney (2004), J. Chem. Inf. Comput. Sci. 44:1000-1005",
         "",
         "Steps:",
         "  1. Load dataset, parse SMILES with RDKit",
         "  2. Generate Morgan fingerprints (ECFP4, 2048 bits)",
         "  3. Calculate physicochemical descriptors",
         "  4. 80/20 train/test split",
         "  5. Train Random Forest regressor",
         "  6. Train XGBoost regressor",
         "  7. Compare: RMSE, R-squared",
         "",
         "Time: ~50 min including code execution | Work in pairs"],
        n,
        notes="TIMING: 93-95 min (intro) then 95-145 min (practical)\n\n2 min intro then hands-on. The Delaney dataset is a classic benchmark. Solubility matters because ~40% of drug candidates have issues. Expected: RF RMSE ~0.6-0.8 logS, R² ~0.85-0.90. XGBoost slightly better.\n\nIMPORTANT: Note that this initial split is RANDOM. In the afternoon, we'll show why this overestimates performance!")

    # --- Slide 14: Practical 1 hands-on ---
    n += 1
    make_section_divider(prs, "🔬 PRACTICAL 1 — Build Your QSAR Model", n,
        notes="TIMING: 95-145 min (Units 3-4, ~50 min coding)\n\nWalk around, help students. Ensure everyone gets RF and XGBoost trained. They'll need these models for the afternoon evaluation session.",
        subtitle="50 minutes of coding — train your first molecular ML model!")

    # --- Slide 15: Afternoon Agenda ---
    n += 1
    make_content_slide(prs,
        "Afternoon Agenda (Units 5-8)",
        ["7. Why standard evaluation fails for molecular data",
         "8. Activity cliffs and scaffold bias",
         "9. Proper validation: scaffold splits, temporal splits",
         "10. Metrics deep dive: RMSE, R², ROC-AUC, precision-recall",
         "11. Explainability with SHAP (Lundberg & Lee 2017)",
         "12. PRACTICAL 2: Re-evaluate your morning model properly",
         "",
         "☕ BREAK after Unit 6 (after ~90 min afternoon)",
         "",
         "Key message: the EVALUATION is more important than the model"],
        n,
        notes="TIMING: 145-148 min (Unit 5 start)\n\nQuick 3-min overview. This afternoon is about rigor. In pharmaceutical companies, a model that looks great on paper but fails in practice costs millions. By the end of today, students will understand why their morning models might be misleading.")

    # --- Slide 16: Hall of Shame ---
    n += 1
    make_story_slide(prs,
        "The Hall of Shame: When Models Lie",
        "Example 1: hERG toxicity model — 95% accuracy on random split. On novel scaffolds: 60%.\n\nExample 2: Published AUC-ROC 0.97. Remove duplicate compounds: 0.72.\n\nExample 3: Scaffold split vs random — AUROC dropped from 0.92 to 0.75.\n\nExample 4 (Neuroscience): A GABA-A binding model looked great, but it learned to recognize the benzodiazepine scaffold, not actual binding features. On novel chemotypes: near-random.\n\nLesson: ALWAYS validate properly.\n\nRef: Wallach & Heifets (2018) — 'Most Ligand-Based Benchmarks Reward Memorization'",
        n,
        notes="TIMING: 148-155 min\n\n5 min cautionary tales. CITATION: Wallach & Heifets (2018) 'Most Ligand-Based Classification Benchmarks Reward Memorization Rather than Generalization.' JCIM 58:916-932. Showed many published models memorize molecular series rather than learning generalizable chemical features.\n\nGABA-A example: GABA-A is same Cys-loop family as GluCl (Dr. Serbe-Kamp's research). A dataset dominated by benzodiazepines (all sharing 1,4-benzodiazepine core) will train models that recognize scaffolds, not binding mechanisms. On random split: great (benzodiazepines in both sets). On scaffold split: fails for novel chemotypes (neurosteroids, barbiturates, Z-drugs).")

    # --- Slide 17: Scaffold Bias ---
    n += 1
    make_content_slide(prs,
        "Scaffold Bias and Proper Splitting",
        ["Scaffold = core ring structure (Bemis-Murcko framework)",
         "Problem: random split → same scaffold in train AND test → memorization",
         "",
         "Solutions:",
         "  Scaffold split: no scaffold overlap (Bemis & Murcko 1996)",
         "  Temporal split: train on older, test on newer data",
         "  Cluster split: cluster molecules, split by cluster",
         "",
         "MoleculeNet uses scaffold split as default (Wu et al. 2018)",
         "",
         "Activity cliffs: very similar molecules, very different activity",
         "  Changing one methyl → chlorine can change IC50 by 1000×",
         "  These are the hardest cases for ML!"],
        n,
        notes="TIMING: 155-163 min\n\n5 min. Draw Bemis-Murcko framework: strip molecule to ring systems + linkers. If train and test share scaffolds, the model memorizes scaffolds. CITATION: Wu et al. (2018) 'MoleculeNet: A Benchmark for Molecular Machine Learning.' Chemical Science 9:513-530. Made scaffold splitting the standard for evaluation.")

    # --- Slide 18: Metrics ---
    n += 1
    make_two_column_slide(prs,
        "Evaluation Metrics for Molecular ML",
        "Regression",
        ["RMSE — root mean squared error",
         "  Penalizes large errors more",
         "R² — coefficient of determination",
         "  1.0 = perfect, <0 = worse than mean",
         "MAE — mean absolute error",
         "  More robust to outliers",
         "",
         "Report ALL three, not just R²!"],
        "Classification",
        ["AUC-ROC — area under ROC curve",
         "  0.5 = random, 1.0 = perfect",
         "  Insensitive to class balance!",
         "Precision-Recall AUC",
         "  Better for imbalanced data (most drug data)",
         "Balanced accuracy",
         "  Average of sensitivity + specificity",
         "",
         "Report PR-AUC for imbalanced data!"],
        n,
        notes="TIMING: 163-170 min\n\n5 min reference slide. Key point: AUC-ROC can be misleading for imbalanced data (common in drug discovery: many inactives, few actives). PR-AUC is more informative. For regression, report all three metrics.")

    # --- Slide 19: Interpretability Section ---
    n += 1
    make_section_divider(prs, "Model Interpretability", n,
        notes="TIMING: 170-172 min\n\nTransition: We've evaluated quantitatively. Can we understand WHY?",
        subtitle="Opening the black box — why does the model predict THIS?")

    # --- Slide 20: Why Interpretability ---
    n += 1
    make_content_slide(prs,
        "Why Explainability Matters",
        ["Regulatory: FDA expects understanding of model behavior",
         "Scientific: chemists won't trust a model they can't understand",
         "Safety: unexplained predictions could mask dangerous failures",
         "Discovery: understanding features drives new hypotheses",
         "",
         "Jimenez-Luna et al. (2020), Nature Machine Intelligence 2:573-584",
         "",
         "Types:",
         "  Global: which features are generally important?",
         "  Local: why did the model predict THIS for THIS molecule?",
         "",
         "EU AI Act: medical AI = high-risk = requires transparency"],
        n,
        notes="TIMING: 172-177 min\n\n3 min. CITATION: Jiménez-Luna et al. (2020) 'Drug discovery with explainable artificial intelligence.' Nature Machine Intelligence 2:573-584. The best review for drug discovery interpretability. If a model says a molecule is toxic but can't explain why, the chemist will be skeptical. If it can say 'the nitro group at position 3 drives the toxicity prediction,' the chemist engages.")

    # --- Slide 21: SHAP ---
    n += 1
    make_content_slide(prs,
        "SHAP: SHapley Additive exPlanations",
        ["Lundberg & Lee (2017), NeurIPS",
         "Based on Shapley values from cooperative game theory",
         "",
         "For each prediction, assign a contribution to each feature:",
         "  Positive SHAP → pushes prediction higher",
         "  Negative SHAP → pushes prediction lower",
         "  Sum of all SHAP values + baseline = predicted value",
         "",
         "Advantages:",
         "  Mathematically grounded (unique solution with desirable properties)",
         "  Works for any model (model-agnostic)",
         "  Local AND global explanations",
         "  Beautiful visualizations (beeswarm, waterfall, force plots)",
         "",
         "For fingerprints: important bit → decode to substructure → chemistry insight"],
        n,
        notes="TIMING: 177-185 min\n\n5 min. Shapley value: each feature is a 'player' in a prediction 'game'. CITATION: Lundberg & Lee (2017) 'A Unified Approach to Interpreting Model Predictions.' NeurIPS 30:4765-4774. TreeSHAP (Lundberg et al. 2020, Nature MI 2:56-67) is exact and fast for RF/XGBoost.\n\nFor molecular fingerprints: if bit 847 is most important and corresponds to a pyridine ring, we learn actual chemistry. RDKit GetMorganFingerprintBitInfo() decodes bits.")

    # --- Slide 22: Applicability Domain ---
    n += 1
    make_content_slide(prs,
        "Applicability Domain",
        ["A model is only reliable within its training data distribution",
         "",
         "Methods to define AD:",
         "  Distance-based: is the new molecule close to training data?",
         "  Descriptor range: are features within training bounds?",
         "  Conformal prediction: distribution-free prediction intervals",
         "",
         "Sahigara et al. (2012), Molecules 17(5):4791-4810",
         "",
         "In practice: flag out-of-domain predictions as low confidence",
         "A confident wrong prediction is worse than an honest 'I don't know'"],
        n,
        notes="TIMING: 185-190 min\n\n3 min. A model trained on kinase inhibitors shouldn't predict GPCR ligands. Conformal prediction is gaining popularity for guaranteed coverage.")

    # --- Slide 23: Break ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="TIMING: 190-205 min (after Unit 6)\n\nSecond break of the day. Practical 2 coming up.",
        subtitle="Final practical of the day coming up!")

    # --- Slide 24: Practical 2 ---
    n += 1
    make_content_slide(prs,
        "Practical 2: Evaluate & Interpret Your QSAR Model",
        ["Open: day2_molecular_ml/Day2_Practical_Evaluation.ipynb",
         "",
         "Steps:",
         "  1. Load your morning QSAR model (or re-train quickly)",
         "  2. Implement scaffold splitting (Bemis-Murcko)",
         "  3. Compare: random split vs. scaffold split performance",
         "  4. Compute ROC curve, precision-recall curve",
         "  5. Calculate SHAP values for your RF model",
         "  6. Create SHAP beeswarm plot: which features matter most?",
         "  7. Decode important fingerprint bits to substructures",
         "  8. Define applicability domain using Tanimoto distance",
         "",
         "Time: ~50 min | The 'aha moment': watch performance DROP with scaffold split"],
        n,
        notes="TIMING: 205-210 min (intro) then 210-260 min (practical)\n\n2 min overview, then hands-on. THE KEY LEARNING MOMENT: students see their R² drop from ~0.85-0.90 (random) to ~0.70-0.80 (scaffold). That's the real performance. The SHAP analysis is the fun part.")

    # --- Slide 25: Practical 2 hands-on ---
    n += 1
    make_section_divider(prs, "🔬 PRACTICAL 2 — Evaluate & Interpret", n,
        notes="TIMING: 210-260 min (Units 7-8, ~50 min coding)\n\nWalk around. Make sure everyone sees the scaffold split effect. Help with SHAP plots.",
        subtitle="50 minutes — discover what your model really knows!")

    # --- Slide 26: Discussion ---
    n += 1
    make_discussion_slide(prs,
        "A model achieves AUC-ROC 0.95 on the test set.\nWould you trust it to guide a $100M drug campaign?",
        ["What kind of split was used? Random or scaffold?",
         "How imbalanced is the dataset? (Use PR-AUC instead!)",
         "Has it been validated prospectively on truly new chemical series?",
         "What is the applicability domain?"],
        n,
        notes="TIMING: 260-270 min\n\n7-min discussion. The answer is NO — not without more info. Key questions: What split? What balance? Prospective validation? In pharma, models go through extensive validation before decisions.")

    # --- Slide 27: Takeaways ---
    n += 1
    make_takeaway_slide(prs,
        ["Morgan/ECFP fingerprints convert molecules to bit vectors",
         "QSAR: predicting activity from structure (since 1964!)",
         "RF and XGBoost: strong, production-ready QSAR baselines",
         "Scaffold splits prevent data leakage (Wallach & Heifets 2018)",
         "SHAP: mathematically grounded explanations for any model",
         "Applicability domain: know your model's limits",
         "The evaluation is more important than the model!",
         "Next (Day 3, 5 May): Deep learning, GNNs, generative AI, AlphaFold"],
        2, n,
        notes="TIMING: 270-275 min\n\nQuick recap. The RF/XGBoost + ECFP + scaffold split + SHAP pipeline is production-ready — used in real pharma companies today.")

    # --- Slide 28: Next Day ---
    n += 1
    make_content_slide(prs,
        "Coming Up: Day 3 — 5 May 2026",
        ["Morning: Deep Learning & Graph Neural Networks",
         "  Molecules as graphs → message passing → GNNs",
         "  Practical: train a GNN on MoleculeNet data",
         "",
         "Afternoon: Generative AI & Protein Targets",
         "  Generate new molecules (VAE, diffusion)",
         "  AlphaFold for drug targets",
         "  Practical: explore generative models or AlphaFold",
         "",
         "Preparation:",
         "  Review today's SHAP analysis — understand your model",
         "  Start discussing project approach with your group"],
        n,
        notes="TIMING: 275-278 min\n\nQuick preview. Day 3 moves from classical ML to deep learning. Encourage students to discuss project approaches with their groups between now and then.")

    os.makedirs("day2_molecular_ml", exist_ok=True)
    prs.save("day2_molecular_ml/slides.pptx")
    print(f"  Day 2: {n} slides → day2_molecular_ml/slides.pptx")


# =====================================================================
# DAY 3 — 5 May 2026 — 8 units (360 min)
# Deep Learning, GNNs, Generative AI, Protein Targets / AlphaFold
# =====================================================================
def generate_day3():
    prs = new_prs()
    n = 0

    # --- Title ---
    n += 1
    make_title_slide(prs,
        "Deep Learning, GNNs & Generative AI",
        "When molecules become graphs and AI creates new drugs",
        3, "AI for Drug Discovery", date_str="5 May 2026",
        notes="TIMING: 0-10 min (Unit 1 start)\n\nWelcome back after the break! Quick recap: on Day 2 you built QSAR models with fingerprints + RF/XGBoost, learned about scaffold splits, SHAP. Today we level up: molecules as graphs → GNNs → generative AI → AlphaFold.\n\nAgenda:\nMorning (Units 1-4): Deep learning, GNNs, Practical 1\nAfternoon (Units 5-8): Generative AI, AlphaFold, Practical 2")

    # --- Slide 2: Morning Agenda ---
    n += 1
    make_content_slide(prs,
        "Morning Agenda (Units 1-4)",
        ["1. From fingerprints to learned representations",
         "2. Molecules as graphs — nodes, edges, features",
         "3. GNN intuition: message passing = learned fingerprints",
         "4. GNN architectures: GCN, GAT, MPNN",
         "5. Real-world GNN applications: pharma & neuroscience",
         "6. PRACTICAL 1: Train a GNN with DeepChem (~50 min)",
         "",
         "☕ BREAK after Unit 2 (after ~90 min)"],
        n,
        notes="TIMING: 10-12 min\n\nKey insight: GNNs do the SAME thing as Morgan fingerprints (expand neighborhoods, aggregate) but LEARN the aggregation instead of hashing.")

    # --- Slide 3: Why Deep Learning? ---
    n += 1
    make_content_slide(prs,
        "From Fingerprints to Learned Representations",
        ["Day 2 pipeline: SMILES → fingerprints → RF/XGBoost → prediction",
         "",
         "Fingerprint limitations:",
         "  Fixed-length bit vectors lose information (hash collisions)",
         "  Cannot capture 3D spatial relationships",
         "  Hand-designed, not learned from data",
         "",
         "Deep learning solution: learn representations end-to-end",
         "  Input: molecular graph (atoms + bonds)",
         "  Model: Graph Neural Network learns optimal features",
         "  Output: property prediction",
         "",
         "Key question: when does this matter? (hint: dataset size)"],
        n,
        notes="TIMING: 12-20 min\n\n5 min. The key limitation: fingerprint information loss from hashing. GNNs solve this by learning representations directly from graph structure. BUT — RF on fingerprints often matches GNNs on small datasets. The advantage emerges with larger datasets (>10K molecules).")

    # --- Slide 4: Molecules as Graphs ---
    n += 1
    make_two_column_slide(prs,
        "Molecules as Graphs",
        "Atom (Node) Features",
        ["Atomic number (C=6, N=7, O=8)",
         "Degree (number of bonds)",
         "Formal charge",
         "Hybridization (sp, sp2, sp3)",
         "Aromaticity",
         "Number of hydrogens"],
        "Bond (Edge) Features",
        ["Bond type (single, double, triple, aromatic)",
         "Is in ring?",
         "Is conjugated?",
         "Stereochemistry (E/Z, cis/trans)",
         "",
         "Same features as fingerprints,",
         "but preserved per atom/bond!"],
        n,
        notes="TIMING: 20-27 min\n\n5 min. Molecules are NATURALLY graphs. The key difference from fingerprints: features are preserved as separate attributes per atom/bond. The GNN learns which features matter and how to combine them.")

    # --- Slide 5: Message Passing ---
    n += 1
    make_content_slide(prs,
        "GNN Intuition: Message Passing",
        ["Core idea: each node gathers info from neighbors, then updates",
         "",
         "Algorithm (for each node v, K iterations):",
         "  1. AGGREGATE: collect features from all neighbors of v",
         "  2. UPDATE: combine message with v's own features",
         "  3. After K iterations: each node knows K-hop neighborhood",
         "",
         "K=1: direct neighbors (≈ Morgan radius=1)",
         "K=2: 2-hop neighborhood (≈ Morgan radius=2)",
         "",
         "This IS what ECFP does — but LEARNED, not hashed!",
         "",
         "After K layers → READOUT: aggregate all node features → prediction",
         "  h_graph = mean/sum/attention pooling of all node representations"],
        n,
        notes="TIMING: 27-37 min\n\n7 min. Analogy: each atom is a person at a party. Round 1: talk to direct friends. Round 2: friends tell you what they learned. After K rounds, you know about people up to K handshakes away.\n\nDraw on board: pick an atom in aspirin, trace neighborhoods at r=1, r=2. The connection to Morgan fingerprints is the KEY insight.\n\nMath: m_v^(k) = AGG({h_u^(k-1): u ∈ N(v)}); h_v^(k) = UPDATE(h_v^(k-1), m_v^(k))")

    # --- Slide 6: GNN Architectures ---
    n += 1
    make_content_slide(prs,
        "Common GNN Architectures",
        ["GCN — Graph Convolutional Network (Kipf & Welling 2017, ICLR)",
         "  Simple, effective, widely adopted (~56,000+ citations as of 2026)",
         "",
         "GAT — Graph Attention Network (Veličković et al. 2018, ICLR)",
         "  Learns attention weights: some neighbors matter more",
         "",
         "MPNN — Message Passing Neural Network (Gilmer et al. 2017, ICML)",
         "  The unifying framework for molecular GNNs",
         "",
         "D-MPNN — Directed MPNN / Chemprop (Yang et al. 2019, JCIM)",
         "  State-of-the-art for molecular property prediction",
         "",
         "3D-aware: SchNet, DimeNet — use atomic coordinates",
         "",
         "Key insight: architecture matters LESS than data quality + featurization"],
        n,
        notes="TIMING: 37-44 min\n\n5 min overview.\nCITATION: GCN: Kipf & Welling (2017) 'Semi-Supervised Classification with Graph Convolutional Networks.' ICLR 2017. ~56,000 citations as of 2026 — one of the most influential ML papers.\nCITATION: GAT: Veličković et al. (2018) 'Graph Attention Networks.' ICLR 2018. Learns attention weights so not all neighbors contribute equally.\nCITATION: MPNN: Gilmer et al. (2017) 'Neural Message Passing for Quantum Chemistry.' ICML 2017, pp. 1263-1272. The unifying framework — shows GCN, GAT, and others are all special cases of message passing.\nCITATION: Chemprop/D-MPNN: Yang et al. (2019) JCIM 59:3370-3388. Best out-of-the-box molecular GNN — use it for projects!")

    # --- Slide 7: Break ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="TIMING: 44-59 min (after Unit 2)\n\nBreak.",
        subtitle="GNN practical coming up next!")

    # --- Slide 8: Real-World Applications ---
    n += 1
    make_content_slide(prs,
        "Real-World GNN Applications",
        ["Drug Discovery:",
         "  Stokes et al. (2020, Cell) — discovered halicin antibiotic with neural nets",
         "  Chemprop: state-of-the-art on MoleculeNet benchmarks",
         "",
         "Neuroscience:",
         "  Drosophila connectome (FlyWire: ~139K neurons, ~54.5M synapses) → graph!",
         "  GNNs predict neuron types from connectivity patterns",
         "  Brain imaging graphs: EEG/fMRI → GNN for disease classification",
         "",
         "Neuroscience & Connectomics:",
         "  Visual motion circuit: physiology + connectomics approach",
         "  Connectomics graph data (MESH repository) = neurons as graph nodes",
         "  Same GNN architecture works on molecules AND brain circuits!"],
        n,
        notes="TIMING: 59-69 min (Unit 3 start)\n\n7 min. CITATION: Stokes et al. (2020) 'A Deep Learning Approach to Antibiotic Discovery.' Cell 180:688-702. Neural network on 2,335 molecules → screened Broad repurposing library → discovered halicin.\n\n=== CONNECTOMICS & PHYSIOLOGY ===\nCITATION: Dorkenwald et al. (2024) 'Neuronal wiring diagram of an adult brain.' Nature 634:124-138. The Drosophila brain (FlyWire): ~139K neurons, ~54.5M synapses = the ultimate graph dataset.\n\nDr. Serbe-Kamp's 2016 Neuron paper: used PHYSIOLOGY (calcium imaging, electrophysiology, optogenetics) to characterize T5 neuron inputs. This is FUNCTIONAL characterization, not connectomics.\n\nDr. Serbe-Kamp CURRENTLY works on connectomics with the MESH repository — this is his ongoing research.\n\nThe connection: physiology tells you WHAT neurons do (functional), connectomics tells you HOW they're wired (structural). Both are graph problems. Same GNN architectures apply.")

    # --- Slide 9: Cross-Domain GNNs ---
    n += 1
    make_content_slide(prs,
        "Cross-Domain: Molecules, Brains, and Physiological Signals",
        ["Same GNN framework, different biological graphs:",
         "",
         "Molecular graph: atoms = nodes, bonds = edges → predict activity",
         "Brain connectome: neurons = nodes, synapses = edges → predict function",
         "Physiological signals: channels = nodes, correlations = edges → classify states",
         "",
         "Neuroscience example:",
         "  Physiology: functional characterization of T5 circuit (Serbe et al. 2016, Neuron)",
         "  Connectomics: structural graph of neural wiring (MESH repository)",
         "  Together: structure constrains function, function validates structure",
         "",
         "For your projects: BYB data (EMG/ECG/EEG) can be graphs too!",
         "  Record from multiple channels → correlation graph → GNN classification"],
        n,
        notes="TIMING: 69-77 min\n\n5 min. Key message: graph learning is a UNIVERSAL framework for biology.\n\nDr. Serbe-Kamp's research spans both:\n1. PHYSIOLOGY (2016 Neuron paper, 2023 papers): calcium imaging, electrophysiology, optogenetics — measuring FUNCTION\n2. CONNECTOMICS (MESH, current work): EM reconstruction — mapping STRUCTURE\nThese are complementary, not the same. The Neuron paper is physiology.\n\nStudent project connection: BYB recordings from multiple channels can be treated as graph data (channels as nodes, cross-correlations as edges).")

    # --- Slide 10: Practical 1 ---
    n += 1
    make_content_slide(prs,
        "Practical 1: Train a GNN with DeepChem",
        ["Open: day3_deep_learning/Day3_Practical_GNN.ipynb",
         "",
         "Steps:",
         "  1. Load MoleculeNet dataset (HIV or BACE)",
         "  2. Featurize molecules as graphs using DeepChem",
         "  3. Train a Graph Convolutional Network (GCN)",
         "  4. Compare to your Day 2 Random Forest baseline",
         "  5. Analyze: when does GNN beat RF? When doesn't it?",
         "  6. (Bonus) Try Chemprop and compare",
         "",
         "Time: ~50 min | GPU recommended (Colab provides free GPU)",
         "HIV: ~41K compounds (large), BACE: ~1.5K (small, Alzheimer's target)"],
        n,
        notes="TIMING: 77-80 min (intro), 80-130 min (practical)\n\n3 min intro. HIV: large dataset, GNN should outperform RF. BACE: small dataset, RF might win. Use SAME scaffold split from Day 2. Enable GPU in Colab: Runtime → Change runtime type → GPU.\n\nImportant: they should COMPARE to their Day 2 RF baseline. This demonstrates when deep learning helps and when it doesn't.")

    # --- Slide 11: Practical hands-on ---
    n += 1
    make_section_divider(prs, "🔬 PRACTICAL 1 — Train Your GNN", n,
        notes="TIMING: 80-130 min (Units 3-4, ~50 min coding)\n\nHands-on. Walk around. Help with GPU setup. GNN training takes longer than RF (~5-10 min vs seconds).",
        subtitle="50 minutes — molecules as graphs!")

    # --- Slide 12: Afternoon start ---
    n += 1
    make_content_slide(prs,
        "Afternoon Agenda (Units 5-8)",
        ["7. Generative AI for molecules — creating new drugs",
         "8. VAE, diffusion models, SELFIES",
         "9. Protein targets & AlphaFold for drug design",
         "10. Binding prediction & structure-based drug design",
         "11. PRACTICAL 2: Generative models or AlphaFold exploration",
         "",
         "☕ BREAK after Unit 6"],
        n,
        notes="TIMING: 130-133 min (Unit 5 start)\n\n2 min overview. Afternoon shifts from predicting properties to GENERATING new molecules and understanding protein targets.")

    # --- Slide 13: Generative AI Section ---
    n += 1
    make_section_divider(prs, "Generative AI for Molecules", n,
        notes="TIMING: 133-135 min\n\nTransition: we've been predicting properties. Now: can AI CREATE new molecules?",
        subtitle="Can AI invent new drugs?")

    # --- Slide 14: Generative Approaches ---
    n += 1
    make_content_slide(prs,
        "Generative AI: Creating New Molecules",
        ["Goal: generate novel molecules with desired properties",
         "  Not just screen existing libraries — DESIGN new compounds!",
         "",
         "Key approaches:",
         "  1. Variational Autoencoder (VAE) on SMILES or molecular graphs",
         "     Gómez-Bombarelli et al. (2018), ACS Central Science",
         "  2. Recurrent Neural Networks (RNN) on SMILES strings",
         "  3. Generative Adversarial Networks (GAN)",
         "  4. Reinforcement Learning — optimize for desired properties",
         "  5. Diffusion Models — state-of-the-art (Hoogeboom et al., ICML 2022)",
         "",
         "SELFIES (Krenn et al. 2020): 100% valid molecular strings!",
         "  Every string decodes to a valid molecule → ideal for generation"],
        n,
        notes="TIMING: 135-148 min\n\n8 min. CITATION: Gómez-Bombarelli et al. (2018) 'Automatic Chemical Design Using a Data-Driven Continuous Representation of Molecules.' ACS Central Science 4:268-276. This was the landmark paper: encode molecules to latent space, decode back. Modify the latent representation to change properties.\n\nSELFIES is important: with SMILES, randomly generated strings are usually invalid. SELFIES guarantees validity — every string is a valid molecule.\n\nDiffusion models (from image generation: DALL-E, Stable Diffusion) are now adapted for molecules.")

    # --- Slide 15: Multi-objective Optimization ---
    n += 1
    make_content_slide(prs,
        "Multi-Objective Drug Design",
        ["Real drug design requires optimizing MANY properties simultaneously:",
         "  High potency (low IC50)  |  Drug-like (Lipinski/CNS rules)",
         "  Low toxicity (no hERG block)  |  Synthetic accessibility",
         "  Good ADMET  |  Novel (patentable) |  Soluble",
         "",
         "Approaches:",
         "  Pareto optimization: no solution dominates all others",
         "  Reward shaping: combine objectives into single score",
         "  Reinforcement learning: agent learns to generate optimal molecules",
         "",
         "This is where domain expertise (neuroscience, pharmacology) is critical!",
         "  You need to know WHICH properties matter for your target"],
        n,
        notes="TIMING: 148-155 min\n\n5 min. Real drug design is multi-objective. For CNS drugs: must cross BBB AND avoid hERG block AND be selective for target. This is where neuroscience knowledge is invaluable — knowing which properties matter for neural targets.\n\nFor student projects using generative models: define 2-3 objectives and optimize simultaneously.")

    # --- Slide 16: Protein Targets Section ---
    n += 1
    make_section_divider(prs, "Protein Targets & Binding Prediction", n,
        notes="TIMING: 155-157 min\n\nTransition: we've covered ligand-based approaches (fingerprints, QSAR, GNN, generation). Now: structure-based — using 3D protein structures for drug design.",
        subtitle="Structure-based drug design with AlphaFold")

    # --- Slide 17: AlphaFold Deep Dive ---
    n += 1
    make_content_slide(prs,
        "AlphaFold2 & AlphaFold3: Revolution in Structural Biology",
        ["AlphaFold2 (Jumper et al. 2021, Nature) — 2024 Nobel Prize",
         "  Solved 50-year protein folding challenge",
         "  200+ million structures predicted in AlphaFold DB",
         "",
         "AlphaFold3 (Abramson et al. 2024, Nature):",
         "  Predicts protein-ligand, protein-DNA, protein-RNA complexes",
         "  Directly applicable to drug design!",
         "",
         "Neuroscience drug targets with AlphaFold structures:",
         "  hERG (cardiac safety)  |  GABA-A (anxiolytics, anesthetics)",
         "  GluCl (ivermectin target, Cys-loop receptor superfamily)",
         "  Serotonin receptors (5-HT2A: psychedelics, antidepressants)",
         "  Dopamine receptors (D2: antipsychotics)  |  nAChR (addiction)"],
        n,
        notes="TIMING: 157-167 min\n\n7 min. CITATION: Jumper et al. (2021) 'Highly accurate protein structure prediction with AlphaFold.' Nature 596:583-589. doi:10.1038/s41586-021-03819-2. AlphaFold2 accuracy matches experimental methods (GDT >90 on CASP14). 2024 Nobel Prize in Chemistry awarded to Demis Hassabis and John Jumper (shared with David Baker for computational protein design).\n\nCITATION: Abramson et al. (2024) 'Accurate structure prediction of biomolecular interactions with AlphaFold 3.' Nature 630:493-500. doi:10.1038/s41586-024-07487-w. AlphaFold3 is the game-changer for drug design: predicts protein-ligand complexes directly. Uses a diffusion-based architecture (different from AF2's end-to-end Evoformer approach).\n\nNeuroscience drug targets: ALL major families benefit:\n- Ion channels (hERG, Nav, GluCl, GABA-A): membrane proteins that were extremely hard to crystallize. AlphaFold structures opened up structure-based drug design for these targets.\n- GPCRs (serotonin 5-HT2A, dopamine D2, opioid receptors): ~34% of all FDA-approved drugs target GPCRs (Saikia et al. 2019, Curr. Drug Targets 20:522-539). AlphaFold structures complement cryo-EM structures.\n- Transporters (SERT, DAT): targets for SSRIs (fluoxetine/Prozac), methylphenidate (Ritalin)\n\nStudent project: pick one of these targets, retrieve AlphaFold structure from uniprot.org → AlphaFold DB, analyze binding site, compare to experimental PDB structure.")

    # --- Slide 18: Molecular Docking ---
    n += 1
    make_content_slide(prs,
        "Molecular Docking & Binding Prediction",
        ["Molecular docking: fit a molecule into a protein binding site",
         "  AutoDock Vina, GNINA, DiffDock (Corso et al. 2023, ICLR)",
         "",
         "Scoring functions: predict binding affinity from pose",
         "  Physics-based: force fields  |  Empirical: trained on data",
         "  ML-based: GNNs on protein-ligand complexes",
         "",
         "Virtual screening workflow:",
         "  1. Get protein structure (experimental or AlphaFold)",
         "  2. Prepare binding site",
         "  3. Dock library of compounds",
         "  4. Score and rank",
         "  5. Validate top hits experimentally (including physiology!)",
         "",
         "Step 5 = physiological validation = what the SpikerBox does!"],
        n,
        notes="TIMING: 167-175 min\n\n5 min. The virtual screening workflow ends with EXPERIMENTAL VALIDATION. For ion channel drugs, this means electrophysiology. This connects back to the physiology-as-readout theme.\n\nCITATION: DiffDock: Corso et al. (2023) 'DiffDock: Diffusion Steps, Twists, and Turns for Molecular Docking.' ICLR 2023. (arXiv preprint: 2022, arXiv:2210.01776.) Uses diffusion models for molecular docking — generates multiple plausible binding poses and scores them. State-of-the-art for blind docking (no prior knowledge of binding site).")

    # --- Slide 19: Break ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="TIMING: 175-190 min (after Unit 6)\n\nBreak.",
        subtitle="Last practical coming up!")

    # --- Slide 20: Practical 2 ---
    n += 1
    make_content_slide(prs,
        "Practical 2: Explore Generative Models or AlphaFold",
        ["Open: day3_deep_learning/Day3_Practical_Advanced.ipynb",
         "",
         "Choose your track (both are in the notebook):",
         "",
         "Track A — Generative Chemistry:",
         "  1. Use SELFIES to generate random valid molecules",
         "  2. Train a simple VAE on molecular data",
         "  3. Interpolate in latent space between known drugs",
         "  4. Score generated molecules with your Day 2 QSAR model",
         "",
         "Track B — AlphaFold Exploration:",
         "  1. Retrieve AlphaFold predictions for a neuro drug target",
         "  2. Visualize with py3Dmol in Colab",
         "  3. Identify binding site residues",
         "  4. Compare AlphaFold vs experimental structure (if available)",
         "",
         "Time: ~50 min | Choose based on your project interest!"],
        n,
        notes="TIMING: 190-195 min (intro) then 195-245 min (practical)\n\n3 min intro. Track A is better for students doing QSAR projects (6). Track B is better for students doing AlphaFold projects (7). BYB project students can choose either for general understanding.\n\nTrack B is particularly important: students learn to use AlphaFold DB, visualize protein structures, and identify drug binding sites. This is directly applicable to project 7.")

    # --- Slide 21: Practical hands-on ---
    n += 1
    make_section_divider(prs, "🔬 PRACTICAL 2 — Generative AI or AlphaFold", n,
        notes="TIMING: 195-245 min (Units 7-8, ~50 min coding)\n\nHands-on. Students choose their track. Walk around and help.",
        subtitle="Choose your track and explore!")

    # --- Slide 22: Discussion ---
    n += 1
    make_discussion_slide(prs,
        "GNNs beat RF on benchmarks. But when would you\nSTILL choose Random Forest for drug discovery?",
        ["Dataset size: GNN needs >10K molecules to shine",
         "Interpretability: SHAP on RF is straightforward",
         "Computational cost: RF trains in seconds, GNN in hours",
         "Neuroscience angle: small circuit (60 neurons) → RF. Full fly brain (140K) → GNN"],
        n,
        notes="TIMING: 245-255 min\n\n7-min discussion. RF wins on small data, when interpretability is critical, or when compute is limited. GNN wins on large data. CITATION: Yang et al. (2019) 'Analyzing Learned Molecular Representations for Property Prediction.' JCIM 59:3370-3388. Chemprop (D-MPNN) beat RF on 7/8 MoleculeNet tasks, but RF won on 1.\n\nNeuroscience analogy: small T5 circuit (~60 neurons) → hand-crafted features + RF. Full 140K-neuron connectome → GNN essential.")

    # --- Slide 23: Takeaways ---
    n += 1
    make_takeaway_slide(prs,
        ["GNNs learn molecular representations end-to-end from graph structure",
         "Message passing ≈ Morgan fingerprints, but LEARNED",
         "Generative AI can CREATE new drug candidates (VAE, diffusion, RL)",
         "AlphaFold: protein structure prediction → structure-based drug design",
         "Same GNN framework applies to molecules, brain circuits, physiology data",
         "Virtual screening ends with physiological validation (SpikerBox!)",
         "ALWAYS compare to RF/XGBoost baseline!",
         "Next (Day 4, 19 May): Ethics, physiology as readout, project workshop"],
        3, n,
        notes="TIMING: 255-260 min\n\nQuick recap. The key progression: Day 1 (pipeline + physiology), Day 2 (classical ML), Day 3 (deep learning + generation). Day 4 wraps up with ethics and project finalization.")

    # --- Slide 24: Next Day ---
    n += 1
    make_content_slide(prs,
        "Coming Up: Day 4 — 19 May 2026",
        ["Ethics, Regulation & Responsible AI in Drug Discovery (2 units)",
         "  EU AI Act, bias, safety, intellectual property",
         "  Discussion: David Willson's mRNA vaccine — ethics of DIY medicine",
         "",
         "Physiology as Drug Readout — The BYB Approach (1 unit)",
         "  From AI prediction to physiological validation",
         "  SpikerBox/SpikerBot demonstrations",
         "  Ion channel pharmacology: the final test of drug action",
         "",
         "Project Finalization Workshop (1 unit)",
         "  Troubleshoot your project, get feedback, prepare for exam",
         "",
         "EXAM: 27 May — presentations/posters"],
        n,
        notes="TIMING: 260-265 min\n\nPreview of Day 4 (only 4 units). Projects should be well underway by then. Encourage students to start coding their projects now.")

    os.makedirs("day3_deep_learning", exist_ok=True)
    prs.save("day3_deep_learning/slides.pptx")
    print(f"  Day 3: {n} slides → day3_deep_learning/slides.pptx")


# =====================================================================
# DAY 4 — 19 May 2026 — 4 units (180 min)
# Ethics, Physiology as Drug Readout, Project Finalization
# =====================================================================
def generate_day4():
    prs = new_prs()
    n = 0

    # --- Title ---
    n += 1
    make_title_slide(prs,
        "Ethics, Physiology & Project Workshop",
        "Responsible AI, the physiological readout of drugs, and your projects",
        4, "AI for Drug Discovery", date_str="19 May 2026",
        notes="TIMING: 0-10 min (Unit 1 start)\n\nLast teaching day before the exam! Today:\n- Units 1-2: Ethics & regulation in AI drug discovery + physiology as drug readout\n- Units 3-4: Project finalization workshop\n\nCheck project progress: every group should have their data/analysis well underway. Exam is 27 May — 8 days from now.")

    # --- Slide 2: Day Agenda ---
    n += 1
    make_content_slide(prs,
        "Today's Agenda",
        ["Unit 1 — Ethics & Regulation in AI Drug Discovery",
         "  AI bias in chemical data, FDA/EMA AI frameworks",
         "  Intellectual property: who owns an AI-designed drug?",
         "  Discussion: David Willson's mRNA vaccine — DIY medicine ethics",
         "",
         "Unit 2 — Physiology: The Last Readout of Drug Action",
         "  From AI prediction to experimental validation",
         "  Ion channel pharmacology: patch-clamp, SpikerBox, SpikerBot",
         "  Physiology research meets drug discovery",
         "",
         "☕ BREAK — 15 min",
         "",
         "Units 3-4 — Project Workshop",
         "  Group consultations, troubleshooting, exam preparation"],
        n,
        notes="TIMING: 10-12 min\n\n2 min overview. This is a compact day (only 4 units = 180 min). Ethics + physiology are front-loaded, then we shift to hands-on project work.")

    # --- Slide 3: Ethics Section ---
    n += 1
    make_section_divider(prs, "Ethics & Regulation", n,
        notes="TIMING: 12-14 min\n\nTransition to ethics. This is required content and genuinely important for responsible AI use.",
        subtitle="Responsible AI in drug discovery")

    # --- Slide 4: AI Bias in Drug Data ---
    n += 1
    make_content_slide(prs,
        "AI Bias in Drug Discovery Data",
        ["Chemical bias: historical datasets over-represent certain scaffolds",
         "  e.g., benzodiazepines dominate GABA-A data → model learns scaffold, not binding",
         "",
         "Target bias: most data on 'popular' targets (kinases, GPCRs)",
         "  Neglected tropical diseases have very little data",
         "",
         "Population bias: clinical trial data skews toward certain demographics",
         "  Pharmacogenomics: drug metabolism varies by population",
         "",
         "Publication bias: positive results published, negative results hidden",
         "  ChEMBL is biased toward active compounds",
         "",
         "Question: how do scaffold splits (Day 2) address some of these biases?"],
        n,
        notes="TIMING: 14-24 min\n\n7 min. Connect back to Day 2: scaffold splits directly address chemical bias by testing on novel scaffolds. But other biases (target, population, publication) require different solutions.\n\nFor neuroscience: CNS drugs have the highest failure rate partly because data quality for brain targets is lower and the biology is more complex.")

    # --- Slide 5: Regulation ---
    n += 1
    make_content_slide(prs,
        "Regulatory Landscape for AI in Drug Discovery",
        ["FDA:",
         "  AI/ML Action Plan (2021): 'Good Machine Learning Practice' principles",
         "  Focus on transparency, clinical validation, performance monitoring",
         "",
         "EMA (European Medicines Agency):",
         "  Reflection paper on AI (2023): emphasizes human oversight",
         "",
         "EU AI Act (2024):",
         "  Medical AI = high-risk → requires conformity assessment",
         "  Transparency, documentation, human oversight obligations",
         "",
         "Intellectual property: who owns an AI-designed drug?",
         "  Patent offices increasingly require human inventor",
         "  AI as a tool, not an inventor (for now)"],
        n,
        notes="TIMING: 24-32 min\n\n5 min. Key regulatory principle: AI is a TOOL used by human scientists. The human remains responsible. FDA wants to see that models are validated, monitored, and that there is human oversight.\n\nIP question: if AI generates a novel molecule, who is the inventor? Current law requires a human inventor. Insilico Medicine listed their human scientists, not the AI.")

    # --- Slide 6: mRNA Discussion ---
    n += 1
    make_discussion_slide(prs,
        "David Willson designed an mRNA cancer vaccine for his dog.\nIs this the future of personalized medicine or a cautionary tale?",
        ["Arguments FOR: democratization of medicine, personalized treatment",
         "Arguments AGAINST: safety risks, no clinical trial, no oversight",
         "What if the vaccine caused harm? Who is responsible?",
         "How does this relate to AI-generated drug candidates?",
         "Should there be an 'open-source medicine' movement?"],
        n,
        notes="TIMING: 32-45 min\n\n10-min discussion. This is NOT a peer-reviewed study — it's a viral news story. The underlying science (mRNA cancer vaccines, neoantigen prediction) IS well-published (BioNTech, Moderna clinical trials).\n\nKey discussion points:\n1. TECHNOLOGY: The pipeline (tumor sequencing → neoantigen prediction → mRNA design) uses AI at every step. Students should understand this.\n2. ETHICS: No IRB, no FDA, no GMP manufacturing. Safety is unknown.\n3. REGULATION: Should people be allowed to make their own medicines? Where's the line?\n4. AI CONNECTION: AI-generated drug candidates raise similar questions — how much validation is needed before human use?\n\nThis discussion is relevant for students considering the mRNA project topic.")

    # --- Slide 7: Physiology Section ---
    n += 1
    make_section_divider(prs, "Physiology: The Last Readout of Drugs", n,
        notes="TIMING: 45-47 min (Unit 2 start)\n\nTransition to the central theme of the course. This slide is about why physiology is indispensable for drug design.",
        subtitle="From AI prediction to biological reality")

    # --- Slide 8: Physiology Hierarchy ---
    n += 1
    make_content_slide(prs,
        "The Drug Readout Hierarchy",
        ["Every drug must ultimately change PHYSIOLOGY in living systems:",
         "",
         "Level 1 — Molecular: binding affinity (Ki, IC50)",
         "  → In silico prediction (QSAR, GNN, docking) — Day 2-3 content",
         "",
         "Level 2 — Cellular: physiological response (channel currents, Ca²⁺ signals)",
         "  → Patch-clamp, calcium imaging — Serbe et al. (2016) Neuron 89:829-841",
         "  → SpikerBox: extracellular recording of drug effects on neural/muscle firing",
         "",
         "Level 3 — Tissue/Organ: cardiac QT interval, neural circuit function",
         "  → ECG, EEG, EMG — BYB Human Signals project",
         "",
         "Level 4 — Organism: behavior, clinical outcome",
         "  → Eye-tracking, VR experiments — student projects 4, 5",
         "",
         "AI predicts levels 1-2. Physiology VALIDATES all levels."],
        n,
        notes="TIMING: 47-57 min\n\n7 min. This is the KEY conceptual slide that ties the course together.\n\nThe drug readout hierarchy maps directly to the course content AND the student projects:\n- Levels 1-2: AI methods from Days 2-3 predict these computationally\n- Level 2: Dr. Serbe-Kamp's 2016 Neuron paper used calcium imaging, electrophysiology, and optogenetics — these are the gold standard for measuring cellular drug effects\n- Level 3: BYB human signal recording (ECG, EEG, EMG) measures organ-level effects\n- Level 4: Eye-tracking and VR measure behavioral effects\n\nThe SpikerBox/SpikerBot make Level 2 accessible: record from cockroach leg, apply lidocaine (Nav blocker), watch action potentials disappear. This IS drug effect measurement.\n\nThis perspective justifies ALL the hardware projects:\n- Project 1 (SpikerBox): Level 2 — cellular physiology\n- Project 2 (SpikerBot): Level 2 — stimulation + recording\n- Project 3 (BYB Human): Level 3 — organ-level physiology\n- Project 4 (Eye-tracking): Level 4 — behavioral\n- Project 5 (VR): Level 4 — behavioral/perceptual")

    # --- Slide 9: Ion Channel Pharmacology ---
    n += 1
    make_content_slide(prs,
        "Ion Channel Pharmacology: Where AI Meets the SpikerBox",
        ["Ion channels = ~6% of primary drug targets (Santos et al. 2017, Nat. Rev. Drug Discov. 16:19-34)",
         "",
         "Critical ion channel drug targets:",
         "  Nav: antiepileptics (carbamazepine), local anesthetics (lidocaine)",
         "  GABA-A: anxiolytics (diazepam), anesthetics (propofol)",
         "  GluCl: antiparasitics (ivermectin) — same family as GABA-A!",
         "  hERG: cardiac safety (MANDATORY screening for ALL drugs)",
         "  nAChR: smoking cessation (varenicline)",
         "",
         "Measuring drug effects on ion channels:",
         "  Lab: patch-clamp electrophysiology (gold standard)",
         "  DIY: SpikerBox (extracellular recording, ~$100)",
         "  Computational: Hodgkin-Huxley simulation (Day 1 notebook!)",
         "",
         "GluClα research: Ammer, Serbe-Kamp et al. (2023, Nat. Neurosci. 26:1894-1905) → ivermectin target"],
        n,
        notes="TIMING: 57-67 min\n\n7 min. CITATION: Santos et al. (2017) 'A comprehensive map of molecular drug targets.' Nature Rev. Drug Discov. 16:19-34. Note: Santos reports ion channels as ~6% of primary efficacy targets, NOT 18%. The 18% figure is sometimes cited but is incorrect — it likely conflates different target categories. GPCRs are ~34% (the largest family). Ion channels remain critically important despite the smaller percentage because of safety (hERG) and neurological disease applications.\n\n=== THE GluCl / IVERMECTIN CONNECTION ===\nGluCl is a glutamate-gated chloride channel in the Cys-loop receptor superfamily (same family as GABA-A, glycine receptors, nAChR). Ivermectin locks GluCl open → Cl⁻ influx → hyperpolarization → paralysis in parasites.\n\n2015 Nobel Prize (Ōmura & Campbell) for ivermectin discovery.\n\nDr. Serbe-Kamp's 2023 Nature Neuroscience paper (Ammer, Serbe-Kamp et al., Nature Neurosci. 26:1894-1905) found GluClα mediates direction-opponent inhibition in the Drosophila visual system. Same channel family, different organism, different function.\n\n=== hERG SAFETY ===\nhERG (human Ether-à-go-go Related Gene) channel is MANDATORY for drug safety. Block of hERG can cause QT prolongation → cardiac arrhythmia → sudden death. EVERY drug candidate must be tested. CITATION: Redfern et al. (2003) 'Relationships between preclinical cardiac electrophysiology, clinical QT interval prolongation and torsade de pointes.' Cardiovasc. Res. 58:32-45. The ICH S7B guideline requires hERG testing.\n\nAI predicts hERG block from molecular structure (QSAR). Physiology validates it (patch-clamp).\n\n=== DIY APPROACH ===\nSpikerBox: apply lidocaine to cockroach leg → Nav block → no action potentials. This IS drug testing at the most fundamental level. Making it accessible is the BYB mission.")

    # --- Slide 10: Physiology Research ---
    n += 1
    make_content_slide(prs,
        "Physiology Research → Drug Discovery",
        ["Key publication: Serbe et al. (2016) Neuron 89:829-841",
         "  'Comprehensive Characterization of the Major Presynaptic Elements",
         "  to the Drosophila OFF Motion Detector'",
         "",
         "  Methods: calcium imaging + electrophysiology + optogenetics",
         "  This is a PHYSIOLOGY paper — functional characterization of neurons",
         "  NOT a connectomics paper (that's the current MESH project)",
         "",
         "Why this matters for drug discovery:",
         "  Same techniques measure drug effects on neural function",
         "  Calcium imaging → FLIPR assays for GPCR drug screens",
         "  Electrophysiology → patch-clamp for ion channel IC50",
         "  Optogenetics → precise control of neural activity",
         "",
         "Current work: MESH repository for connectomics (structural mapping)",
         "Physiology (function) + Connectomics (structure) = complete picture"],
        n,
        notes="TIMING: 67-77 min\n\n7 min. This slide explicitly clarifies the 2016 Neuron paper.\n\n=== IMPORTANT CORRECTION ===\nThe 2016 Neuron paper ('Comprehensive Characterization of the Major Presynaptic Elements to the Drosophila OFF Motion Detector', Serbe et al.) is a PHYSIOLOGY paper. It used:\n1. Calcium imaging (GCaMP6f, two-photon) to measure visual responses of medulla neurons\n2. Whole-cell electrophysiology (patch-clamp) to record membrane potential changes\n3. Optogenetics (CsChrimson, GtACR1) to activate/silence specific neuron types\n\nWhile it references EM data for anatomical context, its primary contribution is FUNCTIONAL characterization — understanding how neurons RESPOND to stimuli.\n\nDr. Serbe-Kamp is CURRENTLY working on connectomics with the MESH repository — this is a SEPARATE project focused on STRUCTURAL mapping of neural circuits.\n\nThe distinction:\n- Physiology (2016 Neuron, 2023 papers) = FUNCTION (what do neurons do?)\n- Connectomics (MESH, current) = STRUCTURE (how are neurons wired?)\n\nFor drug discovery:\n- Physiology techniques directly measure drug effects\n- Connectomics identifies potential drug targets (hub neurons, key circuit nodes)\n- Together: structure-guided target ID → physiology-based drug testing")

    # --- Slide 11: Break ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="TIMING: 77-92 min\n\nBreak. When students return, it's project workshop time.",
        subtitle="Project workshop starts after the break!")

    # --- Slide 12: Workshop Section ---
    n += 1
    make_section_divider(prs, "Project Finalization Workshop", n,
        notes="TIMING: 92-94 min (Unit 3 start)\n\nTransition to workshop mode. This is hands-on time for students to work on their projects with instructor support.",
        subtitle="Get your projects ready for the exam (27 May)")

    # --- Slide 13: Project Checklist ---
    n += 1
    make_content_slide(prs,
        "Project Checklist — What You Need for 27 May",
        ["POSTER (Projects 1-5: own data):",
         "  □ Data collected and cleaned",
         "  □ ML analysis applied (classification, regression, or clustering)",
         "  □ Results visualized (plots, figures, confusion matrices)",
         "  □ Connection to drug discovery / pharmacology explained",
         "  □ Poster formatted and ready to present",
         "",
         "PRESENTATION (Projects 6-7 or publication-based):",
         "  □ Key figures reproduced OR explained in detail",
         "  □ Methods understood (not just results)",
         "  □ Critical analysis: strengths, limitations, alternatives",
         "  □ Connection to course content (QSAR, GNN, AlphaFold)",
         "  □ 15-20 min presentation + Q&A slides ready",
         "",
         "ALL: Be ready for questions! I will probe your understanding."],
        n,
        notes="TIMING: 94-100 min\n\n4 min. Go through the checklist. The key message: depth matters more than breadth. I'd rather see a simple analysis done well and understood deeply than a complex analysis the students can't explain.\n\nFor poster projects: the minimum is data + ML analysis + interpretation + drug discovery connection.\nFor presentations: understanding HOW figures were generated (methods) is as important as WHAT they show (results).")

    # --- Slide 14: Exam Format ---
    n += 1
    make_two_column_slide(prs,
        "Exam: 27 May 2026 — Presentations / Posters",
        "Grading Criteria",
        ["Implementation (40%)",
         "  Is the analysis correct and complete?",
         "  Is the code well-documented?",
         "",
         "Scientific Reasoning (25%)",
         "  Do you understand the biology?",
         "  Can you explain your choices?",
         "",
         "Evaluation (20%)",
         "  Did you evaluate properly?",
         "  Scaffold splits? Cross-validation?"],
        "Format & Logistics",
        ["Presentation Quality (15%)",
         "  Clear, well-structured, within time",
         "",
         "POSTER: printed or digital, 5 min pitch + Q&A",
         "PRESENTATION: 15-20 min + 5-10 min Q&A",
         "",
         "Groups of 3-4, all members must contribute",
         "",
         "Publications in: publications/ folder",
         "  (organized by topic, upload your PDFs)"],
        n,
        notes="TIMING: 100-107 min\n\n5 min. Be explicit about grading. Implementation 40% emphasizes that the code must work. Scientific reasoning 25% means they must understand the biology behind their analysis. Evaluation 20% checks proper methodology. Presentation 15% rewards clear communication.\n\nAll group members must participate in the presentation/poster defense. I will ask individual questions.")

    # --- Slide 15: Workshop Time ---
    n += 1
    make_section_divider(prs, "🔧 WORKSHOP — Work on Your Projects!", n,
        notes="TIMING: 107-165 min (Units 3-4, ~58 min of workshop)\n\nThis is dedicated project time. Circulate among groups, answer questions, troubleshoot code. Check on each group at least once.\n\nCommon issues to anticipate:\n- Data loading problems (ChEMBL, BYB data formats)\n- Feature engineering choices\n- Evaluation methodology (remind: scaffold split!)\n- AlphaFold structure retrieval\n- SpikerBox data analysis pipeline\n\nEncourage students to use all the course notebooks as templates.",
        subtitle="I'm available for questions and troubleshooting!")

    # --- Slide 16: Closing ---
    n += 1
    make_takeaway_slide(prs,
        ["The 4-day journey: pipeline → molecular ML → deep learning → ethics/physiology",
         "Physiology is the FINAL readout of every drug — from patch-clamp to SpikerBox",
         "AI predicts, physiology validates — both are essential",
         "Serbe et al. (2016) Neuron paper = physiology of cells (not connectomics!)",
         "  Current work: MESH repository for connectomics",
         "Ethics matters: bias, regulation, responsible innovation",
         "Your projects bridge AI methods with biological reality",
         "",
         "EXAM: 27 May 2026 — presentations/posters",
         "Good luck! 🎓"],
        4, n,
        notes="TIMING: 165-175 min\n\nFinal recap. Thank the students for their engagement over the 4 days. Remind: exam 27 May. Publications are in the repository. Reach out with questions.\n\nThe 4-day arc:\nDay 1: What is drug discovery? Where does AI fit? What projects will you do?\nDay 2: How to represent molecules and build ML models\nDay 3: Deep learning, GNNs, generative AI, AlphaFold\nDay 4: Ethics, physiology as readout, project finalization\n\nThe key insight: AI and physiology are complementary. AI predicts, physiology validates. The SpikerBox/SpikerBot make that validation accessible.")

    # --- Slide 17: Thank you ---
    n += 1
    make_section_divider(prs, "Thank You & Good Luck! 🎓", n,
        notes="TIMING: 175-180 min (end)\n\nFinal slide. Open for last questions. See everyone on 27 May!",
        subtitle="Exam: 27 May 2026 — See you there!")

    os.makedirs("day4_ethics_physiology", exist_ok=True)
    prs.save("day4_ethics_physiology/slides.pptx")
    print(f"  Day 4: {n} slides → day4_ethics_physiology/slides.pptx")


# =====================================================================
# QUIZ — Introductory Drug Discovery Quiz (10 questions, 20 slides)
# Broad overview: drug types, markets, history, routes, targets, etc.
# =====================================================================
def generate_quiz():
    prs = new_prs()
    n = 0

    # --- Title ---
    n += 1
    make_title_slide(prs,
        "Drug Discovery Quiz",
        "10 questions to test your knowledge — from ancient remedies to billion-dollar blockbusters",
        1, "AI for Drug Discovery", date_str="15 April 2026",
        notes="TIMING: Use this quiz at the very start of Day 1, BEFORE the lecture slides. It takes about 20-30 min. The quiz serves as an icebreaker and gives students a broad overview of the drug discovery landscape. Read each question aloud, give students 30 seconds to think, then advance to the answer slide. Encourage discussion after each answer.")

    # ===== Q1: Highest-grossing drug =====
    n += 1
    make_quiz_question_slide(prs, 1,
        "What is the highest-grossing drug of all time by total revenue?",
        ["Lipitor (atorvastatin) — cholesterol",
         "Humira (adalimumab) — autoimmune diseases",
         "Keytruda (pembrolizumab) — cancer",
         "Viagra (sildenafil) — erectile dysfunction"],
        n,
        notes="Let students guess. Most will say Lipitor or Viagra. This question introduces the concept of blockbuster drugs (>$1B/year revenue). Give them 30 seconds to think.")

    n += 1
    make_quiz_answer_slide(prs, 1,
        "What is the highest-grossing drug of all time by total revenue?",
        ["Lipitor (atorvastatin) — cholesterol",
         "Humira (adalimumab) — autoimmune diseases",
         "Keytruda (pembrolizumab) — cancer",
         "Viagra (sildenafil) — erectile dysfunction"],
        1,  # B = Humira
        "✓ Humira (adalimumab, AbbVie) — ~$230 billion lifetime revenue, peak $21.2B/year (2022).\nLipitor was the previous record holder at ~$125-170B. Humira treats rheumatoid arthritis, Crohn's disease, psoriasis. It is a monoclonal antibody — a biologic, not a small molecule!",
        n,
        notes="CITATION: Statista (2024) 'Top pharmaceutical products by lifetime sales worldwide.' AbbVie annual reports confirm Humira's cumulative sales exceeded $200B by 2023. Peak annual sales: $21.2B in 2022.\n\nEXPLAIN IN DETAIL:\n- Humira (adalimumab) is a monoclonal antibody that blocks TNF-alpha, a key inflammatory cytokine. It was approved in 2002 and became the world's best-selling drug.\n- A 'blockbuster drug' is defined as one earning >$1 billion per year in revenue. Humira earned >$20B/year at its peak.\n- Lipitor (atorvastatin, Pfizer) was the previous record holder — a statin that lowers cholesterol by inhibiting HMG-CoA reductase. Lifetime ~$125-170B.\n- KEY DISTINCTION: Humira is a biologic (large protein molecule, injected), Lipitor is a small molecule (taken orally as a pill). This distinction matters for AI drug discovery — designing small molecules vs. biologics requires very different approaches.\n- Humira's patent expired in 2023 (EU) and biosimilars are now available, which is why its sales are declining.\n- Viagra (sildenafil) is iconic but lifetime revenue is only ~$30-40B — much less than Humira or Lipitor.")

    # ===== Q2: First drug in human history =====
    n += 1
    make_quiz_question_slide(prs, 2,
        "What is considered the oldest drug used by humans?",
        ["Aspirin (from willow bark)",
         "Opium (from poppies)",
         "Alcohol (from fermentation)",
         "Penicillin (from mold)"],
        n,
        notes="This question explores the deep history of pharmacology. Many students may say aspirin. Let them discuss for a moment.")

    n += 1
    make_quiz_answer_slide(prs, 2,
        "What is considered the oldest drug used by humans?",
        ["Aspirin (from willow bark)",
         "Opium (from poppies)",
         "Alcohol (from fermentation)",
         "Penicillin (from mold)"],
        1,  # B = Opium
        "✓ Opium — used since ~3400 BCE by Sumerians (called the 'joy plant').\nWillow bark (salicin → aspirin) was used from ~1500 BCE. Alcohol from ~7000 BCE as beverage, but opium was specifically used as a MEDICINE. Penicillin was discovered in 1928.",
        n,
        notes="CITATION: Brownstein, M.J. (1993) 'A brief history of opiates, opioid peptides, and opioid receptors.' Proc. Natl. Acad. Sci. 90:5391-5393. Archaeological evidence from Sumerian clay tablets (~3400 BCE) describes opium poppy cultivation.\n\nEXPLAIN IN DETAIL:\n- Opium is extracted from Papaver somniferum (the opium poppy). The Sumerians in Mesopotamia (modern Iraq) used it for pain relief and called it 'hul gil' — the 'joy plant.'\n- The active compound is morphine, which was isolated in 1804 by Friedrich Sertürner — this was the first alkaloid ever isolated from a plant, marking the birth of modern pharmacology.\n- TRICKY POINT: Alcohol (ethanol) from fermentation dates to ~7000 BCE, but it was primarily used as a beverage, not as a medicine. If we define 'drug' as a substance intentionally used for therapeutic effect, opium is the clear winner.\n- Willow bark (salicin, precursor to aspirin) was used by Egyptians and Hippocrates (~400 BCE) for pain/fever. Aspirin (acetylsalicylic acid) was synthesized by Felix Hoffmann at Bayer in 1897.\n- Penicillin: discovered 1928 by Alexander Fleming, first clinical use 1941. Much more recent.\n- KEY TAKEAWAY: Humans have been doing 'drug discovery' for over 5,000 years. What's changed is that we now understand WHY these drugs work (opium → opioid receptors, willow bark → COX inhibition). AI helps us find new drugs faster by understanding these mechanisms computationally.")

    # ===== Q3: Drug target families =====
    n += 1
    make_quiz_question_slide(prs, 3,
        "Which protein family is targeted by the largest fraction of FDA-approved drugs?",
        ["Ion channels (~15%)",
         "Kinases (~12%)",
         "Nuclear receptors (~8%)",
         "GPCRs (~34%)"],
        n,
        notes="This question teaches students about drug target families. Most AI/bioinformatics students may not know what a GPCR is. This is a great teaching moment.")

    n += 1
    make_quiz_answer_slide(prs, 3,
        "Which protein family is targeted by the largest fraction of FDA-approved drugs?",
        ["Ion channels (~15%)",
         "Kinases (~12%)",
         "Nuclear receptors (~8%)",
         "GPCRs (~34%)"],
        3,  # D = GPCRs
        "✓ GPCRs — ~34% of ALL FDA-approved drugs target G protein-coupled receptors.\nExamples: beta-blockers (heart), antihistamines (allergies), opioids (pain), SSRIs (depression). GPCRs are the single largest drug target family in pharmacology.",
        n,
        notes="CITATION: Saikia et al. (2019) 'Established and In-trial GPCR Families in Clinical Trials: A Review for Target Selection.' Curr. Drug Targets 20:522-539. Also: Hauser et al. (2017) 'Trends in GPCR drug discovery: new agents, targets and indications.' Nature Rev. Drug Discov. 16:829-842.\n\nEXPLAIN IN DETAIL:\n- GPCRs (G Protein-Coupled Receptors) are a superfamily of ~800 membrane proteins in humans. They have 7 transmembrane domains and signal through G proteins.\n- ~34% of FDA-approved drugs target GPCRs on ~108 unique GPCR targets. This makes them by far the most 'druggable' protein family.\n- Examples students should know: beta-2 adrenergic receptor (salbutamol for asthma), histamine H1 receptor (cetirizine for allergies), mu-opioid receptor (morphine for pain), serotonin receptors (fluoxetine/Prozac for depression), dopamine D2 receptor (haloperidol for schizophrenia).\n- Ion channels are ~15% (important for neurology/cardiology — hERG, Nav, GABA-A).\n- Kinases are ~12% (mostly oncology — imatinib, erlotinib).\n- Nuclear receptors are ~8% (steroid hormones — estrogen receptor, glucocorticoid receptor).\n- WHY THIS MATTERS FOR AI: AlphaFold can now predict GPCR structures. Combined with molecular docking and GNNs, AI can design new GPCR drugs computationally. This is a huge opportunity.")

    # ===== Q4: Most prescribed drug =====
    n += 1
    make_quiz_question_slide(prs, 4,
        "What is the most prescribed drug in the world (by number of prescriptions)?",
        ["Metformin (diabetes)",
         "Amoxicillin (antibiotic)",
         "Atorvastatin (cholesterol)",
         "Omeprazole (stomach acid)"],
        n,
        notes="This question distinguishes between 'most revenue' (Humira) and 'most prescribed' (volume). Very different answers!")

    n += 1
    make_quiz_answer_slide(prs, 4,
        "What is the most prescribed drug in the world (by number of prescriptions)?",
        ["Metformin (diabetes)",
         "Amoxicillin (antibiotic)",
         "Atorvastatin (cholesterol)",
         "Omeprazole (stomach acid)"],
        2,  # C = Atorvastatin
        "✓ Atorvastatin (Lipitor) — ~185 million prescriptions/year in the US alone.\nFollowed by levothyroxine (thyroid), lisinopril (blood pressure), metformin (diabetes). These are all chronic disease drugs — taken daily for life.",
        n,
        notes="CITATION: Oregon DFR (2024) '2024 Insurer Reporting Prescription Drug List.' Also: ClinCalc DrugStats database (clinicalc.com/DrugStats) tracks US prescription volumes.\n\nEXPLAIN IN DETAIL:\n- IMPORTANT DISTINCTION: 'Most prescribed' (by volume) ≠ 'highest revenue.' Atorvastatin is prescribed the most but is now generic and cheap (~$10/month). Humira earns the most revenue because each dose costs ~$5,000-7,000.\n- Atorvastatin is an HMG-CoA reductase inhibitor (statin). It lowers LDL cholesterol and is prescribed for cardiovascular disease prevention. It was originally Pfizer's Lipitor — the highest-grossing drug before Humira.\n- Metformin is the first-line treatment for type 2 diabetes. It works by decreasing hepatic glucose production and increasing insulin sensitivity. It is also being studied for anti-aging effects!\n- The top 5 most prescribed drugs are ALL for chronic conditions: cholesterol (atorvastatin), thyroid (levothyroxine), blood pressure (lisinopril), ADHD (amphetamine salts), diabetes (metformin). These represent the biggest disease burdens globally.\n- WHY THIS MATTERS FOR AI: These drugs target well-understood mechanisms with huge datasets (millions of patient records, thousands of compounds tested). This is where AI/QSAR models have the most training data and highest impact.")

    # ===== Q5: Route of administration =====
    n += 1
    make_quiz_question_slide(prs, 5,
        "What percentage of all medications are administered orally (by mouth)?",
        ["~30%",
         "~50%",
         "~60%",
         "~80%"],
        n,
        notes="This question highlights why oral bioavailability (Lipinski's Rule of Five) matters so much in drug design.")

    n += 1
    make_quiz_answer_slide(prs, 5,
        "What percentage of all medications are administered orally (by mouth)?",
        ["~30%",
         "~50%",
         "~60%",
         "~80%"],
        3,  # D = ~80%
        "✓ ~80-90% of all medications are taken orally. This is why Lipinski's Rule of Five\n(MW ≤ 500, LogP ≤ 5, HBD ≤ 5, HBA ≤ 10) is so central to drug design — it predicts oral absorption.",
        n,
        notes="CITATION: Lipinski et al. (1997) 'Experimental and computational approaches to estimate solubility and permeability in drug discovery.' Adv. Drug Deliv. Rev. 23:3-25. Also: Homayun et al. (2019) 'Challenges and Recent Progress in Oral Drug Delivery.' Pharmaceutics 11(3):129.\n\nEXPLAIN IN DETAIL:\n- The oral route dominates because it is the simplest, cheapest, and most convenient for patients. Patient compliance is highest for oral drugs.\n- But oral drugs face major hurdles: they must survive stomach acid (pH 1-3), be absorbed through the intestinal wall, survive first-pass metabolism in the liver, and reach the target tissue at sufficient concentration.\n- This is why Lipinski's Rule of Five (Ro5) is so important: it defines the physicochemical properties that predict oral absorption. We will learn this on Day 2.\n- Other routes: injection (~10-20%, for drugs that can't survive the gut or need rapid action — e.g., insulin, antibodies like Humira), topical (~2-5%, skin creams), transdermal (<2%, patches like nicotine/fentanyl), inhalation (asthma inhalers), intrathecal (into spinal fluid for CNS drugs).\n- IMPORTANT FOR AI: Most QSAR models predict properties relevant to oral drugs (solubility, permeability, LogP). If your target is a CNS drug, you also need blood-brain barrier (BBB) penetration — an additional filter beyond Ro5.")

    # ===== Q6: Animal vs human pharma market =====
    n += 1
    make_quiz_question_slide(prs, 6,
        "How does the global veterinary drug market compare to the human pharmaceutical market?",
        ["Veterinary is about 50% of human pharma",
         "Veterinary is about 25% of human pharma",
         "Veterinary is about 10% of human pharma",
         "Veterinary is about 3% of human pharma"],
        n,
        notes="Students often overestimate the veterinary market. This question provides perspective on market sizes.")

    n += 1
    make_quiz_answer_slide(prs, 6,
        "How does the global veterinary drug market compare to the human pharmaceutical market?",
        ["Veterinary is about 50% of human pharma",
         "Veterinary is about 25% of human pharma",
         "Veterinary is about 10% of human pharma",
         "Veterinary is about 3% of human pharma"],
        3,  # D = ~3%
        "✓ The veterinary pharmaceutical market (~$53B) is only about 3% of the human pharmaceutical market (~$1.58 trillion). However, veterinary drug discovery is growing fast, especially for companion animals (pets).",
        n,
        notes="CITATION: Straits Research (2024) 'Veterinary Pharmaceutical Drugs Market Size.' Reports veterinary pharma at ~$53B in 2024. IQVIA/Statista (2024) reports human pharma at ~$1.58 trillion.\n\nEXPLAIN IN DETAIL:\n- Human pharmaceutical market (2024): ~$1.58 trillion. This is one of the largest industries on Earth.\n- Veterinary pharmaceutical market (2024): ~$53 billion. About 3.3% of human pharma.\n- BUT veterinary pharma is growing faster (~8-9% CAGR vs. ~5-6% for human pharma), driven by the pet care boom — people increasingly treat pets like family members and spend more on their healthcare.\n- The veterinary market breaks down into: livestock/production animals (~55%) vs. companion animals/pets (~45%). The pet segment is growing fastest.\n- INTERESTING OVERLAP: Many drugs work in both humans and animals! Ivermectin (antiparasitic), gabapentin (pain), fluoxetine (anxiety in dogs), and many antibiotics are used across species. The 'One Health' concept recognizes that human, animal, and environmental health are interconnected.\n- CONNECTION TO THE COURSE: The David Willson mRNA vaccine story (Day 1 slide 7) is a veterinary application. AI drug discovery methods apply equally to both markets. Some companies (like Invetx/Dechra) are using AI specifically for veterinary drug discovery.\n- REGULATORY DIFFERENCE: FDA Center for Veterinary Medicine (CVM) handles animal drugs. Approval is generally faster and cheaper than for human drugs, making it potentially a good entry point for AI-discovered drugs.")

    # ===== Q7: Drug that saved most lives =====
    n += 1
    make_quiz_question_slide(prs, 7,
        "Which drug is estimated to have saved the most human lives in history?",
        ["Insulin",
         "Penicillin",
         "Chloroquine",
         "Oral Rehydration Salts (ORS)"],
        n,
        notes="This is a thought-provoking question. All four have saved millions. Let students debate.")

    n += 1
    make_quiz_answer_slide(prs, 7,
        "Which drug is estimated to have saved the most human lives in history?",
        ["Insulin",
         "Penicillin",
         "Chloroquine",
         "Oral Rehydration Salts (ORS)"],
        1,  # B = Penicillin
        "✓ Penicillin — estimated to have saved over 200 million lives since its discovery.\nDiscovered by Alexander Fleming (1928), first clinical use 1941. Antibiotics collectively prevent ~200K deaths/year in the US alone. ORS has saved ~70 million lives (primarily children with diarrheal disease).",
        n,
        notes="CITATION: Science History Institute: 'To date, penicillin has saved an estimated 200 million lives worldwide.' Guinness World Records lists Penicillium as the 'most life-saving fungi.'\n\nEXPLAIN IN DETAIL:\n- Penicillin was discovered by Alexander Fleming in 1928 when he noticed that Penicillium mold killed bacteria on a petri dish. But it wasn't developed into a usable drug until Howard Florey and Ernst Boris Chain purified it in 1940-41 (all three shared the 1945 Nobel Prize in Physiology or Medicine).\n- Before penicillin, a simple scratch could be lethal if it got infected. Bacterial pneumonia, wound infections, strep throat — all were potential death sentences. Penicillin transformed medicine.\n- HOW IT WORKS: Penicillin inhibits bacterial cell wall synthesis by blocking transpeptidase (penicillin-binding proteins, PBPs). Bacteria can't maintain their cell wall → they lyse and die. Human cells don't have cell walls, so penicillin is selectively toxic to bacteria.\n- ORS is also an incredible answer — it has saved ~70 million lives by treating dehydration from diarrheal diseases, primarily in children in developing countries. It's the simplest 'drug' imaginable: water + salt + sugar.\n- Chloroquine has saved tens of millions from malaria. Insulin has saved millions of Type 1 diabetics since 1922.\n- KEY TAKEAWAY: The most impactful drugs are often simple molecules with well-understood mechanisms. AI can help find the NEXT penicillin by screening billions of compounds against novel bacterial targets — especially important given rising antibiotic resistance (see Stokes et al. 2020, Cell — halicin discovery).")

    # ===== Q8: Average time for drug development =====
    n += 1
    make_quiz_question_slide(prs, 8,
        "On average, how long does it take to develop a new drug from initial discovery to FDA approval?",
        ["3-5 years",
         "6-8 years",
         "10-15 years",
         "20-25 years"],
        n,
        notes="This sets up the motivation for the entire course: why we need AI to speed up drug discovery.")

    n += 1
    make_quiz_answer_slide(prs, 8,
        "On average, how long does it take to develop a new drug from initial discovery to FDA approval?",
        ["3-5 years",
         "6-8 years",
         "10-15 years",
         "20-25 years"],
        2,  # C = 10-15 years
        "✓ 10-15 years on average (median ~12 years). Costs ~$2.6 billion per approved drug.\nThis is why AI is so valuable: Insilico Medicine reached Phase II in ~30 months using AI for both target identification and drug design.",
        n,
        notes="CITATION: DiMasi et al. (2016) 'Innovation in the pharmaceutical industry: New estimates of R&D costs.' J. Health Economics 47:20-33. Reports $2.558B out-of-pocket ($2.87B capitalized). Also: BIO/Informa/QLS (2021) 'Clinical Development Success Rates 2011-2020' for the ~7.9% overall success rate.\n\nEXPLAIN IN DETAIL:\n- The breakdown: Target ID & Validation (1-2 years) → Lead Discovery & Optimization (2-3 years) → Preclinical (1-2 years) → Phase I safety (1-2 years) → Phase II efficacy (2-3 years) → Phase III large-scale (3-4 years) → FDA Review (1-2 years).\n- Total: 10-15 years. Average cost per approved drug: ~$2.6 billion. This includes the cost of ALL the failed drugs along the way (>90% failure rate).\n- WHY AI MATTERS: Insilico Medicine's ISM001-055 (Rentosertib) went from project start to Phase IIa enrollment in ~30 months — compressing the first 5-6 years of the pipeline. CITATION: Ren et al. (2024) Nature Biotechnology. doi:10.1038/s41587-024-02143-0.\n- BUT: AI primarily accelerates the EARLY stages (target ID, virtual screening, lead optimization). Clinical trials still take 5-10 years because you have to test in humans and that takes time for safety reasons.\n- The $2.6B figure is controversial — some argue it's inflated by opportunity cost calculations. But even conservative estimates put the cost at >$1B per approved drug.\n- KEY MESSAGE FOR THE COURSE: AI won't eliminate the need for clinical trials. But it can dramatically reduce the cost and time of the preclinical stages, and improve the success rate by better predicting which drugs will fail.")

    # ===== Q9: First 'magic bullet' drug =====
    n += 1
    make_quiz_question_slide(prs, 9,
        "Paul Ehrlich coined the concept of a 'magic bullet' — a drug that selectively targets disease. What was the first magic bullet drug?",
        ["Aspirin (1897) — targeting pain/inflammation",
         "Salvarsan (1910) — targeting syphilis bacteria",
         "Penicillin (1941) — targeting bacterial infections",
         "Methotrexate (1947) — targeting cancer cells"],
        n,
        notes="This question introduces the concept of TARGETED therapy — the foundation of modern rational drug design and what AI tries to optimize.")

    n += 1
    make_quiz_answer_slide(prs, 9,
        "What was the first 'magic bullet' drug?",
        ["Aspirin (1897) — targeting pain/inflammation",
         "Salvarsan (1910) — targeting syphilis bacteria",
         "Penicillin (1941) — targeting bacterial infections",
         "Methotrexate (1947) — targeting cancer cells"],
        1,  # B = Salvarsan
        "✓ Salvarsan (arsphenamine, compound 606) — discovered by Paul Ehrlich & Sahachiro Hata in 1909, clinical use 1910. First synthetic drug designed to selectively kill a pathogen (Treponema pallidum, syphilis). Ehrlich tested 605 compounds before finding the right one!",
        n,
        notes="CITATION: Ehrlich, P. (1913) 'Address in Pathology on Chemotherapeutics: Scientific Principles, Methods and Results.' The Lancet 182:445-451. Also: Strebhardt & Ullrich (2008) 'Paul Ehrlich's magic bullet concept: 100 years of progress.' Nature Rev. Cancer 8:473-480.\n\nEXPLAIN IN DETAIL:\n- Paul Ehrlich (1854-1915) is the father of chemotherapy. He proposed the 'Zauberkugel' (magic bullet) concept: a chemical that selectively kills disease-causing organisms without harming the patient.\n- He systematically tested hundreds of arsenic-based compounds against Treponema pallidum (the syphilis bacterium). Compound number 606 worked — hence the name 'Salvarsan' ('that which saves by arsenic').\n- This was the FIRST example of rational, systematic drug screening — the precursor to modern high-throughput screening (HTS) and virtual screening. Ehrlich tested 606 compounds manually. Today, AI can screen billions of virtual compounds in silico.\n- The concept of the 'magic bullet' is the FOUNDATION of targeted drug design: find a molecule that binds specifically to a disease target (receptor, enzyme, ion channel) without hitting 'off-targets' that cause side effects.\n- Ehrlich won the 1908 Nobel Prize in Physiology or Medicine for his work on immunity. His chemotherapy work came after the Nobel Prize.\n- CONNECTION TO AI: Modern drug discovery is essentially Ehrlich's approach scaled up with AI. Instead of testing 606 compounds by hand, we use QSAR models and GNNs to predict which of 10^60 possible molecules will be the best 'magic bullet' for a given target. Same concept, exponentially more powerful.")

    # ===== Q10: Therapeutic area =====
    n += 1
    make_quiz_question_slide(prs, 10,
        "Which therapeutic area currently has the most new drug approvals per year (FDA)?",
        ["Cardiovascular disease",
         "Infectious disease",
         "Oncology (cancer)",
         "Neurology / CNS disorders"],
        n,
        notes="Final question. This shows where the industry is currently investing the most.")

    n += 1
    make_quiz_answer_slide(prs, 10,
        "Which therapeutic area currently has the most new drug approvals per year (FDA)?",
        ["Cardiovascular disease",
         "Infectious disease",
         "Oncology (cancer)",
         "Neurology / CNS disorders"],
        2,  # C = Oncology
        "✓ Oncology — ~25-30% of all new FDA approvals in 2023-2024 were cancer drugs.\n50 novel drugs approved by FDA in 2024, with oncology leading by a wide margin. Neurology/CNS is second. Cardiovascular was historically dominant but now has fewer novel approvals.",
        n,
        notes="CITATION: FDA CDER (2024) 'Novel Drug Approvals for 2024' — 50 novel drugs approved. Also: Mullard (2025) 'FDA approvals.' Nature Rev. Drug Discov. (annual review).\n\nEXPLAIN IN DETAIL:\n- Oncology dominates new drug approvals because: (1) cancer is a leading cause of death globally, (2) precision medicine/genomics has enabled targeted therapies, (3) immune checkpoint inhibitors (like Keytruda) revolutionized treatment, (4) FDA has accelerated approval pathways for cancer drugs (breakthrough therapy designation, accelerated approval).\n- In 2024: ~13-15 of 50 novel FDA approvals were oncology drugs. Neurology/CNS was second with ~6-8 approvals.\n- HISTORICAL SHIFT: In the 1990s-2000s, cardiovascular drugs dominated (statins, ACE inhibitors, ARBs). But most cardiovascular targets are now well-served by generics, so pharma has shifted investment to oncology where there's more unmet need and higher prices.\n- CNS drugs have the HIGHEST failure rate of any therapeutic area (~97% failure rate in clinical trials for Alzheimer's drugs). This is because: the blood-brain barrier (BBB) blocks most drugs, brain biology is poorly understood, and clinical endpoints are harder to measure.\n- WHY AI MATTERS: AI is particularly valuable for oncology (massive genomics datasets, precision medicine) and for CNS (where traditional approaches have failed, AI might find new targets and BBB-penetrant molecules). Both are active areas of AI drug discovery research.\n- CONNECTION TO THE COURSE: In Day 3, we discuss AlphaFold structures for CNS targets (GPCRs, ion channels) and how AI can predict BBB penetration (Day 2 QSAR models). The challenge of CNS drug discovery motivates better AI methods.")

    # --- Closing slide ---
    n += 1
    make_section_divider(prs, "Ready for the Course? Let's Dive In! 🚀", n,
        notes="TIMING: End of quiz.\n\nRecap: We've covered drug history, market sizes, targets, timelines, and therapeutic areas. All of these topics will come up again during the 4-day course. Now let's start with the Day 1 lecture!",
        subtitle="Now you know more about drugs than most people — time to learn how AI designs them!")

    os.makedirs("quiz_intro", exist_ok=True)
    prs.save("quiz_intro/slides.pptx")
    print(f"  Quiz: {n} slides → quiz_intro/slides.pptx")



# =====================================================================
# QUIZ 2 — Abbreviations & Drug Biology (easier, for CS students)
# Explains core terminology: QSAR, SMILES, ECFP, ADMET, IC50, BBB, etc.
# =====================================================================
def generate_quiz_abbreviations():
    prs = new_prs()
    n = 0

    n += 1
    make_title_slide(prs,
        "Drug Discovery: Abbreviations & Biology Quiz",
        "10 questions on key terms — from QSAR to receptors to the BBB",
        1, "AI for Drug Discovery", date_str="22 April 2026",
        notes="Use this quiz at the start of Day 2 AFTER the physiology intro and BEFORE the molecular ML lecture. It takes about 20-25 minutes. Questions are deliberately easier than Quiz 1 — the goal is to introduce and reinforce terminology that CS students will encounter all day. Each answer slide has a detailed explanation that teaches the concept from first principles.")

    # Q1: SMILES
    n += 1
    make_quiz_question_slide(prs, 1,
        "What does the abbreviation 'SMILES' stand for?",
        ["Systematic Molecular Identification and Labeling Entry System",
         "Simplified Molecular Input Line Entry System",
         "Structured Molecular Index with Linked Entries System",
         "Standard Method for Indexing Large Element Sets"],
        n,
        notes="Most CS students guess 'Simplified' correctly if they've seen it on Day 1. If not, this is a great vocabulary intro.")

    n += 1
    make_quiz_answer_slide(prs, 1,
        "What does the abbreviation 'SMILES' stand for?",
        ["Systematic Molecular Identification and Labeling Entry System",
         "Simplified Molecular Input Line Entry System",
         "Structured Molecular Index with Linked Entries System",
         "Standard Method for Indexing Large Element Sets"],
        1,
        "✓ SMILES = Simplified Molecular Input Line Entry System.\nInvented by David Weininger (Daylight Chemical, 1988). Converts a 2D chemical structure into a text string. Essential for all molecular ML — it's how we give molecules to a computer.",
        n,
        notes="CITATION: Weininger, D. (1988) 'SMILES, a chemical language and information system. 1. Introduction to methodology and encoding rules.' J. Chem. Inf. Comput. Sci. 28:31-36.\n\nEXPLAIN IN DETAIL:\n- Before SMILES, there was no good standard way to put molecules into computers. Chemists drew structures by hand or used complicated graph formats.\n- Weininger designed SMILES to be human-readable AND machine-parseable. Key insight: if you do a depth-first traversal of the molecular graph and write down each atom and bond you encounter, you get a compact linear string.\n- Example: Ethanol = CCO (a chain of 2 carbons and an oxygen). Aspirin = CC(=O)Oc1ccccc1C(=O)O (more complex: an acetyl group, an ester oxygen, a benzene ring, and a carboxylic acid).\n- The 'Simplified' part: SMILES doesn't capture all stereochemistry by default, and it ignores most hydrogen atoms (they are implicit). A 'Full' version with explicit hydrogens and stereochemistry is InChI (IUPAC standard).\n- WHY IT MATTERS FOR ML: Any ML model that works with molecules needs a way to represent them. SMILES = the standard text input. From SMILES, we compute fingerprints (today's lecture), graph representations (Day 3), or feed directly into language models (SMILES as a 'molecular language').")

    # Q2: QSAR
    n += 1
    make_quiz_question_slide(prs, 2,
        "What does 'QSAR' stand for in drug discovery?",
        ["Quantitative Structure-Activity Relationship",
         "Quality-Scaled Assay Results",
         "Quick Sequence Alignment and Ranking",
         "Qualified Sample Analysis Report"],
        n,
        notes="QSAR is the central concept of today's lecture. Students should know this from Day 1 but may not remember the full expansion.")

    n += 1
    make_quiz_answer_slide(prs, 2,
        "What does 'QSAR' stand for in drug discovery?",
        ["Quantitative Structure-Activity Relationship",
         "Quality-Scaled Assay Results",
         "Quick Sequence Alignment and Ranking",
         "Qualified Sample Analysis Report"],
        0,
        "✓ QSAR = Quantitative Structure-Activity Relationship.\nFoundational concept: molecular structure determines biological activity. QSAR models learn: f(structure) → activity. Originated with Hansch (1964). Today's entire lecture is about building and evaluating QSAR models.",
        n,
        notes="CITATION: Hansch, C. & Fujita, T. (1964) 'ρ-σ-π Analysis. A Method for the Correlation of Biological Activity and Chemical Structure.' J. Am. Chem. Soc. 86:1616-1626. This 1964 paper by Corwin Hansch is the founding paper of computational medicinal chemistry.\n\nEXPLAIN IN DETAIL:\n- The core hypothesis of QSAR: if you know the STRUCTURE of a molecule (its atoms, bonds, shape, electronic properties), you can PREDICT its biological ACTIVITY (how strongly it binds to a protein, how toxic it is, how soluble it is).\n- This is actually a deep biological insight: drugs work by binding to proteins. Binding depends on shape complementarity and chemical compatibility (hydrogen bonds, hydrophobic contacts, electrostatics). Both the drug shape/chemistry AND the protein shape/chemistry are determined by molecular structure. Therefore structure → activity.\n- For CS students: QSAR is essentially supervised learning with molecules as inputs and activity values as outputs. The main challenge is: how do you represent a molecule as a feature vector? (Answer: fingerprints and descriptors — today's lecture.)\n- QUANTITATIVE: not just active/inactive (binary), but the actual IC50 or logS or binding affinity value. This allows regression, not just classification.\n- EXAMPLES: Predict aqueous solubility (today's practical), predict hERG channel blocking, predict BBB penetration, predict metabolic stability — all QSAR tasks.")

    # Q3: ADMET
    n += 1
    make_quiz_question_slide(prs, 3,
        "What does 'ADMET' stand for in pharmacology?",
        ["Activity, Dosage, Metabolism, Efficiency, Toxicology",
         "Absorption, Distribution, Metabolism, Excretion, Toxicity",
         "Analysis, Detection, Mapping, Evaluation, Testing",
         "Administration, Dosage, Mechanism, Effect, Timing"],
        n,
        notes="ADMET is a central concept in drug optimization. Knowing what each letter stands for helps students understand why we care about properties beyond just binding affinity.")

    n += 1
    make_quiz_answer_slide(prs, 3,
        "What does 'ADMET' stand for in pharmacology?",
        ["Activity, Dosage, Metabolism, Efficiency, Toxicology",
         "Absorption, Distribution, Metabolism, Excretion, Toxicity",
         "Analysis, Detection, Mapping, Evaluation, Testing",
         "Administration, Dosage, Mechanism, Effect, Timing"],
        1,
        "✓ ADMET = Absorption, Distribution, Metabolism, Excretion, Toxicity.\nAfter a drug is discovered to bind its target, it must also: get into the body (A), reach the right tissue (D), not be broken down too fast (M), leave the body safely (E), and not be toxic (T). AI predicts all 5!",
        n,
        notes="EXPLAIN IN DETAIL (each letter matters):\n\nA — ABSORPTION: Can the drug get from where you take it (stomach, skin, lung) into the bloodstream? For oral drugs, this means surviving stomach acid, crossing the intestinal wall, and surviving first-pass metabolism in the liver. Lipinski's Ro5 predicts this from molecular structure.\n\nD — DISTRIBUTION: Once in the bloodstream, does the drug reach the RIGHT tissue? For CNS drugs, the blood-brain barrier (BBB) is the key challenge. For cancer drugs, the tumor microenvironment. For heart drugs, cardiac tissue. Predicted by BBB permeability models, LogP, plasma protein binding models.\n\nM — METABOLISM: How quickly is the drug broken down, and into what products? The liver is the main metabolic organ (cytochrome P450 enzymes metabolize most drugs). Fast metabolism = drug is cleared quickly (short half-life). Some metabolites can be MORE toxic than the original drug (e.g., acetaminophen overdose: its metabolite NAPQI is the toxic species). AI predicts CYP450 interactions.\n\nE — EXCRETION: How does the drug leave the body? Mainly kidneys (urine) or bile (feces). Important for dosing schedules. Poor excretion → drug accumulates → toxicity.\n\nT — TOXICITY: Does the drug harm any tissue/organ? The most feared: hERG block (cardiac arrhythmia), hepatotoxicity (liver damage), genotoxicity (DNA damage → cancer risk). Predicting toxicity is one of the main AI applications in drug discovery — if you can eliminate toxic compounds in silico (computationally), you save enormously expensive animal and clinical trials.\n\nWHY AI IS ESSENTIAL: A drug that binds its target but fails ADMET is useless. In drug discovery, roughly 40% of failures are due to ADMET issues. Traditional approaches test ADMET in animals (expensive, slow). AI models trained on large databases can predict ADMET properties from structure alone — potentially before synthesis.")

    # Q4: IC50
    n += 1
    make_quiz_question_slide(prs, 4,
        "What does IC50 measure in pharmacology?",
        ["The minimum concentration at which a drug shows any effect",
         "The concentration that inhibits 50% of a target's activity",
         "The lethal dose for 50% of a test population",
         "The time at which a drug reaches 50% of its maximum concentration"],
        n,
        notes="IC50 is the most common bioactivity measure in medicinal chemistry. Students need to understand this to interpret ChEMBL data and QSAR predictions.")

    n += 1
    make_quiz_answer_slide(prs, 4,
        "What does IC50 measure in pharmacology?",
        ["The minimum concentration at which a drug shows any effect",
         "The concentration that inhibits 50% of a target's activity",
         "The lethal dose for 50% of a test population",
         "The time at which a drug reaches 50% of its maximum concentration"],
        1,
        "✓ IC50 = the concentration that inhibits 50% of a target's activity.\nLower IC50 = more potent (you need less drug). pIC50 = -log10(IC50) — used so higher values = more potent. ChEMBL database stores IC50 values for ~2.5 million compounds.",
        n,
        notes="EXPLAIN IN DETAIL:\n- IC50 = Inhibitory Concentration 50%. The concentration of drug needed to reduce a target's activity by 50%. This is measured in an assay: add increasing concentrations of drug, measure how much the target is inhibited at each concentration, fit a sigmoidal curve, read off the 50% point.\n- Units: usually nM (nanomolar), μM (micromolar), or mM (millimolar). 1 nM = one billionth of a mole per liter.\n- A drug with IC50 = 1 nM is MORE potent than a drug with IC50 = 1000 nM (1 μM), because you need 1000× less of it to achieve the same effect.\n- WHY 50%? It's the midpoint of the dose-response curve, where the measurement is most precise (the slope is steepest). It's a convention, not a magic threshold.\n- pIC50 = -log10(IC50). If IC50 = 1 nM = 10^-9 M, then pIC50 = -log10(10^-9) = 9. Higher pIC50 = better drug. QSAR models often predict pIC50 (or logIC50) because it's a better scale for linear regression.\n- LD50: Lethal Dose 50% — the dose that kills 50% of animals in a toxicity test. Completely different concept from IC50.\n- EC50: Effective Concentration 50% — for agonists (drugs that ACTIVATE a target). The concentration that produces 50% of maximal activation. If you're talking about an agonist (like a drug that activates a receptor), you use EC50. If it's an inhibitor (like an enzyme blocker or channel blocker), you use IC50.\n- Ki: binding affinity constant. Measured differently (competition assay). Lower Ki = tighter binding.\n- ChEMBL: a database maintained by EMBL-EBI containing ~2.5 million compounds with experimental IC50/Ki/EC50 measurements. This is the main training data source for QSAR models.")

    # Q5: BBB
    n += 1
    make_quiz_question_slide(prs, 5,
        "What is the 'BBB' in drug discovery?",
        ["Biologically-Bioavailable Benchmark — a measurement standard",
         "Blood-Brain Barrier — a protective membrane around the brain",
         "Basic Biochemical Binding — a type of protein interaction",
         "Broad Bioactivity Base — a large compound library"],
        n,
        notes="BBB is crucial for CNS drug development. This question teaches why we need special molecular properties for drugs that treat brain diseases.")

    n += 1
    make_quiz_answer_slide(prs, 5,
        "What is the 'BBB' in drug discovery?",
        ["Biologically-Bioavailable Benchmark — a measurement standard",
         "Blood-Brain Barrier — a protective membrane around the brain",
         "Basic Biochemical Binding — a type of protein interaction",
         "Broad Bioactivity Base — a large compound library"],
        1,
        "✓ BBB = Blood-Brain Barrier. A highly selective barrier between blood and brain tissue.\nEndothelial cells with tight junctions block most molecules. Only small, lipophilic molecules cross easily. This is why CNS drugs require special molecular properties: MW < 450, logP 1-3, HBD < 3.",
        n,
        notes="EXPLAIN IN DETAIL:\n- The brain is one of the most protected organs in the body. It sits inside the skull (physical protection) AND has a specialized molecular barrier.\n- Blood-brain barrier structure: brain capillaries (tiny blood vessels) are lined by endothelial cells that are joined by 'tight junctions' — protein complexes that seal the gaps between cells. In other organs, small molecules can slip between endothelial cells. In the brain, they can't.\n- How do molecules cross the BBB? Only a few ways:\n  1. Simple diffusion: if the molecule is small and lipophilic (fat-soluble), it dissolves in the membrane and passes through. This is the main route for CNS drugs.\n  2. Transporter proteins: specific transporters actively carry certain molecules across (glucose, amino acids, some drugs). Also works in REVERSE — efflux transporters (P-glycoprotein) PUMP drugs BACK OUT of the brain!\n  3. Receptor-mediated transcytosis: used by proteins like insulin.\n- The BBB evolved to protect the brain from toxins, pathogens, and fluctuations in blood composition. But for drug discovery, it's an obstacle.\n- WHY THIS MATTERS: L-DOPA is used for Parkinson's instead of dopamine BECAUSE dopamine doesn't cross the BBB well, but L-DOPA does (it uses an amino acid transporter). Once inside the brain, L-DOPA is converted to dopamine by DOPA decarboxylase.\n- MOLECULAR RULES FOR BBB PENETRATION (for QSAR models):\n  MW < 450 Da (smaller → more easily diffuses)\n  logP 1-3 (some lipophilicity needed, but too lipophilic = bad)\n  HBD (hydrogen bond donors) < 3 (hydrogen bonds make molecules sticky, slow diffusion)\n  TPSA (total polar surface area) < 90 Å²\n  These are the Lipinski-like rules for CNS drugs (Pardridge, 2005).\n- CNS DRUG CHALLENGE: ~97% of CNS drug candidates fail in clinical trials. The BBB is a major reason. AI that accurately predicts BBB penetration would be enormously valuable.")

    # Q6: ECFP / Morgan fingerprints
    n += 1
    make_quiz_question_slide(prs, 6,
        "What does 'ECFP4' stand for in molecular ML?",
        ["Exact Chemical Fingerprint Protocol, version 4",
         "Electrochemical Feature Profile with 4 iterations",
         "Extended Connectivity Fingerprint with radius 4",
         "Encoded Chemical Feature Pattern, 4-bit"],
        n,
        notes="This directly relates to today's lecture content. Most students will have seen ECFP4 mentioned on Day 1 but may not know what the number means.")

    n += 1
    make_quiz_answer_slide(prs, 6,
        "What does 'ECFP4' stand for in molecular ML?",
        ["Exact Chemical Fingerprint Protocol, version 4",
         "Electrochemical Feature Profile with 4 iterations",
         "Extended Connectivity Fingerprint with radius 4",
         "Encoded Chemical Feature Pattern, 4-bit"],
        2,
        "✓ ECFP4 = Extended Connectivity FingerPrint with radius 4 (i.e., 2 iterations).\nThe number refers to the diameter (= 2× radius). ECFP4 captures up to 2 atoms away from each center. It creates a 2048-bit binary vector. More widely known today as 'Morgan fingerprints'.",
        n,
        notes="CITATION: Rogers & Hahn (2010) 'Extended-Connectivity Fingerprints.' J. Chem. Inf. Model. 50(5):742-754. This paper formalized Morgan fingerprints as ECFP. The original Morgan algorithm dates to 1965: Morgan, H.L. (1965) J. Chem. Doc. 5:107-113.\n\nEXPLAIN IN DETAIL:\n- A fingerprint is a fixed-length vector of 0s and 1s (bits) that describes a molecule. Each bit says 'yes' or 'no' to whether the molecule contains a specific chemical substructure.\n- Think of it like a boolean feature vector: does the molecule have a benzene ring? (1/0). Does it have a carboxylic acid group? (1/0). Does it have a nitrogen connected to two carbons with an adjacent double bond? (1/0). With 2048 bits, you can encode 2048 such questions.\n- MORGAN/ECFP ALGORITHM:\n  1. Start at each heavy atom in the molecule\n  2. Collect information about the atom itself (atomic number, charge, etc.)\n  3. Look at all atoms 1 bond away ('radius 1 neighborhood') — collect their info, hash it together\n  4. Look at all atoms 2 bonds away ('radius 2 neighborhood') — hash together\n  5. All the hashed identifiers from all atoms at all radii → map to bit positions in a 2048-bit vector\n- ECFP4: each number (4) is the DIAMETER = 2× radius. ECFP4 uses radius 2, meaning it looks 2 bonds away from each center atom. ECFP6 uses radius 3.\n- WHY BINARY? Machine learning algorithms (especially kernel SVMs and similarity calculations) work efficiently with binary vectors. It also compresses information — Tanimoto similarity between two fingerprints can be computed as |A AND B| / |A OR B|.\n- PRACTICAL: In RDKit: from rdkit.Chem import AllChem; fp = AllChem.GetMorganFingerprintAsBitVect(mol, radius=2, nBits=2048)\n- This single line is one of the most important operations in practical molecular ML. We will use it in Practical 1 today.")

    # Q7: HTS
    n += 1
    make_quiz_question_slide(prs, 7,
        "What is 'HTS' in drug discovery?",
        ["Human Testing Stage — Phase I clinical trials",
         "High-Throughput Screening — testing millions of compounds rapidly",
         "Hybrid Target Selection — choosing drug targets using AI",
         "Heterogeneous Training Set — a diverse ML dataset"],
        n,
        notes="HTS vs. virtual screening is a key contrast for understanding why AI is valuable. Many students will not know what HTS is.")

    n += 1
    make_quiz_answer_slide(prs, 7,
        "What is 'HTS' in drug discovery?",
        ["Human Testing Stage — Phase I clinical trials",
         "High-Throughput Screening — testing millions of compounds rapidly",
         "Hybrid Target Selection — choosing drug targets using AI",
         "Heterogeneous Training Set — a diverse ML dataset"],
        1,
        "✓ HTS = High-Throughput Screening.\nRobotic systems test 100,000–1 million compounds/day against a biological target. Costs ~$1M per campaign. AI virtual screening can computationally screen billions of compounds at a fraction of the cost, guiding which compounds to actually synthesize and test.",
        n,
        notes="EXPLAIN IN DETAIL:\n- Traditional HTS: large pharmaceutical companies maintain 'compound libraries' of 1-2 million drug-like molecules, synthesized and stored over decades. To find a drug, they run an automated robot that dispensed tiny amounts (~1-10 nanoliters) of each compound into wells of a 384-well (or 1536-well) plate, adds their target protein (e.g., an enzyme), and measures a signal (fluorescence, luminescence, absorbance) that indicates whether the compound inhibited the target.\n- THROUGHPUT: modern HTS can screen 100,000-500,000 compounds per day. A full library of 1 million compounds takes ~2 weeks to screen.\n- COST: $1-3 per compound tested, so a 1M compound screen costs ~$1-3 million. Plus the cost of library maintenance, assay development, and follow-up validation.\n- VIRTUAL SCREENING: instead of physically testing each compound, we use a computer model (QSAR, docking, or GNN) to predict which compounds are most likely to be active. Then we only TEST the top-predicted compounds.\n- ADVANTAGE: chemical space is estimated at 10^60 drug-like molecules (much larger than 1 million). Virtual screening can explore this space computationally. HTS can only screen what you already have in a bottle.\n- DISADVANTAGE: virtual screening makes predictions based on a model, which has uncertainty. Some good compounds will be missed (false negatives), and some predicted actives will fail in the real assay (false positives). BUT: even if virtual screening misses half the actives, if it enriches the hit rate from 0.1% to 10%, you've reduced the required number of real assays by 100×.\n- THE AI WORKFLOW: virtual screening identifies top candidates (100-10,000 compounds) → these are physically tested in HTS → hits are confirmed and optimized → ADMET testing → clinical trials.")

    # Q8: SHAP
    n += 1
    make_quiz_question_slide(prs, 8,
        "What does 'SHAP' stand for in machine learning?",
        ["Statistical Hypothesis and Prediction",
         "Structural Heatmap and Pattern recognition",
         "SHapley Additive exPlanations",
         "Sequential Hierarchical Analysis Protocol"],
        n,
        notes="SHAP is covered in detail later today. This question introduces the term early so students are primed for it.")

    n += 1
    make_quiz_answer_slide(prs, 8,
        "What does 'SHAP' stand for in machine learning?",
        ["Statistical Hypothesis and Prediction",
         "Structural Heatmap and Pattern recognition",
         "SHapley Additive exPlanations",
         "Sequential Hierarchical Analysis Protocol"],
        2,
        "✓ SHAP = SHapley Additive exPlanations (named after Shapley values from game theory).\nFor each prediction, SHAP assigns a contribution score to each input feature. Allows us to explain WHY the model predicted a specific value. Critical for trust, debugging, and scientific discovery.",
        n,
        notes="CITATION: Lundberg & Lee (2017) 'A Unified Approach to Interpreting Model Predictions.' NeurIPS 30:4765-4774.\n\nEXPLAIN IN DETAIL:\n- Lloyd Shapley was a mathematician and economist who won the 2012 Nobel Prize in Economics. In 1953, he introduced 'Shapley values' as a fair way to distribute payoffs in a cooperative game: if several players cooperate to earn a reward, how much should each player get?\n- The key idea: a player's contribution = average of their marginal contributions across all possible orderings of players joining the coalition.\n- SHAP APPLIES THIS TO ML: instead of 'players', we have 'features'. Instead of 'payout', we have 'prediction'. For a specific molecule with a specific predicted pIC50 = 8.3, SHAP tells you: feature 'aromatic ring count' contributed +0.5, feature 'chlorine at position X' contributed +1.2, feature 'high molecular weight' contributed -0.8, etc. All contributions sum to: prediction - baseline prediction.\n- WHY THIS MATTERS:\n  1. Regulatory acceptance: FDA wants to understand model decisions. SHAP provides audit trail.\n  2. Chemist trust: medicinal chemists can say 'I see — the model predicts this is active because of the methyl group at position 3. I understand that from my experience.' Or they can say 'That SHAP attribution makes no sense chemically — I don't trust this prediction.'\n  3. New knowledge: if SHAP says the chlorine at position 4 is always the most important feature for hERG block, you now know a design rule: avoid chlorine at position 4.\n  4. Debugging: if SHAP says the model is using a structural feature that shouldn't matter (e.g., the SMILES string length), you've found a data leak or artifact.\n- PRACTICAL USE: we use shap.TreeExplainer for Random Forest and XGBoost (computationally exact, fast). For deep learning: KernelSHAP or DeepSHAP (approximate).")

    # Q9: What is a receptor?
    n += 1
    make_quiz_question_slide(prs, 9,
        "What is a 'receptor' in pharmacology?",
        ["A cell that detects external signals (like a neuron or sensory cell)",
         "A protein on a cell surface or inside cells that binds specific molecules and triggers a response",
         "Any molecule that receives electrons in a chemical reaction",
         "A medical device that monitors patient vital signs"],
        n,
        notes="This is a fundamental biology question. CS students may confuse 'receptor' with everyday usage. Understanding what receptors are is essential for making sense of drug mechanisms.")

    n += 1
    make_quiz_answer_slide(prs, 9,
        "What is a 'receptor' in pharmacology?",
        ["A cell that detects external signals (like a neuron or sensory cell)",
         "A protein on a cell surface or inside cells that binds specific molecules and triggers a response",
         "Any molecule that receives electrons in a chemical reaction",
         "A medical device that monitors patient vital signs"],
        1,
        "✓ A receptor is a protein that binds a specific molecule (ligand) and changes its activity as a result.\nMost drug receptors are on the cell surface or inside cells. About 34% of all FDA drugs target GPCRs (G Protein-Coupled Receptors) — a specific family of receptors with 7 membrane-spanning segments.",
        n,
        notes="EXPLAIN IN DETAIL:\n- In pharmacology, 'receptor' has a VERY specific meaning: a macromolecule (almost always a protein) that binds a ligand (drug or endogenous molecule like a neurotransmitter or hormone) with high specificity and affinity, and as a result changes its activity (is 'activated' or 'blocked').\n- DO NOT CONFUSE WITH:\n  1. Sensory receptor = a cell that detects stimuli (photoreceptor in the eye detects light). These CONTAIN receptor proteins, but the cell is not itself what pharmacologists call 'a receptor'.\n  2. Electron acceptor in chemistry = completely different, electron transfer context.\n  3. Medical monitoring device = no relation.\n\n- CLASSES OF DRUG RECEPTORS:\n  1. GPCRs (G Protein-Coupled Receptors): ~800 in the human genome. 7 membrane-spanning domains. When activated by agonist → change shape → activates G protein inside cell → second messenger cascade (cAMP, IP3, etc.) → cellular response. Examples: dopamine D2 receptor, beta-2 adrenergic receptor (asthma inhalers), opioid receptors. ~34% of all drugs target GPCRs.\n  2. Ligand-gated ion channels: protein pore that opens when a specific molecule binds. Examples: GABA-A (opens Cl- pore), nAChR (opens Na+/K+ pore), NMDA receptor (glutamate, opens Ca2+ pore). Also called 'ionotropic receptors'. ~15% of drugs target ion channels overall.\n  3. Receptor tyrosine kinases (RTKs): cell surface proteins that become enzymes when activated. Important in cancer (e.g., EGFR, HER2). Targeted by many kinase inhibitor drugs (imatinib, erlotinib).\n  4. Nuclear receptors: receptors inside the cell that, when bound by lipophilic hormones (steroids, thyroid hormones), move to the nucleus and change gene expression. Examples: estrogen receptor (tamoxifen for breast cancer), glucocorticoid receptor (prednisone).\n\n- THE LOCK-AND-KEY MODEL: drugs bind receptors like a key fits a lock. The drug shape must match the receptor binding pocket shape. This is exactly what molecular docking (Day 3) and QSAR predict: given a molecular structure, does it fit the receptor?\n\n- AGONIST vs. ANTAGONIST: an agonist ACTIVATES the receptor (turns it ON). An antagonist BLOCKS the receptor (prevents activation, turns it OFF). A partial agonist activates it partially. An inverse agonist suppresses below baseline activity.")

    # Q10: LogP
    n += 1
    make_quiz_question_slide(prs, 10,
        "What does 'LogP' measure in medicinal chemistry?",
        ["The logarithm of a drug's potency (pIC50)",
         "The log of the ratio of drug concentration in octanol vs. water — measures lipophilicity",
         "The number of chiral centers in a molecule expressed on a log scale",
         "The logarithm of a drug's molecular weight"],
        n,
        notes="LogP comes up constantly in drug design (Lipinski's Rule of Five, BBB penetration, ADMET). Students need to understand what it measures conceptually.")

    n += 1
    make_quiz_answer_slide(prs, 10,
        "What does 'LogP' measure in medicinal chemistry?",
        ["The logarithm of a drug's potency (pIC50)",
         "The log of the ratio of drug concentration in octanol vs. water — measures lipophilicity",
         "The number of chiral centers in a molecule expressed on a log scale",
         "The logarithm of a drug's molecular weight"],
        1,
        "✓ LogP = log10([drug in octanol]/[drug in water]). Measures lipophilicity (fat-loving vs. water-loving).\nHigher LogP = more lipophilic (oily). Lower = more hydrophilic (watery). Lipinski's Rule of Five requires LogP ≤ 5. CNS drugs need LogP 1-3 to cross the blood-brain barrier.",
        n,
        notes="EXPLAIN IN DETAIL:\n- P = partition coefficient. If you shake a drug with octanol (a fatty oil-like liquid) and water, the drug distributes between the two layers. P = [drug in octanol layer] / [drug in water layer].\n- Octanol is used as a MODEL for biological membranes (cell membrane lipid bilayer). Octanol is a fatty 8-carbon alcohol that mimics the lipid part of biological membranes reasonably well.\n- LogP: since P values range from 0.0001 to 100,000+, we use the log10 scale. LogP = 0 means equal distribution (P=1, 50/50 in both). LogP = 3 means 1000× more in octanol (lipophilic). LogP = -2 means 100× more in water (hydrophilic).\n- WHY DOES THIS MATTER FOR DRUG DISCOVERY?\n  1. Oral absorption: drugs must cross the intestinal cell membrane (lipid bilayer). Some lipophilicity is needed. But too lipophilic → poor water solubility → can't be absorbed from the gut → insoluble tablets!\n  2. BBB penetration: brain capillary membranes are tight. Lipophilic molecules diffuse through more easily. BBB target: LogP 1-3.\n  3. Metabolism: more lipophilic drugs tend to be metabolized faster by the liver (the liver tries to make lipophilic drugs more water-soluble so the kidneys can excrete them). Drug half-life correlates with LogP.\n  4. Toxicity: very lipophilic drugs can accumulate in fatty tissues (liver, brain fat), leading to chronic toxicity.\n- LIPINSKI'S RULE: LogP ≤ 5. This excludes very lipophilic molecules that would be poorly absorbed or highly toxic.\n- CALCULATED vs. EXPERIMENTAL: experimental LogP (shake-flask method) is expensive. Computational LogP (cLogP) can be predicted instantly from molecular structure. RDKit: Chem.Descriptors.MolLogP(mol). Accuracy of computed LogP: ~0.5-1 log unit.\n- RELATED DESCRIPTORS: LogD = LogP at a specific pH (important for drugs that can be ionized). TPSA = Total Polar Surface Area (related to, but different from, LogP — measures the area of polar atoms).\n- FOR TODAY'S PRACTICAL: LogP is one of the key features we include alongside Morgan fingerprints in our QSAR model.")

    # Closing
    n += 1
    make_section_divider(prs, "Great! Now You Speak the Language 📚", n,
        notes="End of Quiz 2. Transition to the main Day 2 lecture on molecular ML. Students should now understand the key abbreviations they'll encounter all day. Reference back to these definitions as needed throughout the day.",
        subtitle="QSAR, SMILES, ECFP, ADMET, IC50, BBB, HTS, SHAP, receptor, LogP — check!")

    os.makedirs("quiz_abbreviations", exist_ok=True)
    prs.save("quiz_abbreviations/slides.pptx")
    print(f"  Quiz 2 (Abbreviations): {n} slides → quiz_abbreviations/slides.pptx")




# =====================================================================
# DAY 2 ANNOTATED — Same content, FULL plain-language presenter notes
# For instructors who want to understand EVERYTHING on each slide
# =====================================================================
def generate_day2_annotated():
    prs = new_prs()
    n = 0

    # --- Title ---
    n += 1
    make_title_slide(prs,
        "Molecular ML: From Fingerprints to Predictions",
        "Represent molecules, train models, evaluate properly, and explain results",
        2, "AI for Drug Discovery — ANNOTATED VERSION", date_str="22 April 2026",
        notes="""ANNOTATED VERSION — COMPREHENSIVE PRESENTER NOTES

=== WHAT THIS DECK IS ===
This is the SAME content as Day 2, but with exhaustive notes explaining every concept in plain language. If you are not familiar with a term on any slide, look at these notes for a complete explanation. You do not need to have prior knowledge of chemistry, pharmacology, or machine learning to deliver this lecture — all the key concepts are explained here.

=== DAY 2 OVERVIEW ===
Day 2 is the most technically dense day of the course. By the end, students will have:
1. Represented molecules as numbers (SMILES → fingerprints → descriptors)
2. Trained a QSAR model (a machine learning model that predicts drug activity from molecular structure)
3. Evaluated the model correctly (NOT with random splits, but with scaffold splits)
4. Explained the model's predictions using SHAP values

The day is split into morning (Units 1-4, ~180 min) and afternoon (Units 5-8, ~180 min) with two 15-minute breaks.

=== HOW TO USE THESE NOTES ===
Before each slide, read the notes. They explain:
- What every acronym means
- Why the concept matters
- What to say to the students
- Common questions and misconceptions
- The biology/chemistry context

After delivering the lecture once, you will be able to improvise. Until then, these notes are your guide.""")

    # --- Morning Agenda ---
    n += 1
    make_content_slide(prs,
        "Morning Agenda (Units 1-4)",
        ["1. SMILES deep dive — syntax rules, practice, limitations",
         "2. Molecular fingerprints — Morgan/ECFP, Tanimoto similarity",
         "3. Feature engineering — descriptors, Lipinski, BBB rules",
         "4. QSAR: Quantitative Structure-Activity Relationships",
         "5. ML baselines: Random Forest & XGBoost",
         "6. PRACTICAL 1: Build a QSAR model for solubility prediction",
         "",
         "☕ BREAK after Unit 2 (after ~90 min)",
         "",
         "CNS drug focus: blood-brain barrier penetration, neuroscience targets"],
        n,
        notes="""=== MORNING AGENDA — EXPLAINING EACH ITEM ===

1. SMILES DEEP DIVE
SMILES = Simplified Molecular Input Line Entry System. This is the text format used to represent molecules as strings. Example: ethanol = 'CCO'. Aspirin = 'CC(=O)Oc1ccccc1C(=O)O'. The deep dive covers the grammar rules (how to read and write SMILES). This matters because ALL molecular ML starts with SMILES as the input.

2. MOLECULAR FINGERPRINTS
A fingerprint converts a molecule into a fixed-length binary vector (a sequence of 0s and 1s). Each bit represents whether a specific chemical substructure is present or absent. Morgan/ECFP fingerprints are the most widely used — they are computed by looking at each atom and its local neighborhood in the molecular graph. Tanimoto similarity is a measure of how similar two fingerprints are (like Jaccard similarity for sets: 0 = completely different, 1 = identical).

3. FEATURE ENGINEERING
In addition to fingerprints, we compute 'descriptors' — quantitative properties of a molecule like: molecular weight (MW), how fat-soluble it is (LogP), how many hydrogen bond donors/acceptors it has (HBD/HBA), how much surface area is polar (TPSA). These are used in Lipinski's Rule of Five (a set of rules that predict whether a drug can be taken as a pill) and BBB rules (which predict if a drug can cross the blood-brain barrier to reach the brain).

4. QSAR — QUANTITATIVE STRUCTURE-ACTIVITY RELATIONSHIPS
This is the main topic. QSAR = using ML to predict a biological activity (like 'how strongly does this molecule inhibit this protein') from molecular structure. Workflow: get molecular features (fingerprints + descriptors) → train RF/XGBoost → evaluate → predict for new molecules.

5. ML BASELINES: RANDOM FOREST & XGBOOST
Two powerful, non-deep-learning ML algorithms that work very well for QSAR:
- Random Forest: build many independent decision trees, average their predictions
- XGBoost: build trees sequentially, each correcting the previous tree's errors
Both are excellent starting points before considering more complex models.

6. PRACTICAL 1
Students build their own QSAR model using Python, RDKit, and scikit-learn. The dataset is Delaney solubility (1,128 molecules with measured aqueous solubility values). Goal: predict log(solubility) from molecular structure.

CNS DRUG FOCUS: blood-brain barrier (BBB) penetration and neuroscience drug targets come up repeatedly as examples throughout the morning. This is relevant because many students' projects will involve CNS drugs.""")

    # --- SMILES Section Divider ---
    n += 1
    make_section_divider(prs, "SMILES: Deep Dive", n,
        notes="""=== SMILES DEEP DIVE — WHAT THIS SECTION IS ===

SMILES (Simplified Molecular Input Line Entry System) was invented by David Weininger at Daylight Chemical Information Systems in 1988. The problem he solved: how do you put a chemical structure into a computer?

Before SMILES, there were complicated graph-based formats (MDL Molfile) that were hard to type and not human-readable. Weininger realized: if you do a depth-first traversal of the molecular graph and write down each atom and bond, you get a compact, human-readable string.

Why teach SMILES as the first topic?
Because EVERY molecular ML pipeline starts with SMILES. The molecule enters the model as a SMILES string, and from that string, all other representations are computed (fingerprints, graphs, 3D coordinates). If students don't understand SMILES, they can't understand the rest.

TIMING: You should get through this section (slides 4-5) in about 15 minutes total. Use the whiteboard or a drawing on screen to supplement the slides.""",
        subtitle="Weininger (1988), J. Chem. Inf. Comput. Sci. 28:31-36")

    # --- SMILES Syntax ---
    n += 1
    make_content_slide(prs,
        "SMILES Syntax Rules",
        ["Atoms: C, N, O, S, P, F, Cl, Br, I (organic subset, implicit H)",
         "Single bond: implicit (CC = ethane)  |  Double: = (C=O)  |  Triple: # (C#N)",
         "Aromatic: lowercase (c1ccccc1 = benzene)",
         "Branches: parentheses — CC(=O)O = acetic acid",
         "Rings: matching digits — C1CCCCC1 = cyclohexane",
         "Charges: [NH4+], [O-]  |  Stereochemistry: / \\ for E/Z, @ @@ for R/S",
         "",
         "Canonical SMILES: unique representation (RDKit Chem.MolToSmiles())",
         "",
         "Alternatives:",
         "  SELFIES — 100% valid strings (Krenn et al. 2020) — great for generative AI",
         "  InChI — IUPAC standard  |  DeepSMILES — neural network friendly"],
        n,
        notes="""=== SMILES SYNTAX — FULL EXPLANATION ===

ATOMS:
Each atom is written as its elemental symbol. For the 'organic subset' (C, N, O, S, P, F, Cl, Br, I), hydrogens are IMPLICIT — you don't write them. The number of hydrogens is inferred by the standard valence rules:
- Carbon (C): valence 4. If written as 'C' with 1 bond, it has 3 implicit H. With 4 bonds, it has 0 implicit H.
- Nitrogen (N): valence 3 normally. 'N' with no bonds = NH3 (three implicit H).
- Oxygen (O): valence 2. 'O' with no bonds = H2O (two implicit H).
Atoms outside the organic subset (like metal ions) must be written in brackets with explicit hydrogens: [Fe++], [NH4+].

BONDS:
- Single bond: write nothing between atoms (CC = ethane: two carbons each connected to the other and implicit H)
- Double bond: use '=' sign (C=C is ethylene, C=O is formaldehyde/carbonyl)
- Triple bond: use '#' (C#N is hydrogen cyanide, C#C is acetylene)
- Aromatic bond: use ':' or lowercase letters (c:c or cc in aromatic ring)

BRANCHING:
When the chain branches, put the branch in parentheses.
CC(=O)O = acetic acid (vinegar):
- First C = methyl
- Second C connects to: =O (double bond oxygen, the carbonyl) AND O (hydroxyl, from the parenthesis and the continuing chain)
- Full structure: CH3-C(=O)-OH = CH3COOH = acetic acid

RINGS:
To represent a ring, you start at one atom, give it a number (the 'ring closure digit'), traverse the ring, and when you return to close it, write the same number again.
C1CCCCC1 = cyclohexane: start at carbon (labeled '1'), traverse 5 more carbons, close back to '1'. The ring is formed.
c1ccccc1 = benzene: same but lowercase (aromatic). The Kekulé structure would alternate single and double bonds, but SMILES shorthand uses lowercase.

CANONICAL SMILES:
Multiple SMILES strings can represent the same molecule (you can traverse the graph starting from any atom). CCO and OCC and C(C)O all represent ethanol. Canonical SMILES is the unique, standardized form — always the same for a given molecule. RDKit: Chem.MolToSmiles(mol).

ALTERNATIVES:
- SELFIES (Self-Referencing Embedded Strings, Krenn et al. 2020): a newer encoding where every possible SELFIES string decodes to a valid molecule (SMILES can produce invalid strings). Better for generative ML models.
- InChI (International Chemical Identifier): IUPAC standard, hierarchical. More information than SMILES but not as human-readable.
- DeepSMILES: modified SMILES designed for character-based neural networks (avoids some issues with branch parentheses and ring closure digits).

DRAW ON THE BOARD: Draw benzene (a hexagon) and show how c1ccccc1 traverses it. Draw aspirin and show how the SMILES represents each part. This visual always helps.""")

    # --- SMILES Practice ---
    n += 1
    make_content_slide(prs,
        "SMILES Practice: Can You Read These?",
        ["1. NCCc1c[nH]c2ccc(O)cc12  →  ?  (hint: neurotransmitter, mood)",
         "2. NCCc1ccc(O)c(O)c1  →  ?  (hint: reward pathway)",
         "3. NCCCC(=O)O  →  ?  (hint: main inhibitory NT in the brain)",
         "4. CC(=O)Oc1ccccc1C(=O)O  →  ?  (hint: common painkiller)",
         "5. Cn1c(=O)c2c(ncn2C)n(C)c1=O  →  ?  (hint: in your coffee)",
         "",
         "Answers: 1) Serotonin  2) Dopamine  3) GABA  4) Aspirin  5) Caffeine",
         "",
         "Note: GABA is simpler than serotonin, but both are critical NTs",
         "GABA-A receptor is in the same Cys-loop superfamily as GluCl (ivermectin target)"],
        n,
        notes="""=== SMILES PRACTICE — DETAILED EXPLANATIONS ===

This is an interactive slide. Give students 2 minutes to try to decode the SMILES, then reveal one by one.

1. SEROTONIN: NCCc1c[nH]c2ccc(O)cc12
Let's parse this:
- NCC: an ethylamine chain (NH2 - CH2 - CH2)
- c1c[nH]c2ccc(O)cc12: an indole ring system. The '1' and '2' are ring closure labels. The [nH] is an aromatic nitrogen with one hydrogen.
- Full structure: 5-hydroxytryptamine (5-HT). The 'N' side chain is the amine, the 'OH' at the ccc(O) part is the 5-hydroxy group.
- PHARMACOLOGY: Serotonin is a neurotransmitter synthesized from tryptophan. Its receptors include: 5-HT1A (targeted by buspirone for anxiety), 5-HT2A (targeted by atypical antipsychotics, LSD), 5-HT3 (ion channel, targeted by ondansetron for nausea), SERT (serotonin reuptake transporter, targeted by SSRIs like fluoxetine/Prozac).
- DOES NOT CROSS BBB: Serotonin is hydrophilic (LogP ~ 0.2) and does not cross the BBB well. SSRIs work by blocking reuptake of serotonin IN the brain (they cross the BBB because they're more lipophilic).

2. DOPAMINE: NCCc1ccc(O)c(O)c1
- NCC: ethylamine (same as serotonin's side chain!)
- c1ccc(O)c(O)c1: catechol ring (a benzene ring with two adjacent OH groups)
- Full structure: 3,4-dihydroxyphenethylamine. The catechol + ethylamine structure.
- PHARMACOLOGY: Dopamine = motivation, reward, motor control. Parkinson's disease = loss of dopaminergic neurons in the substantia nigra. Treatment: L-DOPA (which crosses BBB, is converted to dopamine by DOPA decarboxylase). Antipsychotics (haloperidol, clozapine) block D2 receptors.
- DOES NOT CROSS BBB: dopamine doesn't cross BBB (too hydrophilic, no BBB transporter). L-DOPA does (uses the large neutral amino acid transporter).

3. GABA: NCCCC(=O)O
- N: amine group (NH2)
- CCC: three carbons
- C(=O)O: carboxylic acid (COOH)
- Full structure: gamma-aminobutyric acid. The main INHIBITORY neurotransmitter.
- PHARMACOLOGY: GABA-A (Cl- channel, ionotropic) and GABA-B (GPCR, metabotropic). Diazepam (Valium) potentiates GABA-A. Alcohol activates GABA-A. Gabapentin (for epilepsy, pain) modulates GABA metabolism. GABA-A is in the Cys-loop receptor superfamily, same as GluCl.

4. ASPIRIN: CC(=O)Oc1ccccc1C(=O)O
- CC(=O)O: acetyl ester (the acetyl group connected via oxygen)
- c1ccccc1: benzene ring
- C(=O)O: carboxylic acid
- Full structure: acetylsalicylic acid
- PHARMACOLOGY: COX-1/COX-2 inhibitor. Blocks prostaglandin synthesis → reduces inflammation, pain, fever. Low dose: prevents platelet aggregation → prevents heart attacks/strokes.

5. CAFFEINE: Cn1c(=O)c2c(ncn2C)n(C)c1=O
- Complex purine ring system (xanthine scaffold with 3 methyl groups)
- Adenosine A1/A2A receptor ANTAGONIST. Adenosine normally makes you sleepy. Caffeine blocks adenosine → you feel alert.
- LogP = -0.07 (quite water-soluble), MW = 194 Da. Very efficiently crosses the BBB!

KEY TEACHING POINT: Ask students to notice how GABA (NCCCC(=O)O) is much simpler than serotonin or caffeine. But GABA is the main inhibitory neurotransmitter in the brain — biological importance doesn't correlate with molecular complexity.""")

    # --- Fingerprints Section ---
    n += 1
    make_section_divider(prs, "Molecular Fingerprints", n,
        notes="""=== MOLECULAR FINGERPRINTS — WHAT THIS SECTION IS ===

WHY DO WE NEED FINGERPRINTS?
SMILES is text. Machine learning algorithms need numbers (floating point vectors or binary vectors). Fingerprints convert the molecular STRUCTURE into a fixed-length vector of 0s and 1s. This is the representation used for virtually all classical ML methods (Random Forest, XGBoost, SVM) in drug discovery.

WHY NOT JUST USE THE SMILES STRING DIRECTLY?
You could train a recurrent neural network or Transformer on SMILES strings (some models do this). But for classical ML (Random Forest, XGBoost), you need a fixed-length input vector. Fingerprints provide this. Also, fingerprints are rotation/translation invariant — the same molecule always gets the same fingerprint regardless of how you draw it.

WHY NOT JUST USE DESCRIPTORS?
Descriptors (MW, LogP, etc.) lose structural information — you can't recover the molecule from its descriptors. Fingerprints preserve more structural information. In practice, combining fingerprints + descriptors is the best classical approach.""",
        subtitle="From molecular structure to numerical vectors")

    # --- Fingerprint Concept ---
    n += 1
    make_content_slide(prs,
        "Molecular Fingerprints: The Concept",
        ["A fingerprint converts a molecule into a fixed-length bit vector",
         "Each bit indicates presence/absence of a substructure",
         "",
         "Types:",
         "  Structural keys (MACCS, 166 bits) — predefined patterns",
         "  Topological (RDKit FP) — paths through molecular graph",
         "  Circular (Morgan/ECFP) — local atom environments at radius r",
         "",
         "Morgan fingerprints are the most widely used today",
         "  Rogers & Hahn (2010), JCIM 50:742-754",
         "  ECFP4 = radius 2, 2048 bits = default starting point",
         "",
         "In RDKit: AllChem.GetMorganFingerprintAsBitVect(mol, radius=2, nBits=2048)"],
        n,
        notes="""=== MOLECULAR FINGERPRINTS — FULL EXPLANATION ===

THE CONCEPT:
A fingerprint is a binary vector (a list of 0s and 1s). Imagine a 2048-element list. Each position in the list represents a specific chemical substructure. If the molecule CONTAINS that substructure, the bit is 1. If not, it's 0.

ANALOGY: Think of a 'chemical checklist'. Does the molecule have a benzene ring? → bit 47 = 1. Does it have a carboxylic acid? → bit 312 = 1. Does it have a fluorine on an aromatic ring? → bit 891 = 0 (if it doesn't). The full list of 0s and 1s is the 'fingerprint'.

TYPE 1 — STRUCTURAL KEYS (MACCS):
Uses 166 predefined substructure questions. Was the dominant fingerprint type in the 1990s. The questions are fixed (e.g., 'does it have a thiocarbonyl?'). Limitation: only 166 bits → limited resolution. A molecule can be complex but the fingerprint only captures the 166 predefined patterns.

TYPE 2 — TOPOLOGICAL (RDKit FP):
Enumerate all paths up to a certain length through the molecular graph. Hash each path to a bit position. More flexible than structural keys but still path-based.

TYPE 3 — CIRCULAR (Morgan/ECFP) — the standard today:
Instead of paths, look at the LOCAL ENVIRONMENT around each atom. For each atom, collect information about all atoms within radius r (r=2 for ECFP4). Hash the combined information to a bit position.

WHY 'CIRCULAR'? If you draw a circle of radius r around each atom and look at all atoms inside that circle, you get a 'circular' neighborhood. Radius 2 = 2 bonds away from the center atom.

WHY THIS WORKS WELL:
Biological activity depends on LOCAL chemical environments. A nitrogen next to two aromatic rings in a particular configuration creates a specific 3D shape and electronics that match a receptor pocket. ECFP captures this local pattern.

THE CODE (RDKit):
from rdkit import Chem
from rdkit.Chem import AllChem
mol = Chem.MolFromSmiles('CCO')  # Parse SMILES
fp = AllChem.GetMorganFingerprintAsBitVect(mol, radius=2, nBits=2048)
fp_array = list(fp)  # Convert to Python list of 0s and 1s
# This 2048-element list is the fingerprint

For scikit-learn, you need a numpy array:
import numpy as np
fp_array = np.array(fp)""")

    # --- Morgan Algorithm ---
    n += 1
    make_content_slide(prs,
        "Morgan / ECFP: How They Work",
        ["Algorithm (for each atom):",
         "  1. Start: identifier = atom type + properties",
         "  2. Iteration 1: collect identifiers from radius-1 neighbors",
         "  3. Hash combined info → new identifier",
         "  4. Iteration 2: expand to radius-2, hash again",
         "  5. Map all identifiers to fixed-length bit vector",
         "",
         "ECFP4 = radius 2  |  ECFP6 = radius 3",
         "",
         "Similarity: Tanimoto coefficient = |A∩B| / |A∪B|",
         "  Range: 0 (different) to 1 (identical)",
         "  Tanimoto > 0.85 → likely similar activity (but exceptions exist!)",
         "  In RDKit: DataStructs.TanimotoSimilarity(fp1, fp2)"],
        n,
        notes="""=== MORGAN ALGORITHM — STEP BY STEP ===

Let me walk through a concrete example. Consider a molecule with 5 atoms: A-B-C-D-E (a simple chain). We'll compute the ECFP4 (radius 2) fingerprint for atom C.

STEP 1 — INITIAL IDENTIFIERS:
Each atom gets an initial identifier based on its properties:
- Atomic number (is it C, N, O...?)
- Degree (how many bonds does it have?)
- Number of hydrogens
- Formal charge
- Is it in a ring?
- Is it aromatic?
These are hashed into a single integer (the 'initial identifier').
For atom C (the middle): it has bonds to B and D, so degree 2. Identifier = hash(C_properties).

STEP 2 — RADIUS 1 ITERATION:
For atom C: collect the identifiers of A (radius-2 from C), B (radius-1), D (radius-1), E (radius-2).
At radius 1: we see B and D.
New identifier for C at radius 1 = hash(C_r0, B_r0, D_r0) where r0 means the initial identifier.

STEP 3 — RADIUS 2 ITERATION:
At radius 2 from C: we also see A and E.
New identifier for C at radius 2 = hash(C_r1, A_r0, B_r1, D_r1, E_r0).

STEP 4 — COLLECT ALL IDENTIFIERS:
We do this for EVERY atom in the molecule (A, B, C, D, E), at each radius level (0, 1, 2 for ECFP4). We collect all the integers.

STEP 5 — BIT MAPPING:
Each integer is mapped to a bit position: integer mod 2048 (for 2048-bit fingerprint). Set that bit to 1.

COLLISIONS: Two different substructures might map to the same bit (hash collision). With 2048 bits, this is rare but happens. Count fingerprints (not bit fingerprints) avoid this but are not fixed-length.

WHY ECFP4 > ECFP6 FOR MANY TASKS?
ECFP4 captures local environments (2 bonds). ECFP6 captures larger environments (3 bonds). For activity prediction: activity is often determined by a specific functional group + local context (2-3 bonds). ECFP6 can sometimes over-specify (the fingerprint becomes too specific, two similar molecules look very different).

TANIMOTO COEFFICIENT:
Mathematical definition: |A ∩ B| / |A ∪ B| where A and B are the SET of 'on' bits in the two fingerprints.
Example: Fingerprint A has bits {1, 5, 7, 12} on. Fingerprint B has bits {1, 3, 7, 15} on.
A ∩ B (shared): {1, 7} = 2 bits
A ∪ B (combined): {1, 3, 5, 7, 12, 15} = 6 bits
Tanimoto = 2/6 = 0.33

SIMILAR PROPERTY PRINCIPLE: molecules with similar structures tend to have similar biological activity. The rule of thumb: Tanimoto > 0.85 → likely in the same chemical series, similar activity. But ACTIVITY CLIFFS exist — this will be discussed in the afternoon.

CONNECTION TO DAY 3: The Morgan algorithm is essentially the same as message passing in a Graph Neural Network (GNN). In GNNs, instead of hashing the neighborhood information, you LEARN how to combine it (trainable weights). Same concept, much more powerful.""")

    # --- Feature Engineering ---
    n += 1
    make_content_slide(prs,
        "Feature Engineering & Lipinski's Rule of Five",
        ["Physicochemical descriptors:",
         "  MW, LogP (lipophilicity), HBD, HBA, TPSA, rotatable bonds, aromatic rings",
         "",
         "Lipinski (Pfizer, 1997) — most cited paper in medicinal chemistry:",
         "  Orally bioavailable if: MW ≤ 500, LogP ≤ 5, HBD ≤ 5, HBA ≤ 10",
         "",
         "For CNS drugs — crossing the blood-brain barrier (Pardridge 2005):",
         "  MW < 450, LogP 1-3, HBD < 3, TPSA < 90 Å²",
         "",
         "Example: serotonin (MW=176, LogP=0.2) does NOT cross BBB well",
         "  → That's why we give SSRIs (which cross BBB) not serotonin itself",
         "  → L-DOPA crosses BBB, dopamine doesn't → Parkinson's treatment",
         "",
         "RDKit: 200+ descriptors  |  Mordred: >1800 descriptors"],
        n,
        notes="""=== FEATURE ENGINEERING — FULL EXPLANATION ===

WHAT ARE 'PHYSICOCHEMICAL DESCRIPTORS'?
These are numbers computed from molecular structure that describe PHYSICAL and CHEMICAL properties of the molecule. Unlike fingerprints (which are binary), descriptors are continuous numbers.

KEY DESCRIPTORS EXPLAINED:

MW (Molecular Weight):
The mass of the molecule in Daltons (Da) or g/mol. Ethanol = 46 Da. Aspirin = 180 Da. Herceptin (trastuzumab, a monoclonal antibody) = ~148,000 Da. Small molecule drugs are typically 200-500 Da.
For ML: larger molecules are harder to dose (you'd need a very large pill), harder to absorb, harder to get across cell membranes.

LogP (Octanol-Water Partition Coefficient):
If you shake a drug with a mixture of octanol (a fatty liquid) and water, some drug goes to the octanol phase and some to the water phase. P = [octanol] / [water]. LogP = log10(P).
LogP = 0: equally distributed (50/50).
LogP = 3: 1000× more in octanol → lipophilic (fatty/oily).
LogP = -2: 100× more in water → hydrophilic (water-loving).
Matters because: oral absorption requires crossing the lipid bilayer of intestinal cells → needs some lipophilicity. BBB penetration also requires lipophilicity. But too lipophilic → poor aqueous solubility → can't dissolve in stomach contents.

HBD (Hydrogen Bond Donors):
Atoms with N-H or O-H bonds that can DONATE a hydrogen bond. Examples: -OH (hydroxyl), -NH2 (amine), -COOH (carboxylic acid). Count of these groups.
Too many HBDs → molecule is 'sticky', can't easily pass through lipid membranes.

HBA (Hydrogen Bond Acceptors):
Atoms (typically N and O) that can ACCEPT a hydrogen bond (they have lone electron pairs). Count of N and O atoms.

TPSA (Total Polar Surface Area):
The surface area of all polar atoms (N, O, and attached H). Measured in Å² (square angstroms).
TPSA > 140 Å² → likely can't cross cell membranes.
TPSA < 90 Å² → recommended for CNS drugs.

LIPINSKI'S RULE OF FIVE (Ro5):
In 1997, Christopher Lipinski at Pfizer analyzed all FDA-approved drugs and identified that those with good oral bioavailability had:
1. MW ≤ 500 (molecular weight)
2. LogP ≤ 5 (lipophilicity)
3. HBD ≤ 5 (hydrogen bond donors)
4. HBA ≤ 10 (hydrogen bond acceptors)
All four numbers are divisible by 5 (hence 'Rule of Five'). Called 'Rule of Five' but it's actually four rules. If 2+ are violated, oral absorption is poor.

EXCEPTIONS: biologics (antibodies, proteins) BREAK the Ro5 but are injected, not taken orally. Natural products (antibiotics like ivermectin, erythromycin) often break Ro5 but use active transporters. The Ro5 applies to passive diffusion.

RDKIT CODE:
from rdkit.Chem import Descriptors
mw = Descriptors.MolWt(mol)
logp = Descriptors.MolLogP(mol)
hbd = Descriptors.NumHDonors(mol)
hba = Descriptors.NumHAcceptors(mol)
tpsa = Descriptors.TPSA(mol)""")

    # --- QSAR ---
    n += 1
    make_content_slide(prs,
        "QSAR: Quantitative Structure-Activity Relationships",
        ["Core idea: molecular structure determines biological activity",
         "  Activity = f(molecular features)",
         "  Hansch (1964) — founded the field",
         "",
         "Modern QSAR workflow:",
         "  1. Curate dataset from ChEMBL (target + activity data)",
         "  2. Generate features (fingerprints + descriptors)",
         "  3. Split data (scaffold split, NOT random!)",
         "  4. Train model (RF, XGBoost, GNN)",
         "  5. Evaluate rigorously (RMSE, R², AUC-ROC)",
         "  6. Interpret (SHAP, feature importance)",
         "",
         "This is the core pipeline you'll use for your projects!"],
        n,
        notes="""=== QSAR — FULL EXPLANATION ===

WHAT IS QSAR?
QSAR = Quantitative Structure-Activity Relationship. The idea: if you know the molecular STRUCTURE of a drug (its atoms, bonds, 3D shape, electronic properties), you can PREDICT its biological ACTIVITY (how strongly it binds to a protein, how toxic it is, how soluble it is).

This isn't magic — it's based on physical chemistry: drug binding depends on shape complementarity and chemical interactions (hydrogen bonds, van der Waals, electrostatics) between the drug and its protein target. Both the drug shape/chemistry AND the protein shape/chemistry are determined by molecular structure. Therefore structure → binding → activity.

HISTORY:
Corwin Hansch (UC Riverside) formalized QSAR in 1964 using linear regression: biological activity = a × LogP + b × Hammett_sigma + c × electronic_parameter + constant. This was a breakthrough — it showed that simple, measurable physical parameters could predict complex biological effects. The 1964 paper has been cited thousands of times.

Modern QSAR uses ML instead of linear equations, but the concept is the same.

THE WORKFLOW (step by step):

Step 1 — DATA: Get a dataset of (molecule, activity) pairs. ChEMBL (maintained by EMBL-EBI) is the main source: it contains ~2.5 million compounds with experimental IC50/Ki/EC50 measurements against hundreds of protein targets. Download target data, clean duplicates, standardize units.

Step 2 — FEATURES: For each molecule, compute:
- Morgan fingerprints (2048-bit binary vector)
- Physicochemical descriptors (MW, LogP, HBD, HBA, TPSA...)
Concatenate into a single feature vector. Shape: (n_molecules, 2048+200) approximately.

Step 3 — SPLIT: Divide molecules into training set (80%) and test set (20%). CRITICAL: use scaffold split (see afternoon), not random split!

Step 4 — TRAIN: Fit RF or XGBoost on training set.

Step 5 — EVALUATE: Compute RMSE, R², etc. on test set. Look at the actual vs. predicted scatter plot.

Step 6 — INTERPRET: Use SHAP to understand which features drive predictions. Decode important fingerprint bits to chemical substructures.

WHY THIS MATTERS FOR AI DRUG DISCOVERY:
This pipeline is EXACTLY what is used in pharmaceutical companies and biotech today. Companies like Schrödinger, Insilico Medicine, and others built businesses on this pipeline. The practical skills students gain today are directly applicable in industry.""")

    # --- RF & XGBoost ---
    n += 1
    make_two_column_slide(prs,
        "ML Baselines: Random Forest & XGBoost",
        "Random Forest",
        ["Ensemble of decision trees (Breiman 2001)",
         "Each tree trained on bootstrap sample",
         "Final prediction = average (regression) or vote",
         "Handles high-dimensional fingerprints well",
         "Resistant to overfitting",
         "Easy to interpret (feature importance)"],
        "XGBoost",
        ["Sequential trees correct errors (Chen & Guestrin 2016)",
         "Often higher accuracy than RF",
         "Built-in regularization",
         "Handles missing values",
         "Dominant in Kaggle competitions",
         "Sheridan (2016, JCIM 56:2353-2360): XGBoost modestly > RF for QSAR"],
        n,
        notes="""=== RF AND XGBOOST — FULL EXPLANATION ===

WHY ARE WE USING THESE MODELS?
For molecular data (fingerprints + descriptors), Random Forest and XGBoost consistently outperform or match deep learning models on datasets with fewer than ~10,000 molecules. This is because:
1. Molecular datasets are often small (expensive to generate experimentally)
2. Tree models handle high-dimensional sparse data (fingerprints are 2048-bit but usually <5% set) very well
3. They don't need feature scaling (deep learning does)
4. They have excellent out-of-box performance with minimal hyperparameter tuning

RANDOM FOREST:
Invented by Leo Breiman (UC Berkeley, 2001). The idea:
1. Create B decision trees (typically 100-500)
2. Each tree is trained on a BOOTSTRAP SAMPLE: sample n molecules WITH replacement from the training set of n molecules. About 63% of unique molecules appear; 37% are 'out-of-bag' (OOB)
3. At each split in the tree, only consider sqrt(n_features) features (feature subsampling). This forces diversity between trees.
4. Final prediction: average of all trees' predictions (regression) or majority vote (classification).

WHY IT WORKS: Individual trees overfit, but they overfit to DIFFERENT aspects (due to bootstrap + feature sampling). Averaging cancels out the errors = low-variance ensemble.

FEATURE IMPORTANCE: The Gini importance (how much each feature reduces the variance across all trees). Can be a rough guide to which features matter most. But SHAP is more reliable (see afternoon).

XGBOOST:
Invented by Tianqi Chen (now at NVIDIA, 2016). XGBoost = eXtreme Gradient BOOSTing. The idea:
1. Start with a simple model (constant prediction = mean of training values)
2. Compute residuals (errors) for each training example
3. Train a NEW decision tree to predict the residuals
4. Add the new tree to the model (weighted by a learning rate η)
5. Compute new residuals. Repeat.
Each tree 'boosts' the previous ones by correcting their errors. Like a committee where each new member fixes what the previous members got wrong.

WHY XGBOOST IS OFTEN BETTER:
- Sequential learning can capture patterns that random forests miss
- Built-in L1 and L2 regularization prevents overfitting
- Handles missing values automatically (decides at each split which branch to take)
- Won many Kaggle competitions for tabular data

WHICH TO USE?
- Start with RF (easy, robust, fewer hyperparameters)
- Try XGBoost to potentially get better performance
- In practice: try both, choose the one with better validation performance
- In the practical today: we compare both on the same dataset

SCIKIT-LEARN CODE:
from sklearn.ensemble import RandomForestRegressor
rf = RandomForestRegressor(n_estimators=100, random_state=42)
rf.fit(X_train, y_train)
y_pred = rf.predict(X_test)

from xgboost import XGBRegressor
xgb = XGBRegressor(n_estimators=100, learning_rate=0.1, random_state=42)
xgb.fit(X_train, y_train)
y_pred = xgb.predict(X_test)""")

    # --- Break ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="15-MINUTE BREAK. Students should stretch, grab coffee, and mentally prepare for Practical 1. The practical uses the concepts from the last 90 minutes (SMILES parsing, Morgan fingerprints, descriptors, RF, XGBoost). Remind students before the break: after the break, open the notebook Day2_Practical_QSAR.ipynb.",
        subtitle="Practical 1 starts right after the break!")

    # --- Practical 1 ---
    n += 1
    make_content_slide(prs,
        "Practical 1: Build a QSAR Model",
        ["Open: day2_molecular_ml/Day2_Practical_QSAR.ipynb",
         "",
         "Dataset: Delaney solubility (1,128 molecules, logS values)",
         "  Delaney (2004), J. Chem. Inf. Comput. Sci. 44:1000-1005",
         "",
         "Steps:",
         "  1. Load dataset, parse SMILES with RDKit",
         "  2. Generate Morgan fingerprints (ECFP4, 2048 bits)",
         "  3. Calculate physicochemical descriptors",
         "  4. 80/20 train/test split",
         "  5. Train Random Forest regressor",
         "  6. Train XGBoost regressor",
         "  7. Compare: RMSE, R-squared",
         "",
         "Time: ~50 min including code execution | Work in pairs"],
        n,
        notes="""=== PRACTICAL 1 — WHAT STUDENTS ARE DOING AND WHY ===

THE DATASET — DELANEY SOLUBILITY:
John Delaney (AstraZeneca) published a dataset in 2004 with 1,128 small organic molecules and their experimentally measured aqueous solubility (log S, units: log mol/L). This is a classic benchmark for QSAR.

WHY SOLUBILITY? Aqueous solubility is one of the most important drug properties. About 40% of drug candidates fail because they are insoluble in water — they can't dissolve in the stomach and get absorbed. Predicting solubility from molecular structure (before synthesis!) would save enormous resources.

logS VALUES: typically range from -11 (almost insoluble) to +1 (very soluble). logS = -6 means 10^-6 mol/L ≈ 0.001 mg/mL (very insoluble). logS = 0 means 1 mol/L (very soluble, like sugar).

WHAT STUDENTS DO:
1. LOAD DATA: read the CSV file, look at the SMILES and logS columns
2. PARSE SMILES: for each SMILES string, use RDKit to create a molecule object: mol = Chem.MolFromSmiles(smiles). If the SMILES is invalid, this returns None — need to handle errors.
3. FINGERPRINTS: for each mol, compute Morgan fingerprints: fp = AllChem.GetMorganFingerprintAsBitVect(mol, 2, 2048). Convert to numpy: np.array(fp).
4. DESCRIPTORS: compute MW, LogP, HBD, HBA, TPSA for each molecule using RDKit Descriptors module. Stack these 5 numbers alongside the 2048-bit fingerprint. Final feature vector: 2053 dimensions per molecule.
5. SPLIT: use sklearn.model_selection.train_test_split with test_size=0.2, random_state=42
6. TRAIN RF: RandomForestRegressor(n_estimators=100, random_state=42).fit(X_train, y_train)
7. TRAIN XGBOOST: XGBRegressor(n_estimators=100, learning_rate=0.1).fit(X_train, y_train)
8. EVALUATE: compute RMSE and R² using sklearn.metrics.mean_squared_error and r2_score

EXPECTED RESULTS:
RF: RMSE ~0.6-0.8 logS, R² ~0.85-0.90
XGBoost: RMSE ~0.5-0.7 logS, R² ~0.87-0.92
These look great! BUT (afternoon spoiler) the random split inflates performance because similar molecules end up in both train and test. With scaffold split, performance drops. This 'aha moment' is the key learning of the afternoon.

COMMON STUDENT ERRORS TO WATCH FOR:
- Forgetting to handle None values from invalid SMILES
- Using the wrong feature matrix shape (need 2D array, not list)
- Confusing RMSE with MSE (RMSE = sqrt(MSE))
- Getting confused about what R² means (0 = baseline mean prediction, 1 = perfect, negative = worse than mean)""")

    # --- Practical 1 hands-on ---
    n += 1
    make_section_divider(prs, "🔬 PRACTICAL 1 — Build Your QSAR Model", n,
        notes="PRACTICAL 1 SESSION (~50 minutes). Walk around the room. Help students who are stuck. Common issues: import errors (make sure rdkit, sklearn, xgboost are installed), SMILES parsing errors (remind them to check for None), shape errors in numpy arrays. The goal: every student gets RF and XGBoost trained with RMSE and R² computed by the end.",
        subtitle="50 minutes of coding — train your first molecular ML model!")

    # --- Afternoon Agenda ---
    n += 1
    make_content_slide(prs,
        "Afternoon Agenda (Units 5-8)",
        ["7. Why standard evaluation fails for molecular data",
         "8. Activity cliffs and scaffold bias",
         "9. Proper validation: scaffold splits, temporal splits",
         "10. Metrics deep dive: RMSE, R², ROC-AUC, precision-recall",
         "11. Explainability with SHAP (Lundberg & Lee 2017)",
         "12. PRACTICAL 2: Re-evaluate your morning model properly",
         "",
         "☕ BREAK after Unit 6 (after ~90 min afternoon)",
         "",
         "Key message: the EVALUATION is more important than the model"],
        n,
        notes="""=== AFTERNOON AGENDA — WHAT WE'RE DOING AND WHY ===

The afternoon is about RIGOUR. The morning built a model that LOOKS good (R² ~0.9). The afternoon reveals that the model might not be as good as it looks — and teaches the correct way to evaluate.

WHY IS EVALUATION SO IMPORTANT?
In pharmaceutical drug discovery, a model that falsely appears to work well can guide expensive decisions. If your model says 'compound X is active against Target Y', and you spend $100,000 synthesizing and testing it, only to find it doesn't work — that's a very expensive false positive. If the model was trained and tested with data leakage (contaminated by scaffold similarity), all those AUC-ROC 0.95 results were illusory.

THE CORE PROBLEM:
Standard ML practice: randomly split data 80/20, train on 80%, test on 20%. Works great for image recognition (a test image of a cat is genuinely unseen). But for molecules: if molecule A and molecule B are structurally similar (same scaffold), and A is in the training set, then B in the test set is not really 'unseen' — the model has seen its scaffold. The test is easier than it would be for a truly novel molecule.

HOW TO FIX IT:
Scaffold split: cluster molecules by their core ring system (scaffold). Put all molecules with the same scaffold in the SAME split (either all train or all test). This ensures the test set contains genuinely novel scaffolds.

SHAP:
After proper evaluation, we want to understand WHAT the model learned. SHAP (SHapley Additive exPlanations) is the standard tool for this. It answers: for this specific molecule's prediction, which features contributed the most (and in which direction)?

THE 'AHA MOMENT' for students: compare their morning performance (random split, R² ~0.90) with their afternoon performance (scaffold split, R² ~0.75-0.80). That difference — 0.10-0.15 R² points — represents the 'scaffold memorization' the model was doing. The scaffold split reveals the true performance on novel chemical series.""")

    # --- Hall of Shame ---
    n += 1
    make_story_slide(prs,
        "The Hall of Shame: When Models Lie",
        "Example 1: hERG toxicity model — 95% accuracy on random split. On novel scaffolds: 60%.\n\nExample 2: Published AUC-ROC 0.97. Remove duplicate compounds: 0.72.\n\nExample 3: Scaffold split vs random — AUROC dropped from 0.92 to 0.75.\n\nExample 4 (Neuroscience): A GABA-A binding model looked great, but it learned to recognize the benzodiazepine scaffold, not actual binding features. On novel chemotypes: near-random.\n\nLesson: ALWAYS validate properly.\n\nRef: Wallach & Heifets (2018) — 'Most Ligand-Based Benchmarks Reward Memorization'",
        n,
        notes="""=== HALL OF SHAME — EXPLAINING EACH EXAMPLE ===

These are real failure modes from the drug discovery literature. Understanding them is CRITICAL for doing AI drug discovery properly.

EXAMPLE 1 — hERG MODEL:
hERG (pronounced 'h-ERG') = human Ether-à-go-go Related Gene. It encodes the Kv11.1 potassium channel in cardiac cells. If a drug blocks this channel, it can cause a life-threatening cardiac arrhythmia called 'Long QT syndrome' (QT = a measurement on an EKG/ECG that represents how long it takes the heart to electrically reset after each beat). Many promising drug candidates were withdrawn from clinical trials or market because they unexpectedly blocked hERG (terfenadine = Seldane, cisapride = Propulsid, etc.).

The hERG QSAR model: trained on all available hERG data with RANDOM split → 95% accuracy. But hERG binders are dominated by certain chemical scaffolds (basic nitrogen + aromatic ring). If those scaffolds are in both train and test, the model learns 'does it have this scaffold?' (not 'does it bind hERG?'). On scaffolds not seen in training: 60%. NOT MUCH BETTER THAN RANDOM!

EXAMPLE 2 — DUPLICATE COMPOUNDS:
Some databases have duplicate entries (the same molecule measured multiple times, or structurally similar molecules). If you don't remove duplicates before splitting, you literally have the same molecule in train and test. Published AUC 0.97 → remove duplicates → 0.72. The published result was meaningless.

EXAMPLE 3 — GENERAL SCAFFOLD SPLIT VS. RANDOM:
This is the most common issue. AUC drops from 0.92 to 0.75 when properly evaluated. Citation: Wallach & Heifets (2018) showed this in a systematic study of multiple drug discovery benchmarks. Almost all published benchmarks at that time were overestimating performance.

EXAMPLE 4 — GABA-A MODEL (neurologically relevant):
GABA-A is the main inhibitory receptor in the brain. It's targeted by: benzodiazepines (diazepam/Valium → anxiety, sleep), barbiturates (old sedatives), Z-drugs (zolpidem/Ambien → insomnia), neurosteroids (naturally occurring anesthetics), alcohol.

Most GABA-A binding data is from benzodiazepines (they were extensively studied in the 1970s-1990s). A QSAR model trained on this data learns to recognize the 1,4-benzodiazepine scaffold. It predicts that anything with a 1,4-benzodiazepine skeleton will bind GABA-A. TRUE for all the training data, but MISLEADING for novel chemotypes (neurosteroids, Z-drugs, THIP, etc.). On scaffold split: near-random performance for non-benzodiazepine chemotypes.

THE LESSON:
Every step between 'I have a great AUC' and 'I should trust this model in drug discovery' requires rigorous validation. In the afternoon practical, students will directly OBSERVE this performance drop with scaffold splits.""")

    # --- Scaffold Bias ---
    n += 1
    make_content_slide(prs,
        "Scaffold Bias and Proper Splitting",
        ["Scaffold = core ring structure (Bemis-Murcko framework)",
         "Problem: random split → same scaffold in train AND test → memorization",
         "",
         "Solutions:",
         "  Scaffold split: no scaffold overlap (Bemis & Murcko 1996)",
         "  Temporal split: train on older, test on newer data",
         "  Cluster split: cluster molecules, split by cluster",
         "",
         "MoleculeNet uses scaffold split as default (Wu et al. 2018)",
         "",
         "Activity cliffs: very similar molecules, very different activity",
         "  Changing one methyl → chlorine can change IC50 by 1000×",
         "  These are the hardest cases for ML!"],
        n,
        notes="""=== SCAFFOLD BIAS AND PROPER SPLITTING — FULL EXPLANATION ===

WHAT IS A SCAFFOLD?
In medicinal chemistry, a 'scaffold' is the core structural framework of a drug molecule — typically the ring system (cycles) plus the atoms connecting them, with side chains removed.

The Bemis-Murcko framework (Bemis & Murcko 1996, J. Med. Chem. 39:2887-2893): strip off all side chains (substituents), keep the ring systems and the linker atoms connecting them. The result is the 'scaffold'.

Example: Aspirin (benzene ring + two side chains), ibuprofen (different side chains but a similar propionic acid + aryl scaffold), naproxen (naphthalene ring + propionic acid). These three have different scaffolds. Aspirin's scaffold = benzene. Ibuprofen's scaffold = isobutylbenzene. If aspirin is in train and ibuprofen is in test, they share some structural similarity but are 'scaffold different'.

THE PROBLEM WITH RANDOM SPLITS:
Drug discovery datasets often contain chemical series — groups of related compounds where medicinal chemists systematically varied substituents on a core scaffold (e.g., series of 50 compounds all with the same quinazoline core but different R groups). If you randomly split, many of these series will have molecules in BOTH train and test. The model learns to recognize the scaffold → predicts test set well → but this performance doesn't generalize to novel scaffolds.

SCAFFOLD SPLIT IMPLEMENTATION:
1. For each molecule, compute the Bemis-Murcko scaffold using RDKit: from rdkit.Chem.Scaffolds import MurckoScaffold; scaffold = MurckoScaffold.GetScaffoldForMol(mol)
2. Group molecules by scaffold
3. Put each scaffold entirely into train OR test (never split)
4. Fill to 80/20 split approximately (start with largest scaffold groups in train, fill test until ~20% is reached)

TEMPORAL SPLIT:
Instead of structural split, split by TIME. Train on all data before date X, test on all data after date X. This mimics the real use case: you train the model in year Y, use it to predict in year Y+1 (on molecules that don't exist yet).

ACTIVITY CLIFFS:
The term for when structurally VERY similar molecules have VERY different activities. Example: compound A (benzene ring + NH2 group) has IC50 = 1 nM (very potent). Compound B (same benzene ring + NHMe group, just added one methyl) has IC50 = 1 μM (1000× weaker).

Activity cliffs arise because: a single functional group (like NH2 vs. NMe) can make a completely different hydrogen bond. NH2 can donate TWO hydrogen bonds (two N-H bonds). NMe can donate ZERO (no N-H bond). If the receptor requires an N-H hydrogen bond for binding, adding the methyl destroys activity.

WHY ACTIVITY CLIFFS ARE HARD FOR ML:
Fingerprints make these two compounds look similar (Tanimoto ~0.9). But their activities differ by 1000×. The model can't explain this from fingerprints alone — it would need the 3D structure of the binding pocket (molecular docking, Day 3) to understand why the methyl group is so destructive.""")

    # --- Metrics ---
    n += 1
    make_two_column_slide(prs,
        "Evaluation Metrics for Molecular ML",
        "Regression",
        ["RMSE — root mean squared error",
         "  Penalizes large errors more",
         "R² — coefficient of determination",
         "  1.0 = perfect, <0 = worse than mean",
         "MAE — mean absolute error",
         "  More robust to outliers",
         "",
         "Report ALL three, not just R²!"],
        "Classification",
        ["AUC-ROC — area under ROC curve",
         "  0.5 = random, 1.0 = perfect",
         "  Insensitive to class balance!",
         "Precision-Recall AUC",
         "  Better for imbalanced data (most drug data)",
         "Balanced accuracy",
         "  Average of sensitivity + specificity",
         "",
         "Report PR-AUC for imbalanced data!"],
        n,
        notes="""=== EVALUATION METRICS — FULL EXPLANATION ===

REGRESSION METRICS:

RMSE (Root Mean Squared Error):
RMSE = sqrt(mean((y_true - y_pred)²))
For solubility prediction: if RMSE = 0.7 logS, it means the average prediction error is 0.7 log units. Since each log unit = 10× in solubility, RMSE = 0.7 means the model's average prediction is off by ~5× (10^0.7 ≈ 5).
Why square first, then root? Squaring penalizes large errors more than small ones (2² = 4, 4² = 16). A model with many small errors is better than one with a few catastrophic errors. Root brings it back to the original units (logS).

R² (R-squared, Coefficient of Determination):
R² = 1 - (SS_res / SS_tot)
where SS_res = sum of squared residuals = sum((y_true - y_pred)²)
SS_tot = total sum of squares = sum((y_true - mean(y_true))²)
Interpretation:
R² = 1.0: perfect predictions (all residuals = 0)
R² = 0: model no better than always predicting the mean value
R² < 0: model WORSE than always predicting the mean!
Typically for good QSAR models: R² > 0.7 is acceptable, R² > 0.85 is good.

MAE (Mean Absolute Error):
MAE = mean(|y_true - y_pred|)
Simpler interpretation than RMSE (average absolute error). More robust to outliers (extreme errors are not squared, so they have less influence). Useful if you have noisy experimental data.

CLASSIFICATION METRICS:

AUC-ROC (Area Under the Receiver Operating Characteristic Curve):
The ROC curve plots True Positive Rate (TPR = sensitivity) vs. False Positive Rate (FPR = 1 - specificity) at different decision thresholds.
AUC = 0.5: random classifier (ROC curve = diagonal)
AUC = 1.0: perfect classifier
AUC = 0.8: good classifier (80% chance that a randomly chosen positive example will be scored higher than a randomly chosen negative)

PROBLEM: AUC-ROC is insensitive to class imbalance. If 99% of compounds are inactive and 1% are active, a model that always predicts 'inactive' achieves ROC-AUC ≈ 0.5 (not great) but PR-AUC ≈ 0.01 (reveals it's useless). Drug discovery data is almost always imbalanced (many more inactive compounds than active ones).

PRECISION-RECALL AUC (PR-AUC):
Precision = true positives / (true positives + false positives) = of the things I predicted positive, how many were actually positive?
Recall = true positives / (true positives + false negatives) = of all the actual positives, how many did I find?
PR-AUC = area under the Precision-Recall curve. For perfectly balanced data: random classifier = 0.5. For imbalanced data with 5% positives: random classifier ≈ 0.05. Much more informative for drug discovery.

PRACTICAL RECOMMENDATION:
For QSAR regression: report RMSE, R², and MAE together.
For activity classification: report PR-AUC and ROC-AUC both.
For comparison across datasets with different imbalance ratios: use balanced accuracy or Matthew's Correlation Coefficient (MCC).""")

    # --- Interpretability Section ---
    n += 1
    make_section_divider(prs, "Model Interpretability", n,
        notes="Transition slide. After evaluating performance quantitatively (RMSE, R², AUC), we want to understand WHY the model makes specific predictions. This is the SHAP section (next two slides).",
        subtitle="Opening the black box — why does the model predict THIS?")

    # --- Why Interpretability ---
    n += 1
    make_content_slide(prs,
        "Why Explainability Matters",
        ["Regulatory: FDA expects understanding of model behavior",
         "Scientific: chemists won't trust a model they can't understand",
         "Safety: unexplained predictions could mask dangerous failures",
         "Discovery: understanding features drives new hypotheses",
         "",
         "Jimenez-Luna et al. (2020), Nature Machine Intelligence 2:573-584",
         "",
         "Types:",
         "  Global: which features are generally important?",
         "  Local: why did the model predict THIS for THIS molecule?",
         "",
         "EU AI Act: medical AI = high-risk = requires transparency"],
        n,
        notes="""=== WHY EXPLAINABILITY MATTERS — FULL EXPLANATION ===

The core tension in modern AI: deep learning models are extremely powerful but hard to interpret ('black boxes'). In drug discovery, a black box is a liability, not an asset.

REGULATORY (FDA):
The FDA is increasingly paying attention to AI-driven drug discovery. They have published guidance on 'Good Machine Learning Practice' for medical devices (including AI tools in drug development). The key expectation: you must be able to explain WHY the model makes a prediction, especially for safety-critical decisions (predicting toxicity, predicting which patients respond).

SCIENTIFIC CREDIBILITY:
A medicinal chemist with 30 years of experience will not trust a model that says 'synthesize compound X — I predict it's active but I can't explain why.' They will ask: 'What structural features drive that prediction?' If you can say 'the model predicts this compound is active because of the hydroxyl group at position 4 (SHAP value +1.5) and the fluorine at position 2 (SHAP +0.8), while the large alkyl chain is hurting it (-0.4)', then the chemist can engage. They might say 'yes, that makes sense — I've seen that hydroxyl group be important in this receptor family.' Or they might say 'that fluorine shouldn't matter — there's no electrophilic site in the receptor pocket. I suspect the model is wrong.' Either way, the interpretation drives a scientific discussion.

SAFETY:
If a model predicts 'non-toxic' for a compound, but can't explain why, you're making a safety decision on faith. If the model can say 'predicted non-toxic because: low LogP (-0.3), no nitro groups (-1.0), no reactive Michael acceptors (-0.5)', you can check: 'Are these the right reasons?' If the model is making the right prediction for the wrong reasons, it will fail on new chemical types.

DISCOVERY (new science):
When SHAP says 'for this antiviral drug series, the fluorine at position 3 is consistently the most important feature', you've discovered a chemical design rule. Maybe no one knew this before. This is AI contributing to medicinal chemistry knowledge, not just prediction.

GLOBAL VS. LOCAL EXPLANATIONS:
Global: 'What are the most important features across all predictions?' → SHAP beeswarm plot (1000 points, each molecule, each dot colored by feature value, showing distribution of SHAP values).
Local: 'Why did the model predict IC50 = 5 nM for THIS specific molecule?' → SHAP waterfall plot (single prediction, showing positive and negative contributions of each feature).

EU AI ACT:
The European Union's AI Act (2023) classifies medical AI as 'high risk'. High-risk AI systems require: transparency (documentation of training data, model design, performance), explainability (ability to understand why decisions are made), robustness (performance on diverse inputs), and human oversight (humans can override). This will affect all AI drug discovery tools used in Europe.""")

    # --- SHAP ---
    n += 1
    make_content_slide(prs,
        "SHAP: SHapley Additive exPlanations",
        ["Lundberg & Lee (2017), NeurIPS",
         "Based on Shapley values from cooperative game theory",
         "",
         "For each prediction, assign a contribution to each feature:",
         "  Positive SHAP → pushes prediction higher",
         "  Negative SHAP → pushes prediction lower",
         "  Sum of all SHAP values + baseline = predicted value",
         "",
         "Advantages:",
         "  Mathematically grounded (unique solution with desirable properties)",
         "  Works for any model (model-agnostic)",
         "  Local AND global explanations",
         "  Beautiful visualizations (beeswarm, waterfall, force plots)",
         "",
         "For fingerprints: important bit → decode to substructure → chemistry insight"],
        n,
        notes="""=== SHAP — FULL EXPLANATION FROM GAME THEORY TO MOLECULAR ML ===

GAME THEORY BACKGROUND:
Lloyd Shapley (Princeton, 1953) asked: 'If several players cooperate to earn a reward, how should we distribute the reward fairly?' His answer: each player's fair share = their AVERAGE MARGINAL CONTRIBUTION across all possible orderings of players joining the coalition.

EXAMPLE: Three players A, B, C earn rewards based on who is in the coalition:
{A}=3, {B}=2, {C}=4, {A,B}=7, {A,C}=8, {B,C}=6, {A,B,C}=12
Shapley value for A: average of A's marginal contribution in all 6 orderings (3! = 6 permutations of A,B,C joining the coalition): in some orderings A joins first (marginal contribution = 3), in others A joins after B (marginal = 7-2=5), etc. The exact Shapley values are fair and unique (they are the ONLY values satisfying efficiency, symmetry, dummy player, and additivity axioms).

APPLYING TO ML:
Players = features. Payout = prediction. SHAP value for feature i = the average marginal contribution of feature i to the prediction across all possible subsets of features.

FORMULA: SHAP_i = sum over all subsets S not containing i of [|S|! × (|F|-|S|-1)! / |F|!] × [f(S∪{i}) - f(S)]
Where F is the set of all features and f(S) is the model's prediction when only features in S are present.

WHY THIS IS EXPENSIVE: There are 2^n subsets (n = number of features). For 2048-bit fingerprints, this is astronomical. TreeSHAP (Lundberg et al. 2020) computes exact Shapley values for tree-based models (RF, XGBoost) in O(T × L²) time where T = number of trees and L = max number of leaves. Fast and exact!

HOW TO READ SHAP VALUES:
For a molecule predicted to have pIC50 = 7.8:
- Baseline (mean prediction for all training molecules) = 6.5
- SHAP value for 'aromatic amine group' = +1.0 (pushes prediction up by 1.0)
- SHAP value for 'high molecular weight' = -0.3 (pushes it down)
- SHAP value for 'fluorine on ring' = +0.6
- Sum: 6.5 + 1.0 - 0.3 + 0.6 + ... = 7.8 (predicted value)

DECODING FINGERPRINT BITS:
If bit 847 has the highest SHAP value, you want to know what molecular feature it represents. In RDKit:
from rdkit.Chem import AllChem
bitInfo = {}
fp = AllChem.GetMorganFingerprintAsBitVect(mol, radius=2, nBits=2048, bitInfo=bitInfo)
# bitInfo[847] gives you the (atom_index, radius) pair
# Draw.DrawMorganBit(mol, bit=847, bitInfo=bitInfo) shows the substructure
This is how you go from 'bit 847 matters' to 'the pyridine ring at position X matters'.

VISUALIZATIONS:
- Beeswarm plot: shows SHAP values for all molecules and all features. Each dot = one molecule. X-axis = SHAP value. Color = feature value (red=high, blue=low). Most informative global view.
- Waterfall plot: for one specific molecule. Shows exactly which features pushed the prediction up/down from the baseline to the final value.
- Force plot: like waterfall but horizontal. Good for presentations.

CODE:
import shap
explainer = shap.TreeExplainer(rf_model)
shap_values = explainer.shap_values(X_test)
shap.summary_plot(shap_values, X_test)  # beeswarm""")

    # --- Applicability Domain ---
    n += 1
    make_content_slide(prs,
        "Applicability Domain",
        ["A model is only reliable within its training data distribution",
         "",
         "Methods to define AD:",
         "  Distance-based: is the new molecule close to training data?",
         "  Descriptor range: are features within training bounds?",
         "  Conformal prediction: distribution-free prediction intervals",
         "",
         "Sahigara et al. (2012), Molecules 17(5):4791-4810",
         "",
         "In practice: flag out-of-domain predictions as low confidence",
         "A confident wrong prediction is worse than an honest 'I don't know'"],
        n,
        notes="""=== APPLICABILITY DOMAIN — FULL EXPLANATION ===

WHAT IS THE APPLICABILITY DOMAIN?
Every ML model has a 'comfort zone' — the region of chemical space where it has training data and can make reliable predictions. Outside this zone (the 'applicability domain'), predictions become unreliable.

ANALOGY: A QSAR model trained on kinase inhibitors (cancer drugs that block protein kinases) should NOT be used to predict activity of GPCRs. The model has no training data in that chemical space. Similarly, a model trained on compounds with MW 200-400 Da should not be applied to a 900 Da macrocycle without flagging it as 'out of domain'.

METHODS:

1. Distance-based:
For each new compound, compute its similarity to the nearest training compound(s) using Tanimoto similarity on fingerprints. If max_Tanimoto < threshold (e.g., 0.3), the compound is too dissimilar from anything in training → flag as outside AD.

2. Descriptor range check:
For each physicochemical descriptor, check if the new compound's value is within [min_train, max_train]. If LogP = 8 but training data had LogP 0-5, the compound is outside the training distribution for this descriptor.

3. Leverage h statistic:
From linear regression statistics. h_i = x_i^T (X^T X)^{-1} x_i. High h = influential point, far from the center of the training data. Williams plot: SHAP residual vs. h_i. Points with high h AND high residual = problem predictions.

4. Conformal prediction:
A rigorous statistical framework. Gives prediction INTERVALS with guaranteed coverage: 'I predict pIC50 = 7.5 ± 1.2, and I guarantee that 95% of the time, the true value is within this interval.' Based on p-values from comparing the new point to calibration set points. Does not require strong statistical assumptions.

PRACTICAL IMPLEMENTATION:
from sklearn.neighbors import NearestNeighbors
nn = NearestNeighbors(n_neighbors=5, metric='jaccard')  # Tanimoto = 1 - Jaccard
nn.fit(X_train_fps)
distances, _ = nn.kneighbors(X_test_fps)
ad_flag = distances.mean(axis=1) > 0.7  # True = outside AD

THE KEY PRINCIPLE:
A model that SAYS 'I don't know' is MORE VALUABLE than a model that confidently gives wrong answers. In pharmaceutical development, a false positive prediction (we predicted active, it's actually inactive) costs millions. Flagging uncertain predictions as 'low confidence' prevents these costly mistakes.""")

    # --- Break ---
    n += 1
    make_section_divider(prs, "☕ BREAK — 15 minutes", n,
        notes="SECOND 15-MINUTE BREAK of Day 2. Students should have their morning model trained (from Practical 1). The afternoon practical (Practical 2) will: (1) implement scaffold split, (2) re-evaluate the morning model with this proper split, and (3) compute SHAP values. Remind: open Day2_Practical_Evaluation.ipynb after the break.",
        subtitle="Final practical of the day coming up!")

    # --- Practical 2 ---
    n += 1
    make_content_slide(prs,
        "Practical 2: Evaluate & Interpret Your QSAR Model",
        ["Open: day2_molecular_ml/Day2_Practical_Evaluation.ipynb",
         "",
         "Steps:",
         "  1. Load your morning QSAR model (or re-train quickly)",
         "  2. Implement scaffold splitting (Bemis-Murcko)",
         "  3. Compare: random split vs. scaffold split performance",
         "  4. Compute ROC curve, precision-recall curve",
         "  5. Calculate SHAP values for your RF model",
         "  6. Create SHAP beeswarm plot: which features matter most?",
         "  7. Decode important fingerprint bits to substructures",
         "  8. Define applicability domain using Tanimoto distance",
         "",
         "Time: ~50 min | The 'aha moment': watch performance DROP with scaffold split"],
        n,
        notes="""=== PRACTICAL 2 — WHAT STUDENTS DO AND WHAT TO EXPECT ===

SETUP: Students should have their morning model (RF + XGBoost) ready from Practical 1. If someone didn't finish, they can quickly re-train in the first 5 minutes.

STEP 2 — SCAFFOLD SPLIT:
from rdkit.Chem.Scaffolds import MurckoScaffold
def get_scaffold(smiles):
    mol = Chem.MolFromSmiles(smiles)
    scaffold = MurckoScaffold.MurckoScaffoldSmiles(mol=mol, includeChirality=False)
    return scaffold

Group molecules by scaffold. Sort scaffold groups by size (largest first). Put largest scaffolds in train until 80% filled. Rest in test.

STEP 3 — THE AHA MOMENT:
Students will see their R² drop from ~0.85-0.90 (random split) to ~0.70-0.80 (scaffold split). This is the KEY LEARNING. The model was memorizing scaffold patterns, not learning fundamental structure-activity relationships. On novel scaffolds, performance is worse.

Make sure to discuss this observation with the class after the practical. Ask: 'Why did your R² drop?' 'What does this mean for using your model to predict a truly new molecule?'

STEP 5-7 — SHAP:
import shap
explainer = shap.TreeExplainer(rf_model)
shap_values = explainer.shap_values(X_test)  # shape: (n_test, n_features)

COMMON FINDINGS:
- The most important features are often LogP, molecular weight, and a few specific fingerprint bits
- LogP and MW show negative SHAP for high values (more lipophilic/heavier molecules tend to be LESS soluble)
- Some fingerprint bits corresponding to aromatic rings or specific functional groups consistently show positive SHAP
- The SHAP beeswarm plot makes chemical sense: high LogP → high SHAP magnitude for LogP feature → molecules that are very lipophilic are predicted to be less soluble (correct!)

STEP 8 — APPLICABILITY DOMAIN:
Simple implementation: for each test molecule, find its nearest training neighbor (Tanimoto similarity). Flag molecules where similarity < 0.3 as 'outside AD'.

EXPECTED OUTCOMES:
By the end of Practical 2, students should have:
1. A properly evaluated QSAR model (scaffold split R² ~0.75-0.80)
2. A SHAP beeswarm plot showing feature importances
3. Identified the most important molecular features (LogP, MW, some fingerprint bits)
4. An applicability domain flag for each test molecule

THINGS TO HELP WITH:
- shap.TreeExplainer with RF: use rf_model.estimators_ if there are issues with the ensemble. Or use shap.Explainer (newer API)
- SHAP plots can be slow for large datasets: subsample to 200-500 test molecules
- Fingerprint bit decoding: see the bitInfo dictionary from GetMorganFingerprintAsBitVect()""")

    # --- Practical 2 hands-on ---
    n += 1
    make_section_divider(prs, "🔬 PRACTICAL 2 — Evaluate & Interpret", n,
        notes="PRACTICAL 2 SESSION (~50 minutes). Walk around and help. The scaffold split implementation is the trickiest part — help students who get stuck on the MurckoScaffold API. The SHAP beeswarm plot is the most visually rewarding part — it's always a crowd-pleaser when you can see that LogP and MW are the most important features and they make chemical sense.",
        subtitle="50 minutes — discover what your model really knows!")

    # --- Discussion ---
    n += 1
    make_discussion_slide(prs,
        "A model achieves AUC-ROC 0.95 on the test set.\nWould you trust it to guide a $100M drug campaign?",
        ["What kind of split was used? Random or scaffold?",
         "How imbalanced is the dataset? (Use PR-AUC instead!)",
         "Has it been validated prospectively on truly new chemical series?",
         "What is the applicability domain?"],
        n,
        notes="""=== DISCUSSION QUESTIONS — GUIDE FOR DISCUSSION ===

The answer to the main question is: NO, not without more information.

WALK THROUGH EACH QUESTION:

Q: What kind of split was used?
Good answer: scaffold split, ideally temporal split. Random split overfits. The student who just did Practical 2 saw their R² drop when they used scaffold split. The same applies to AUC-ROC. A published AUC 0.95 with random split might be 0.75 with scaffold split.

Q: How imbalanced is the dataset?
Drug discovery datasets are typically very imbalanced: 1-5% active, 95-99% inactive. In this case, AUC-ROC is misleading because it weights false positives and false negatives equally, ignoring class imbalance. PR-AUC is much more informative. A model with AUC-ROC 0.95 on a 1% positive rate dataset might have PR-AUC 0.15 (terrible — most of its 'active' predictions are false positives).

Q: Prospective validation?
Has the model been used in practice to predict TRULY NEW molecules (synthesized AFTER the model was built)? Not just retrospective validation on historical data. This is the ultimate test. If the model successfully guided synthesis of 5 novel actives that were predicted before synthesis, that's strong evidence. If it's only been evaluated on historical data: much weaker evidence.

Q: Applicability domain?
If the model is asked to predict for a molecule very different from its training set, the prediction is unreliable. Did anyone check whether the test set molecules are in the AD? Many published benchmarks don't report AD analysis.

CLOSING THOUGHT:
The answer 'I would not trust a $100M decision to this model without more validation' is NOT pessimism about AI. It's scientific rigor. Models that are properly validated — with scaffold splits, checked AD, prospective validation — ARE trustworthy and ARE used to guide real drug campaigns (Insilico Medicine, Schrödinger, Recursion Pharmaceuticals all do this). The key is: validate properly before making expensive decisions.""")

    # --- Takeaways ---
    n += 1
    make_takeaway_slide(prs,
        ["Morgan/ECFP fingerprints convert molecules to bit vectors",
         "QSAR: predicting activity from structure (since 1964!)",
         "RF and XGBoost: strong, production-ready QSAR baselines",
         "Scaffold splits prevent data leakage (Wallach & Heifets 2018)",
         "SHAP: mathematically grounded explanations for any model",
         "Applicability domain: know your model's limits",
         "The evaluation is more important than the model!",
         "Next (Day 3, 5 May): Deep learning, GNNs, generative AI, AlphaFold"],
        2, n,
        notes="""=== KEY TAKEAWAYS — EMPHASIZE THESE ===

1. MORGAN/ECFP FINGERPRINTS:
The most important molecular representation for classical ML. Converts any molecule (SMILES) to a 2048-bit binary vector. Computed with 3 lines of RDKit code. Used in virtually every industrial QSAR model.

2. QSAR IS 60 YEARS OLD:
It's not new! Hansch (1964) showed that structure predicts activity. What's new is using deep learning (Day 3) and much larger datasets (ChEMBL). The fundamental idea is unchanged.

3. RF AND XGBOOST:
For datasets <10,000 molecules, these are often as good as deep learning. Students can use these RIGHT NOW for their projects. Both are in scikit-learn (RF) and xgboost library (XGBoost).

4. SCAFFOLD SPLITS:
This is the most common mistake in published molecular ML papers. Random split inflates performance by 10-20% R² points. Always use scaffold split for molecular property prediction. MoleculeNet (the main benchmark) uses scaffold split as default.

5. SHAP:
The gold standard for interpreting ML predictions. Mathematically rigorous (Shapley values), model-agnostic (works for RF, XGBoost, neural networks), and connects computer science (ML prediction) to chemistry (substructure importance). TreeSHAP is exact and fast for RF/XGBoost.

6. APPLICABILITY DOMAIN:
Know when to say 'I don't know.' A model outside its training distribution can give confident wrong answers. Flag low-similarity predictions as uncertain.

7. THE EVALUATION IS MORE IMPORTANT THAN THE MODEL:
A mediocre model properly evaluated is more valuable than an impressive model improperly evaluated. This is the key practical lesson of today. Industry hiring managers will ask about validation strategies — not just what model you used.

8. WHAT'S COMING ON DAY 3:
GNNs (Graph Neural Networks): instead of Morgan/ECFP hashing, LEARN how to aggregate neighborhood information. AlphaFold: predict protein 3D structure from sequence, enabling structure-based drug design. Generative AI: design new molecules from scratch.""")

    # --- Next Day ---
    n += 1
    make_content_slide(prs,
        "Coming Up: Day 3 — 5 May 2026",
        ["Morning: Deep Learning & Graph Neural Networks",
         "  Molecules as graphs → message passing → GNNs",
         "  Practical: train a GNN on MoleculeNet data",
         "",
         "Afternoon: Generative AI & Protein Targets",
         "  Generate new molecules (VAE, diffusion)",
         "  AlphaFold for drug targets",
         "  Practical: explore generative models or AlphaFold",
         "",
         "Preparation:",
         "  Review today's SHAP analysis — understand your model",
         "  Start discussing project approach with your group"],
        n,
        notes="""=== PREVIEW OF DAY 3 — WHAT IS COMING AND HOW IT BUILDS ON TODAY ===

GRAPH NEURAL NETWORKS (GNNs):
Today: Morgan/ECFP fingerprints hash the neighborhood information into a fixed vector.
Day 3: GNNs LEARN how to aggregate the neighborhood information (message passing with trainable weights). The connection: ECFP is essentially one round of message passing with hashing. GNNs do it with neural networks instead. More expressive, learns task-specific features, but requires more data and more tuning.

KEY MESSAGE: Today's lecture is NECESSARY before Day 3. Students who understand ECFP fingerprints will immediately understand why GNNs are better. Students who don't understand ECFP will be lost.

GENERATIVE AI:
VAE (Variational Autoencoder): encode molecule to latent vector → modify latent vector → decode to new molecule. Can generate molecules with specific properties.
Diffusion models: add noise to a molecule, learn to denoise → generate new molecules from noise. State of the art for molecular generation.

ALPHAFOLD:
Google DeepMind's protein structure prediction AI (2021). Given a protein sequence, predict its 3D structure with atomic accuracy. REVOLUTIONIZED drug discovery because:
1. Drug binding depends on 3D structure of the protein target
2. Before AlphaFold, getting a 3D structure required experimental crystallography (expensive, slow, often fails)
3. Now: any protein's structure is available in minutes
4. AlphaFold + molecular docking (predicting where a drug fits in the 3D structure) = structure-based drug design

WHAT STUDENTS SHOULD DO BETWEEN NOW AND DAY 3:
1. Review their SHAP analysis from today's practical. Try to explain the top 5 most important features chemically.
2. Finalize their project groups (sign up today!).
3. Think about whether their project uses QSAR (today's tools) or needs GNNs (Day 3 tools). If the dataset is <5,000 molecules: RF/XGBoost might be sufficient. If >10,000 or involves 3D structure: consider GNNs/AlphaFold.""")

    os.makedirs("day2_molecular_ml", exist_ok=True)
    prs.save("day2_molecular_ml/slides_annotated.pptx")
    print(f"  Day 2 (Annotated): {n} slides → day2_molecular_ml/slides_annotated.pptx")



# ===== MAIN — see bottom of file =====


# =====================================================================
# QUIZ 3  Physiology, Neuroscience & SpikerBox
# =====================================================================
def generate_quiz_physiology():
    prs = new_prs()
    n = 0

    n += 1
    make_title_slide(prs,
        "Physiology & Neuroscience Quiz",
        "Action potentials, ion channels, Hodgkin-Huxley, neural circuits, physiological drug readouts",
        1, "AI for Drug Discovery", date_str="22 April 2026",
        notes="Use this quiz at the END of the Day 2 physiology intro, or at the start of the afternoon. 10 questions, ~20 minutes. Tests whether students absorbed the morning physiology material. Questions go from basic vocabulary to mechanistic drug pharmacology.")

    # Q1
    n += 1
    make_quiz_question_slide(prs, 1,
        "What is the resting membrane potential of a typical neuron?",
        ["+40 mV (inside more positive than outside)",
         "0 mV (no difference across membrane)",
         "-70 mV (inside more negative than outside)",
         "-200 mV (very strongly negative)"],
        n, notes="Basic vocabulary. Most CS students won't know this. Use the answer to explain the ion gradient concept.")

    n += 1
    make_quiz_answer_slide(prs, 1,
        "What is the resting membrane potential of a typical neuron?",
        ["+40 mV (inside more positive than outside)",
         "0 mV (no difference across membrane)",
         "-70 mV (inside more negative than outside)",
         "-200 mV (very strongly negative)"],
        2,
        "Correct: -70 mV. Inside of a resting neuron is ~70 millivolts more negative than outside.\nMaintained by the Na+/K+ ATPase pump: 3 Na+ out, 2 K+ in per ATP. K+ leaks out through resting K+ channels. When an AP fires, membrane rapidly depolarizes to +40 mV, then repolarizes.",
        n,
        notes="""WHY -70 mV?
The Na+/K+ pump moves 3 Na+ OUT and 2 K+ IN per ATP cycle. Net result: positive charge accumulates outside. K+ channels are also open at rest — K+ flows out (concentration gradient), leaving negative charge inside.
Resting K+ channels make the membrane selectively permeable to K+ at rest. This brings the membrane toward E_K (about -90 mV). The pump and the leak together settle at ~-70 mV.

DRUG RELEVANCE: local anesthetics, antiepileptics, and cardiac drugs all target the channels that control the resting potential and the action potential. Understanding the baseline helps you understand what the drug is changing.""")

    # Q2
    n += 1
    make_quiz_question_slide(prs, 2,
        "During the RISING phase of an action potential, which ion flows INTO the neuron?",
        ["K+ (potassium) — driven by its concentration gradient",
         "Cl- (chloride) — driven inward by electrical gradient",
         "Na+ (sodium) — through voltage-gated Na+ channels",
         "Ca2+ (calcium) — through L-type calcium channels"],
        n, notes="Tests mechanism of action potential. Na+ drives the rising phase. Ca2+ enters at synaptic terminals but is NOT the main depolarising ion.")

    n += 1
    make_quiz_answer_slide(prs, 2,
        "During the RISING phase of an action potential, which ion flows INTO the neuron?",
        ["K+ (potassium) — driven by its concentration gradient",
         "Cl- (chloride) — driven inward by electrical gradient",
         "Na+ (sodium) — through voltage-gated Na+ channels",
         "Ca2+ (calcium) — through L-type calcium channels"],
        2,
        "Correct: Na+ (sodium). Voltage-gated Na+ channels (Nav) open at threshold (~-55 mV).\nNa+ is 145 mM outside, 12 mM inside: BOTH concentration gradient AND electrical gradient (negative inside) drive Na+ in. Membrane shoots from -70 mV to +40 mV in < 1 ms.",
        n,
        notes="""THE RISING PHASE MECHANISM:
Step 1: stimulus depolarises past -55 mV threshold
Step 2: voltage-gated Nav snap open. The S4 helix in each domain acts as a voltage sensor — positively charged arginines move in the electric field when voltage changes, pulling the channel gate open.
Step 3: Na+ floods in. Two driving forces (thermodynamic): (a) concentration gradient: 145 mM outside vs 12 mM inside = 12x higher outside; (b) electrical gradient: inside is -70 mV, Na+ (positive ion) is pulled toward negative interior.
Step 4: membrane shoots to +40 mV (close to E_Na = +55 mV) in < 1 ms.

WHY NOT Ca2+? Ca2+ DOES enter at synaptic terminals to trigger neurotransmitter release. But this uses Cav2.1/2.2 channels, not the Nav channels of the axon. In cardiac muscle, Ca2+ does drive the plateau phase of the cardiac action potential. But in neurons, Na+ drives the spike.

DRUG TARGETS (Nav):
- Lidocaine/bupivacaine: local anesthetics. Block Nav. No spike = no pain signal.
- Carbamazepine/phenytoin: anticonvulsants. Block Nav preferentially in rapidly-firing neurons (use-dependent block). Reduce seizure frequency.
- Tetrodotoxin (TTX): from pufferfish. Extremely potent Nav blocker. Research tool. Deadly at nanomolar concentrations.""")

    # Q3
    n += 1
    make_quiz_question_slide(prs, 3,
        "What did Hodgkin & Huxley model in their Nobel Prize-winning 1952 work?",
        ["The structure of DNA in nerve cells",
         "The mathematical equations describing action potential generation in squid axons",
         "The first computer simulation of the human brain",
         "The discovery of synaptic vesicles by electron microscopy"],
        n, notes="Historical context. Students may confuse with Watson & Crick (DNA, 1953).")

    n += 1
    make_quiz_answer_slide(prs, 3,
        "What did Hodgkin & Huxley model in their Nobel Prize-winning 1952 work?",
        ["The structure of DNA in nerve cells",
         "The mathematical equations describing action potential generation in squid axons",
         "The first computer simulation of the human brain",
         "The discovery of synaptic vesicles by electron microscopy"],
        1,
        "Correct: Hodgkin & Huxley (1952) wrote 4 coupled ODEs describing action potential generation using voltage-clamp data from squid giant axons.\n1963 Nobel Prize in Physiology or Medicine. First quantitative biological model of cellular excitability. Still the foundation of all computational neuroscience.",
        n,
        notes="""HISTORICAL CONTEXT:
Alan Hodgkin and Andrew Huxley (Cambridge, UK) asked: HOW does a nerve impulse work mathematically? They chose the squid giant axon (Loligo) because it is 1 mm in diameter — 1000x larger than a human axon — easy to impale with metal wire electrodes. They could also perfuse the inside with different salt solutions.

VOLTAGE CLAMP: They invented (with Cole & Marmont) the voltage clamp. A feedback amplifier holds the membrane voltage at a set value and measures the current needed to do so. Since V is constant, the current directly reflects channel conductance: I = g(V-E).

KEY EXPERIMENTS:
1. Replace Na+ outside with choline (impermeant) → the fast inward current disappears. Confirms it is Na+.
2. Replace K+ inside with Cs+ (blocks K+ channels) → the slow outward current disappears. Confirms it is K+.
3. Fit empirical equations to the kinetics of each conductance: m³h for Na+, n⁴ for K+.

THE MODEL: 5 papers published in J. Physiology in 1952. The last paper (Hodgkin & Huxley 1952d) gives the full quantitative model and shows it correctly predicts: (1) action potential shape and amplitude, (2) threshold phenomenon, (3) refractory period, (4) accommodation. All from 4 equations fit to voltage clamp data.

IMPACT: Every neuron model since 1952 uses HH-style conductance equations. The NEURON simulator (used by most computational neuroscientists) is essentially HH + cable theory. Understanding HH = understanding all modern computational neuroscience.""")

    # Q4
    n += 1
    make_quiz_question_slide(prs, 4,
        "What does the SpikerBox (Backyard Brains) record?",
        ["The electrical activity of individual genes during transcription",
         "Extracellular neural spikes (~100-500 µV) from nearby neurons or muscles",
         "Blood oxygen levels using near-infrared spectroscopy",
         "The force of muscle contractions using strain gauges"],
        n, notes="Tests SpikerBox understanding from the physiology intro slides.")

    n += 1
    make_quiz_answer_slide(prs, 4,
        "What does the SpikerBox (Backyard Brains) record?",
        ["The electrical activity of individual genes during transcription",
         "Extracellular neural spikes (~100-500 µV) from nearby neurons or muscles",
         "Blood oxygen levels using near-infrared spectroscopy",
         "The force of muscle contractions using strain gauges"],
        1,
        "Correct: Extracellular neural spikes — tiny voltage deflections (~100-500 µV) caused by action potentials in nearby neurons/muscles.\nA metal electrode near a neuron picks up the ionic current loop of the spike. Amplified ~1000x. Can be heard as 'pops' through a speaker. Drug experiment: add lidocaine → spikes disappear.",
        n,
        notes="""HOW EXTRACELLULAR RECORDING WORKS:
When a neuron fires, Na+ flows IN at the spike site. This creates a current loop: Na+ enters the cell, creating a 'current sink' — a region where positive charge is being absorbed from the extracellular space. The extracellular fluid has to supply this Na+, meaning current flows TOWARD the spike site from adjacent regions (source-to-sink).

An electrode placed near the neuron detects this current loop as a voltage fluctuation. The signal is tiny (~100-500 microvolts = 0.1-0.5 mV vs. 1.5 V in a AA battery = 3000-15000x smaller).

THE SPIKERBOX ELECTRONICS:
- Amplification: ~60 dB (x1000) to bring signal to audible/displayable range
- Bandpass filter: 300-3000 Hz. This passes spike-frequency content, rejects: (a) motion artifact (very low freq), (b) electrode noise (very high freq), (c) 50/60 Hz power line interference (notch filter)
- Output: audio jack (you literally HEAR the spikes as pops!) or USB to computer

DRUG EXPERIMENT (classic demonstration):
1. Set up cockroach leg preparation (remove one leg, keep alive with saline)
2. Record baseline spontaneous activity (see spikes every few seconds)
3. Add lidocaine (0.1-1 mg/mL): Na+ channels blocked → spikes slow, then stop
4. Wash with saline: spikes return (lidocaine diffuses away)
5. This IS a dose-response experiment — you could measure IC50!

CONNECTION TO QSAR: The measurement you just did (how much lidocaine blocks 50% of spikes?) gives an IC50 value. This is EXACTLY what QSAR models predict from molecular structure. The SpikerBox is the physiological validation of an AI prediction.""")

    # Q5
    n += 1
    make_quiz_question_slide(prs, 5,
        "What is the main function of GABA (gamma-aminobutyric acid) in the adult brain?",
        ["Main EXCITATORY neurotransmitter — promotes action potential firing",
         "Main INHIBITORY neurotransmitter — reduces neuronal firing via Cl- influx",
         "An oxygen carrier molecule in glial cells",
         "A hormone released from the pituitary that controls sleep"],
        n, notes="GABA vs. glutamate is fundamental. Tests NT knowledge from SMILES practice slides.")

    n += 1
    make_quiz_answer_slide(prs, 5,
        "What is the main function of GABA (gamma-aminobutyric acid) in the adult brain?",
        ["Main EXCITATORY neurotransmitter — promotes action potential firing",
         "Main INHIBITORY neurotransmitter — reduces neuronal firing via Cl- influx",
         "An oxygen carrier molecule in glial cells",
         "A hormone released from the pituitary that controls sleep"],
        1,
        "Correct: GABA is the main INHIBITORY neurotransmitter. GABA-A receptor = Cl- channel.\nGABA binds → channel opens → Cl- flows in → membrane hyperpolarizes → less likely to fire.\nDrugs: diazepam (Valium), alcohol, propofol all potentiate GABA-A. Same Cys-loop superfamily as GluCl (ivermectin target in flies!).",
        n,
        notes="""GABA-A RECEPTOR STRUCTURE AND MECHANISM:
GABA-A is a pentameric (5-subunit) ligand-gated ion channel. The 5 subunits (2α + 2β + 1γ most commonly) surround a central pore. The pore is selective for Cl-. Two GABA molecules bind simultaneously (one at each α-β interface). Binding → conformational change → pore opens → Cl- flows in.

WHY DOES Cl- ENTRY INHIBIT THE NEURON?
Cl- has a reversal potential (E_Cl) near or more negative than the resting potential (~-70 to -75 mV). When GABA-A opens:
- If V > E_Cl: Cl- flows IN (makes inside more negative = hyperpolarization)
- If V = E_Cl: no net Cl- flow (but the conductance increase still "short-circuits" excitatory currents — a mechanism called "shunting inhibition")
Net result: harder to depolarize the neuron to threshold.

DRUG BINDING SITES ON GABA-A:
1. GABA binding site (α-β interface): endogenous ligand
2. Benzodiazepine site (α-γ interface): diazepam, lorazepam, alprazolam. POTENTIATE GABA-A: increases frequency of channel opening when GABA is present. Cannot open channel without GABA → reason BZDs are safer than barbiturates.
3. Barbiturate site (β subunit transmembrane domain): phenobarbital. At high concentrations, opens channel WITHOUT GABA. Explains narrow therapeutic window and high overdose risk.
4. Neurosteroid site (transmembrane): allopregnanolone, alphaxalone, propofol.
5. Alcohol site: poorly defined, possibly multiple sites.

CYS-LOOP CONNECTION TO IVERMECTIN:
The Cys-loop superfamily includes GABA-A, glycine receptor (GlyR), nAChR, 5-HT3, AND GluCl (invertebrate glutamate-gated Cl- channel). Ivermectin opens GluCl in worms/insects → sustained Cl- influx → paralysis. GluCl doesn't exist in vertebrates (we have GlyR instead, which ivermectin also modulates but less potently). This pharmacological specificity is why ivermectin is safe for humans at antiparasitic doses.""")

    # Q6
    n += 1
    make_quiz_question_slide(prs, 6,
        "In fly (Drosophila) motion vision, what is the role of T4 neurons?",
        ["Detect overall brightness levels in the visual field",
         "Detect ON-motion (a brightening edge moving in one of 4 directions)",
         "Process color information from UV-sensitive photoreceptors",
         "Send feedback signals from the brain back to the eye"],
        n, notes="Tests fly motion vision knowledge from the physiology intro.")

    n += 1
    make_quiz_answer_slide(prs, 6,
        "In fly (Drosophila) motion vision, what is the role of T4 neurons?",
        ["Detect overall brightness levels in the visual field",
         "Detect ON-motion (a brightening edge moving in one of 4 directions)",
         "Process color information from UV-sensitive photoreceptors",
         "Send feedback signals from the brain back to the eye"],
        1,
        "Correct: T4 neurons are elementary motion detectors for ON-motion (brightening edges).\n4 subtypes: front-to-back, back-to-front, upward, downward. T5 neurons do the same for OFF-motion (darkening edges). Together they feed into LPTCs for whole-field optic flow. Requires GluCl-mediated inhibition (Ammer, Serbe-Kamp et al. 2023, Nat. Neurosci.).",
        n,
        notes="""THE T4/T5 CIRCUIT IN DETAIL:
Photoreceptors → Lamina (L1-L5) → Medulla → T4/T5 → Lobula Plate Tangential Cells (LPTCs)

WHAT IS AN ELEMENTARY MOTION DETECTOR?
An EMD detects motion in one direction. The classic Reichardt detector model: take two adjacent photoreceptor inputs, delay one, multiply (correlate) the two signals. If the bright edge moves from left to right, it hits receptor A first, then B slightly later. If you delay signal A and then multiply A×B, the multiplication is large only when both signals are coincident — which happens when the edge is moving left-to-right. Moving right-to-left gives a small or negative output.

T4 neurons implement something like this Reichardt detector using circuits with excitatory (AMPA-type glutamate) and inhibitory (GluCl) inputs from specific medulla neurons.

GluCl ROLE (Ammer, Serbe-Kamp et al. 2023):
T5 neurons receive direction-opponent inhibition through GluCl channels. This inhibition is crucial for DIRECTION SELECTIVITY — without it, T5 responds to motion in all directions equally (loses selectivity). The specific circuit: Mi9 (inhibitory medulla neuron) → T5 via GluCl. Mi9 responds to motion in the NULL direction. By providing inhibitory input through GluCl when the null direction is present, T5's response to null direction is suppressed → direction selectivity.

FOR DRUG DISCOVERY:
This circuit shows how a specific ion channel (GluCl) at a specific synapse (Mi9→T5) implements a specific computation (direction selectivity). Blocking GluCl with ivermectin eliminates this computation. This is the same logic as: blocking Nav with lidocaine eliminates pain signaling. Understanding which computation depends on which ion channel helps predict drug effects on behavior.""")

    # Q7
    n += 1
    make_quiz_question_slide(prs, 7,
        "Why is electron microscopy (EM) essential for connectomics (mapping neural circuits)?",
        ["EM is faster than light microscopy for large tissue volumes",
         "EM works on living tissue while light microscopy requires fixation",
         "Synapses (20-40 nm) are below the ~200 nm resolution limit of light",
         "EM can image fluorescent proteins that light microscopes cannot detect"],
        n, notes="Tests understanding of EM from the morning physiology slides.")

    n += 1
    make_quiz_answer_slide(prs, 7,
        "Why is electron microscopy (EM) essential for connectomics (mapping neural circuits)?",
        ["EM is faster than light microscopy for large tissue volumes",
         "EM works on living tissue while light microscopy requires fixation",
         "Synapses (20-40 nm) are below the ~200 nm resolution limit of light",
         "EM can image fluorescent proteins that light microscopes cannot detect"],
        2,
        "Correct: Synapses are ~20-40 nm wide — far below the diffraction limit of visible light (~200 nm).\nElectrons have wavelength ~0.001 nm → 200x better resolution. EM can resolve individual synaptic vesicles (40 nm), cell membranes (5 nm thick), and postsynaptic densities. Essential for identifying every synapse in a neural circuit.",
        n,
        notes="""THE PHYSICS BEHIND THE RESOLUTION LIMIT:
Abbe diffraction limit: minimum resolvable feature ≈ λ / (2 × NA)
- Visible light: λ = 400-700 nm, NA ≤ 1.4 (oil immersion) → resolution ~140-250 nm
- Electrons at 80 kV: de Broglie wavelength ≈ 0.004 nm → theoretical resolution ~0.002 nm, practical ~0.1-5 nm for biological specimens

WHY DO WE NEED TO SEE SYNAPSES?
A synapse is the connection between two neurons. It has:
- Presynaptic terminal: ~1 µm in diameter, contains 100s of synaptic vesicles (40 nm each, filled with NT)
- Synaptic cleft: 20 nm gap
- Postsynaptic density: a protein scaffold holding receptors, ~30 nm thick

With light microscopy you can see a synaptic bouton (the ~1 µm terminal) but you CANNOT see: whether a synapse is present (vs. just axon passing by), the type of synapse (excitatory vs. inhibitory — different ultrastructure), or the number of vesicles (which correlates with synaptic strength).

With EM you see all of this.

CONNECTOMICS WORKFLOW:
1. Fix tissue in glutaraldehyde (cross-links proteins) + osmium tetroxide (stains membranes with heavy metal → electron-dense = dark)
2. Embed in epoxy resin (hard plastic)
3. Ultramicrotome slices tissue into 30-50 nm sections (thinner than most bacteria!)
4. Image each section: scanning EM (SEM) in backscatter mode, or transmission EM (TEM). Each section = one 2D image, ~terabyte per mm³.
5. AI (CNNs for semantic segmentation) identifies cell membrane boundaries in each slice
6. 3D reconstruction: stack slices → connect membranes → 3D neuron shapes
7. Synapse detection: look for presynaptic specialization (T-bar in fly, active zone density) opposite postsynaptic density
8. Build graph: nodes=neurons, edges=synapses with weights=synapse count

AI'S ROLE: Steps 5-7 are automated by AI. The FlyWire connectome (2024) used deep learning (flood-filling networks) to segment 139,000 neurons. Without AI, this would take a human lifetime.""")

    # Q8
    n += 1
    make_quiz_question_slide(prs, 8,
        "Why do opioids (morphine, fentanyl) cause pinpoint pupils (miosis)?",
        ["Opioids block sympathetic pupil dilation receptors",
         "Opioids directly constrict the iris sphincter muscle",
         "Opioids activate the Edinger-Westphal nucleus (parasympathetic) → pupil constriction",
         "Opioids reduce retinal photoreceptor sensitivity"],
        n, notes="Connects eye-tracking pharmacology to the opioid mechanism. Classic clinical sign used by police and paramedics.")

    n += 1
    make_quiz_answer_slide(prs, 8,
        "Why do opioids (morphine, fentanyl) cause pinpoint pupils (miosis)?",
        ["Opioids block sympathetic pupil dilation receptors",
         "Opioids directly constrict the iris sphincter muscle",
         "Opioids activate the Edinger-Westphal nucleus (parasympathetic) → pupil constriction",
         "Opioids reduce retinal photoreceptor sensitivity"],
        2,
        "Correct: Opioids activate µ-opioid receptors on the Edinger-Westphal (EW) nucleus.\nEW → ciliary ganglion → sphincter pupillae → miosis (small pupils). Tolerance does NOT develop to miosis (unlike analgesia, euphoria). Pinpoint pupils + respiratory depression = opioid overdose triad.",
        n,
        notes="""AUTONOMIC CONTROL OF THE PUPIL:
The pupil is controlled by two antagonistic muscles:
1. SPHINCTER PUPILLAE (constrictor): ring of smooth muscle. Parasympathetically controlled. CN III → Edinger-Westphal nucleus → ciliary ganglion → sphincter → MIOSIS (small pupil)
2. DILATOR PUPILLAE: radial smooth muscle. Sympathetically controlled. Hypothalamus → cervical sympathetic → superior cervical ganglion → dilator → MYDRIASIS (large pupil)

OPIOID MECHANISM:
µ-opioid receptors (MOR) are GPCRs (Gi-coupled). In the Edinger-Westphal nucleus, opioids activate MOR → decrease cAMP → hyperpolarise neurons → increases their output. Wait — but opioids typically INHIBIT neurons? 

Here's the complexity: MOR activation in EW neurons DISINHIBITS the EW output. The EW has GABA interneurons that inhibit the EW projection neurons. Opioids suppress these interneurons → EW projection neurons become MORE active → more parasympathetic tone → more sphincter contraction → miosis.

WHY DOESN'T TOLERANCE DEVELOP TO MIOSIS?
Good question! Tolerance to opioid effects (analgesia, euphoria, respiratory depression) develops quickly — probably due to receptor desensitization and downregulation. But EW miosis is mediated through a different downstream signaling pathway that apparently doesn't desensitize as quickly. This means miosis is a reliable indicator of opioid PRESENCE even in chronic users.

CLINICAL USE:
- Paramedics: pinpoint pupils + unconscious + slow breathing = opioid overdose → give naloxone (opioid antagonist)
- Police: pinpoint pupils can indicate current opioid use
- Anesthesiologists: monitor pupil size during surgery to gauge depth of anesthesia and opioid dosing

AI + PUPILLOMETRY FOR DRUG DISCOVERY:
Tracking pupil diameter over time (pupillary light reflex or spontaneous fluctuations) creates a time-series that encodes autonomic nervous system state. AI can classify: (1) current drug state (opioid vs. stimulant vs. sober), (2) dosage estimation, (3) predict side effects. This is a non-invasive, continuous physiological measure — ideal for AI classification models.""")

    # Q9
    n += 1
    make_quiz_question_slide(prs, 9,
        "Which is the correct order of events at a chemical synapse?",
        ["NT release → Ca2+ entry → receptor activation → vesicle fusion",
         "AP arrives → Ca2+ enters (Cav) → vesicle fuses → NT released → receptor binds",
         "Receptor activation → AP generated → Ca2+ exits → vesicle forms",
         "Ca2+ exits → vesicle fuses → NT uptake → receptor blocked"],
        n, notes="Tests understanding of synaptic transmission sequence. Many drug targets are at specific steps of this process.")

    n += 1
    make_quiz_answer_slide(prs, 9,
        "Which is the correct order of events at a chemical synapse?",
        ["NT release → Ca2+ entry → receptor activation → vesicle fusion",
         "AP arrives → Ca2+ enters (Cav) → vesicle fuses → NT released → receptor binds",
         "Receptor activation → AP generated → Ca2+ exits → vesicle forms",
         "Ca2+ exits → vesicle fuses → NT uptake → receptor blocked"],
        1,
        "Correct: AP arrives at terminal → Cav channels open → Ca2+ in → synaptotagmin senses Ca2+ → SNARE proteins fuse vesicle → NT released → diffuses across 20 nm cleft → binds postsynaptic receptors.\nBotox cleaves SNARE proteins = blocks step 3. SSRIs block NT reuptake after step 5.",
        n,
        notes="""SYNAPTIC TRANSMISSION STEP BY STEP:

1. ACTION POTENTIAL ARRIVES at axon terminal (bouton):
The AP propagates down the axon (saltatory in myelinated axons). The bouton membrane depolarizes.

2. VOLTAGE-GATED Ca2+ CHANNELS (Cav) OPEN:
Cav2.1 (P/Q-type) and Cav2.2 (N-type) are the main subtypes at fast synapses. Ca2+ is 1-2 mM outside, ~100 nM inside → 10,000-20,000x gradient! AND electrical gradient (negative inside) also pulls Ca2+ in. Ca2+ flows rapidly. Near-channel concentration spikes to ~100 µM transiently.

3. Ca2+ BINDS SYNAPTOTAGMIN:
Synaptotagmin is a Ca2+ sensor protein on synaptic vesicle membranes. It has two C2 domains that bind Ca2+. When Ca2+ binds, synaptotagmin changes conformation and interacts with SNARE proteins.

4. SNARE COMPLEX ZIPS UP → VESICLE FUSES:
SNARE proteins: VAMP/synaptobrevin (on vesicle), syntaxin (on plasma membrane), SNAP-25 (on plasma membrane). The SNARE complex forms a 4-helix bundle that zips from the N-terminus toward the C-terminus, pulling the vesicle membrane and plasma membrane together until they fuse. NT is released into the cleft.

5. NT DIFFUSES AND BINDS RECEPTORS:
The cleft is only 20 nm — diffusion takes microseconds. NT binds postsynaptic ionotropic (ion channel) or metabotropic (GPCR) receptors.

6. NT CLEARANCE:
(a) Reuptake transporters: SERT (serotonin), DAT (dopamine), NET (norepinephrine), GAT1 (GABA). Bring NT back into presynaptic terminal. These are the main SSRI/SNRI/cocaine targets.
(b) Enzymatic degradation: AChE cleaves ACh → choline + acetate. MAO-A/B degrade dopamine, serotonin. These are drug targets too (rivastigmine for Alzheimer's, selegiline for Parkinson's/depression).
(c) Diffusion away from cleft (spillover to extrasynaptic receptors).

DRUG TARGETS AT EACH STEP:
- Step 2 (Cav): Ziconotide (Prialt) = Cav N-type blocker, for severe pain. Ca2+ channel blockers (verapamil, amlodipine) = cardiac.
- Step 4 (SNARE): Botulinum toxin A: cleaves SNAP-25. Botulinum toxin B: cleaves VAMP. Both prevent vesicle fusion → no NT release → paralysis. Used medically (spasticity, dystonia, cosmetic wrinkles).
- Step 5 (receptors): vast target space (see all of drug discovery!)
- Step 6 (reuptake): SSRIs (fluoxetine, sertraline, escitalopram), SNRIs (venlafaxine), cocaine, amphetamine, MDMA.""")

    # Q10
    n += 1
    make_quiz_question_slide(prs, 10,
        "What is the main benefit of the Computational SpikerBox for drug discovery?",
        ["It costs millions and is only available at top research institutes",
         "It records from human brain organoids for direct clinical relevance",
         "It simulates Hodgkin-Huxley in software — test drug effects on ion channels in silico first",
         "It replaces ChEMBL as the source of molecular activity data"],
        n, notes="Tests understanding of the in silico → in vitro → in vivo pipeline and the role of simulation.")

    n += 1
    make_quiz_answer_slide(prs, 10,
        "What is the main benefit of the Computational SpikerBox for drug discovery?",
        ["It costs millions and is only available at top research institutes",
         "It records from human brain organoids for direct clinical relevance",
         "It simulates Hodgkin-Huxley in software — test drug effects on ion channels in silico first",
         "It replaces ChEMBL as the source of molecular activity data"],
        2,
        "Correct: Free HH simulator — instantly test what happens if Nav is 80% blocked, or Kv is prolonged.\nThe 3i pipeline: AI predicts IC50 (in silico) → HH simulation predicts physiological consequence (in silico) → SpikerBox validates with real neural tissue (in vitro) → then animal models (in vivo). Each step filters candidates, reducing cost.",
        n,
        notes="""THE 3i PIPELINE:

IN SILICO (computer simulation):
Step 1: QSAR/GNN model predicts: 'compound X has predicted IC50 = 50 nM for Nav1.2'
Step 2: Computational SpikerBox / HH simulation: 'If I set g_Na to 20% of baseline (80% block), what happens to firing?'
- Maybe the neuron stops firing completely → potential anesthetic/anticonvulsant
- Maybe it only reduces firing frequency → less side effects
- Maybe nothing changes → channel not important at this neuron's firing frequency
- Maybe the neuron fires MORE (paradoxical, but happens with some partial blockers)
Cost: essentially zero. Can test 1000 compounds in minutes.

IN VITRO (glass = laboratory, not whole animal):
After filtering with in-silico, take the 10-50 best candidates.
Test on: cockroach SpikerBox, cultured neurons, patch-clamp electrophysiology.
Measure: real IC50 for neural firing. Compare to QSAR prediction.
Cost: ~$100-10,000 per compound depending on assay.

IN VIVO (living animal):
Only the 2-5 best in-vitro candidates go to animal models.
Efficacy: does the drug reduce seizures / pain / arrhythmia in mice/rats?
Safety: does it cause cardiac arrhythmia, sedation, weight change?
Cost: ~$50,000-500,000 per compound.

CLINICAL TRIALS:
Only 1 in 5000 in-silico candidates typically makes it to clinical trials.
But with better AI + better in-silico filtering, this ratio is improving.
Insilico Medicine used AI to design a new drug (ISM001-055 for IPF) in 18 months (vs. typical 5-6 years).

THE COMPUTATIONAL SPIKERBOX'S ROLE:
It's the bridge between 'we predict this blocks Nav' (QSAR) and 'we validated this blocks neural firing' (SpikerBox). By letting you run the HH model interactively (drag a slider for g_Na from 0% to 100%), you build intuition about which ion channel contributions matter at which voltage and firing frequency. This intuition improves your interpretation of QSAR predictions.""")

    # Closing
    n += 1
    make_section_divider(prs, "Neuroscience Unlocked! Ready for Drug Discovery AI? 🧪⚡", n,
        notes="End of Quiz 3 (Physiology). Students should now have a solid working vocabulary of action potential physiology, HH model, SpikerBox, neurotransmitters, fly vision, EM, eye-tracking pharmacology, and synaptic transmission. Transition to the HH explainer deck, or directly to the molecular ML content.",
        subtitle="Action potentials, HH, GABA, synaptic transmission, SpikerBox, T4/T5, EM, opioid pupils — check!")

    os.makedirs("quiz_physiology", exist_ok=True)
    prs.save("quiz_physiology/slides.pptx")
    print(f"  Quiz 3 (Physiology): {n} slides -> quiz_physiology/slides.pptx")



# =====================================================================
# HODGKIN-HUXLEY EXPLAINER DECK
# Plain-language walk through of the HH equations, gating variables,
# channel extensions, and drug targets
# =====================================================================
def generate_hh_explainer():
    prs = new_prs()
    n = 0

    n += 1
    make_title_slide(prs,
        "The Hodgkin-Huxley Model — A Complete Guide",
        "Understanding the equations of the action potential, channel extensions, and drug targets",
        1, "AI for Drug Discovery", date_str="22 April 2026",
        notes="""WHO IS THIS DECK FOR?
This deck is for instructors or students who want to understand the Hodgkin-Huxley model deeply. It assumes basic calculus (what a derivative is) and basic chemistry (ions, concentrations, gradients). No prior neuroscience required.

WHAT IS COVERED:
1. Biological context: the neuron as an electrical circuit
2. The 4 HH ODEs — each explained line by line
3. Gating variables m, h, n — physical meaning and kinetics
4. How voltage-clamp experiments led to the model
5. Channel extensions: A-current (Kv4), Cav (Ca2+ channels), Ih (HCN), persistent Nav
6. Synaptic conductances: AMPA, NMDA, GABA-A
7. Multicompartment models and cable theory
8. Drug targets at every level
9. Python code snippets (numpy / scipy ODE solver)
10. Connection to QSAR: why HH helps interpret predictions

HOW TO USE:
Work through this deck alongside the Day 2 physiology intro. The presenter notes contain the full explanation of every concept. The slides show the key equations and terms.""")

    # --- 1. The neuron as a circuit ---
    n += 1
    make_content_slide(prs,
        "The Neuron as an Electrical Circuit",
        ["The HH model treats a patch of membrane as an electrical circuit:",
         "  Membrane (lipid bilayer) = CAPACITOR (stores charge)",
         "  Ion channels (Nav, Kv, leak) = RESISTORS in parallel",
         "  Ion pumps (Na+/K+ ATPase) = BATTERIES (maintain gradients)",
         "  Cytoplasm = conducting WIRE along the axon",
         "",
         "Key parameters:",
         "  Cm = membrane capacitance = 1 µF/cm² (universal for all neurons)",
         "  g_Na, g_K, g_L = maximum conductances (how many open channels)",
         "  E_Na, E_K, E_L = reversal potentials (the 'battery voltages')",
         "",
         "Kirchhoff's current law: total current in = total current out",
         "  Cm × dV/dt = external current - sum of ionic currents"],
        n,
        notes="""THE MEMBRANE CAPACITOR:
A capacitor stores charge by separating two conductors with an insulator. The cell membrane is exactly this: conducting cytoplasm | insulating lipid bilayer (5 nm) | conducting extracellular fluid. Membrane capacitance = 1 µF/cm². This is a universal constant — all biological membranes have ~1 µF/cm² regardless of cell type or species.

What does Cm = 1 µF/cm² mean physically? Q = Cm × V. If V changes by 100 mV (as in an action potential), the charge that moved is: ΔQ = 1 µF/cm² × 0.1 V = 0.1 µC/cm² = 1e-7 C/cm². Since one Na+ ion carries 1.6e-19 C, this is 0.1e-6 / 1.6e-19 = 6.25e11 ions/cm² = ~625,000 ions/µm². Sounds like a lot, but the bulk Na+ concentration barely changes (the gradient is maintained).

THE ION CHANNEL CONDUCTANCES:
Conductance g (unit: Siemens = A/V = 1/Ohm). Current through a channel type: I = g × (V - E_reversal). E_reversal is the reversal potential (the voltage where no net current flows despite the channel being open). For Na+: E_Na ≈ +55 mV. For K+: E_K ≈ -90 mV. If V = -70 mV and Na+ channels open: driving force = V - E_Na = -70 - 55 = -125 mV. Current = g_Na × (-125 mV) = large inward current.

THE NERNST POTENTIAL (battery voltage):
E_ion = (RT/zF) × ln([ion]_out / [ion]_in)
At 37°C: RT/F = 26.7 mV
E_Na = 26.7 × ln(145/12) = 26.7 × 2.5 = +66 mV (approximated as +55 mV in HH)
E_K = 26.7 × ln(4/155) = 26.7 × (-3.65) = -97 mV (approximated as -77 mV in HH)
E_L = -54.4 mV (empirical, includes all non-Nav/Kv channels)

KIRCHHOFF'S CURRENT LAW:
The sum of all currents at a node = 0. Currents flowing into the membrane patch:
- Injected current I_ext (from electrode or previous compartment)
- Capacitive current I_cap = Cm × dV/dt (charges/discharges the capacitor)
- Ionic currents (Na+, K+, leak, etc.)
Setting sum = 0: Cm × dV/dt + I_Na + I_K + I_L = I_ext
→ Cm × dV/dt = I_ext - I_Na - I_K - I_L""")

    # --- 2. The 4 HH equations ---
    n += 1
    make_content_slide(prs,
        "The Four Hodgkin-Huxley Equations",
        ["Equation 1 — Membrane voltage V (in mV):",
         "  Cm × dV/dt = I_ext - g_Na×m³h×(V-E_Na) - g_K×n⁴×(V-E_K) - g_L×(V-E_L)",
         "",
         "Equation 2 — Na+ activation gate m (0 ≤ m ≤ 1):",
         "  dm/dt = αm(V)×(1-m) - βm(V)×m",
         "",
         "Equation 3 — Na+ inactivation gate h (0 ≤ h ≤ 1):",
         "  dh/dt = αh(V)×(1-h) - βh(V)×h",
         "",
         "Equation 4 — K+ activation gate n (0 ≤ n ≤ 1):",
         "  dn/dt = αn(V)×(1-n) - βn(V)×n",
         "",
         "Parameters: g_Na=120, g_K=36, g_L=0.3 mS/cm² | Cm=1 µF/cm²",
         "E_Na=+55, E_K=-77, E_L=-54.4 mV | V_rest=-65 mV"],
        n,
        notes="""EQUATION 1 — MEMBRANE VOLTAGE:
Cm × dV/dt = I_ext - g_Na×m³h×(V-E_Na) - g_K×n⁴×(V-E_K) - g_L×(V-E_L)

Breaking it down:
- Cm × dV/dt: rate of change of voltage × capacitance = the capacitive current. This is how fast the membrane is charging up or discharging.
- I_ext: injected current (from an electrode, or from a neighboring compartment in a cable model). This is the "input" that drives the neuron.
- g_Na×m³h: the instantaneous Na+ conductance. m³ means THREE activation gates must be open. h is the inactivation gate (must also be open = not inactivated). g_Na = 120 mS/cm² is the maximum possible Na+ conductance (when all channels are open).
- g_K×n⁴: instantaneous K+ conductance. FOUR activation gates (n⁴). g_K = 36 mS/cm².
- g_L = 0.3 mS/cm² is the leak conductance (constant — it doesn't change with voltage).

Why m³ and n⁴?
Hodgkin & Huxley fit the kinetics empirically. The Na+ activation data fitted best with a third-order process (three independent gates). We now know why: the Nav channel has 4 domains, each with a voltage sensor (S4 helix), but not all 4 need to move simultaneously. The effective kinetics appear as m³. The K+ channel (Kv) has 4 identical subunits, each contributing one gate → n⁴.

EQUATION 2 — Na+ ACTIVATION GATE (m):
dm/dt = αm(V)×(1-m) - βm(V)×m

Physical meaning:
- m = fraction of activation gates that are in the OPEN state (ranges 0 to 1)
- (1-m) = fraction that are CLOSED
- αm(V) = rate constant for OPENING (per ms), depends on voltage
- βm(V) = rate constant for CLOSING (per ms), depends on voltage
- dm/dt = (rate of opening × currently closed) - (rate of closing × currently open)

At steady state (dm/dt = 0):
m∞(V) = αm / (αm + βm)   [steady-state open probability]
τm(V) = 1 / (αm + βm)     [time constant of equilibration]

Empirical functions (Hodgkin & Huxley 1952, adjusted for V shifted by -65 mV):
αm(V) = 0.1 × (V+40) / (1 - exp(-(V+40)/10))  [if V ≠ -40 mV]
βm(V) = 4 × exp(-(V+65)/18)
At -65 mV: m∞ ≈ 0.05 (5% of channels activated), τm ≈ 0.4 ms
At +20 mV: m∞ ≈ 0.99 (99% activated), τm ≈ 0.1 ms  ← very fast!

EQUATION 3 — Na+ INACTIVATION GATE (h):
Same structure, BUT:
- h = fraction of inactivation gates in the NON-INACTIVATED (= permissive) state
- h DECREASES when the cell is depolarized (inactivation gates CLOSE at depolarized voltages)
- At rest (-65 mV): h∞ ≈ 0.60 (most channels not inactivated)
- At +20 mV: h∞ ≈ 0.02 (almost all inactivated)  ← slow process, τh ≈ 8 ms
The refractory period: after the spike, h is near 0 (inactivated). Even though m wants to be near 1, h × m³ ≈ 0 → no Na+ current → neuron cannot fire again until h recovers.

αh(V) = 0.07 × exp(-(V+65)/20)
βh(V) = 1 / (exp(-(V+35)/10) + 1)

EQUATION 4 — K+ ACTIVATION GATE (n):
Same structure as m, but:
- n opens more slowly than m (τn ≈ 1-5 ms vs τm ≈ 0.1-0.5 ms)
- n⁴ means ALL FOUR gates must be open for K+ current
- n∞(V): sigmoidal, half-activated at about -15 mV
- At rest: n∞ ≈ 0.32 → n⁴ ≈ 0.01 (very little K+ current at rest)
- After spike peak (+40 mV): n∞ ≈ 0.75 → n⁴ ≈ 0.32 (repolarizing K+ current)

αn(V) = 0.01 × (V+55) / (1 - exp(-(V+55)/10))
βn(V) = 0.125 × exp(-(V+65)/80)

COUPLING: All 4 equations are coupled. V drives m, h, n. m, h, n determine the conductances that drive V. This is a nonlinear dynamical system — its behavior (single spikes, burst firing, oscillations) depends on the parameters.""")

    # --- 3. Gating variables explained ---
    n += 1
    make_content_slide(prs,
        "Gating Variables m, h, n — Physical Meaning",
        ["Each gating variable represents a protein conformation probability:",
         "",
         "  m = Na+ channel activation gate (OPEN = 1)",
         "     Fast opening on depolarization (τm ~ 0.1-0.5 ms)",
         "     m∞(V): sigmoidal, half-activated at ~-40 mV",
         "",
         "  h = Na+ channel inactivation gate (PERMISSIVE = 1, INACTIVATED = 0)",
         "     Slow closing on depolarization (τh ~ 5-10 ms at spike peak)",
         "     h∞(V): sigmoidal but DECREASING — starts near 1 at rest",
         "     Causes refractory period: h → 0 after spike → must recover",
         "",
         "  n = K+ channel activation gate (OPEN = 1)",
         "     Intermediate speed (τn ~ 1-5 ms)",
         "     n∞(V): sigmoidal, half-activated at ~-15 mV",
         "",
         "All three: two-state Markov chain. α = opening rate, β = closing rate",
         "Steady-state: x∞ = α/(α+β)  |  Time constant: τx = 1/(α+β)"],
        n,
        notes="""PHYSICAL INTERPRETATION OF m, h, n:

These are not just mathematical variables — they have direct physical meaning as probabilities of conformational states of individual protein gates.

THE m GATE (Na+ activation):
Real structure: each Nav channel has 4 voltage-sensing domains (I-IV). Each has a positively-charged S4 helix that moves outward when the membrane depolarizes (positive charges move toward the outside = toward the positive side). When ~3 of the 4 S4 helices have moved, the channel activation gate opens (there is actually a physical gate — the S6 helices form the bottom of the pore, they move apart to open).

m³ in HH: three independent gates each with probability m. For the channel to be open, all three must be open → probability m×m×m = m³.

In reality there are 4 subunits in Nav (each contributing a voltage sensor), but the S4 of domain IV is slower and couples to the inactivation gate, so the effective kinetics is m³ not m⁴. (More modern models use m³ or sometimes different formulations.)

THE h GATE (Na+ inactivation):
Physical structure: a cytoplasmic loop (the "ball-and-chain" or "hinged lid") that physically blocks the open pore from the inside. At rest (-70 mV), the ball is floating freely and the pore is open to current (assuming m is also open). After depolarization (during the spike), the ball moves into the pore and plugs it — fast inactivation (the "ball" mechanism).

Slower inactivation (C-type, P/C) involves the outer mouth of the pore collapsing.

h = 1 means the inactivation gate is NOT blocking the pore (permissive). h = 0 means the pore is blocked (inactivated). So h = 1 at rest → h = 0 during sustained depolarization.

REFRACTORY PERIOD MECHANISM (important!):
During the rising phase: m shoots to ~1 (fast), h is still ~0.6 (slow to change)
At the peak: m starts to decrease (Na+ channels inactivate — now h falls rapidly)
During falling phase: K+ n⁴ is large (K+ flows out), m is falling, h is near 0
After repolarization: h slowly recovers (τh at -70 mV ≈ 5-10 ms)
During recovery: the cell cannot fire (absolute refractory) or needs larger stimulus (relative refractory)

This refractory period limits maximum firing rate to ~100-500 Hz in most neurons.

THE n GATE (K+ activation):
Kv channels (voltage-gated K+) have 4 identical subunits, each with a voltage sensor (S4 helix). n represents the probability that ONE subunit's voltage sensor is in the activated (open) conformation. All 4 must be activated → n⁴.

n opens more slowly than m: at the same voltage step, n lags behind m by ~1-2 ms. This delay means K+ current (n⁴) rises while the Na+ current (m³h) is already declining. The K+ current then repolarizes the membrane.

DRUG TARGETING OF GATING:
Local anesthetics (lidocaine, bupivacaine): bind inside the Nav pore when the channel is OPEN or INACTIVATED. They preferentially bind the inactivated state → stabilise h=0 → slower recovery from inactivation → use-dependent block (more block at higher firing rates). This is why they preferentially block rapidly-firing pain neurons.
TEA (tetraethylammonium): blocks Kv channels (blocks n⁴ channels). Used in research to isolate Na+ currents by blocking K+ currents. Not used clinically (too non-selective).
4-AP (4-aminopyridine): also Kv blocker. Clinically approved for multiple sclerosis (Ampyra) — improves walking by increasing AP duration in demyelinated axons.""")

    # --- 4. The action potential walkthrough ---
    n += 1
    make_content_slide(prs,
        "The Action Potential Step-by-Step (HH Simulation)",
        ["t=0 ms: V = -65 mV, m=0.05, h=0.60, n=0.32",
         "  → m³h = 0.05³×0.60 = 0.0001 (tiny Na+ current)",
         "  → n⁴ = 0.32⁴ = 0.01 (small K+ current)",
         "  → Resting: balanced currents",
         "",
         "t=0.1 ms: I_ext steps up → V depolarises to -55 mV",
         "  → m∞(-55) ≈ 0.20 → m starts rising (τm ~ 0.3 ms)",
         "",
         "t=0.5 ms: m ≈ 0.8, V ≈ +20 mV (SPIKE PEAK)",
         "  → m³ = 0.51, h still ≈ 0.50 → I_Na huge (inward)",
         "  → n still rising slowly → I_K not yet large",
         "",
         "t=1.0 ms: h → 0.05 (inactivated), n → 0.70",
         "  → Na+ current collapses, K+ current dominates → REPOLARIZATION",
         "",
         "t=2.0 ms: V undershoots to -80 mV (after-hyperpolarisation, AHP)",
         "  → n still elevated (slow to close) → extra K+ out",
         "",
         "t=5+ ms: h recovers → n closes → V returns to -65 mV"],
        n,
        notes="""HH SIMULATION WALKTHROUGH — NUMERICAL EXAMPLE:

This slide traces a single action potential through the HH equations numerically. Knowing the sequence helps you debug a simulation and understand drug effects.

RESTING STATE (t=0):
V = -65 mV (HH original uses this; some textbooks use -70 mV — just a shift convention)
m = m∞(-65) ≈ 0.05: only 5% of Na+ activation gates open
h = h∞(-65) ≈ 0.60: 60% of Na+ channels not inactivated
n = n∞(-65) ≈ 0.32: 32% of K+ gates open

Na+ conductance: g_Na × m³h = 120 × (0.05)³ × 0.60 = 120 × 0.000125 × 0.60 = 0.009 mS/cm²
K+ conductance: g_K × n⁴ = 36 × (0.32)⁴ = 36 × 0.0105 = 0.38 mS/cm²
Leak: g_L = 0.3 mS/cm² (constant)

At rest: total current = 0 (V is at a fixed point of the dynamical system). The numbers above are the steady-state solution.

THE THRESHOLD (t ≈ 0.1 ms):
When V reaches about -55 mV (the threshold), m∞(-55) ≈ 0.20. The m gate begins opening rapidly (τm ≈ 0.3 ms). If the driving current is large enough to bring V to -55 mV faster than m can open, threshold is crossed.

Positive feedback loop (regenerative): V rises → m opens → more Na+ in → V rises more → m opens more... This is the all-or-nothing principle. Once past threshold, the spike fires to completion regardless.

THE SPIKE PEAK (t ≈ 0.5-0.8 ms):
V ≈ +30 to +40 mV (close to E_Na = +55 mV but not reaching it because inactivation and K+ current start kicking in)
m ≈ 0.90 (almost fully activated)
h ≈ 0.45 (starting to fall — slow τh at these voltages ≈ 5 ms)
n ≈ 0.50 (rising — τn at +30 mV ≈ 1-2 ms)

Na+ current: 120 × (0.90)³ × 0.45 = 120 × 0.73 × 0.45 = 39 mS/cm² × (V-E_Na) = 39 × (30-55) = -975 µA/cm² (huge inward current)

THE FALL (t ≈ 1-2 ms):
h → 0.05 (most channels inactivated)
n → 0.70 (K+ gates open)
Na+ current: 120 × (0.90)³ × 0.05 = ~4 mS/cm² (now much smaller — inactivation dominated)
K+ current: 36 × (0.70)⁴ = 36 × 0.24 = 8.6 mS/cm² × (V-E_K) with V now at 0: = 8.6 × (0-(-77)) = 663 µA/cm² (large outward = repolarizing)

THE UNDERSHOOT / AHP (t ≈ 2-3 ms):
V dips to -80 mV (below rest). Why? n is still elevated (τn at -65 mV ≈ 5 ms). Extra K+ conductance hyperpolarises toward E_K (-77 mV). This is the absolute refractory period (h still near 0 AND V below rest — both prevent a new spike).

RECOVERY (t > 5 ms):
h recovers at -80 mV (τh ≈ 5-10 ms at these voltages)
n closes (τn ≈ 5 ms at rest)
V returns to -65 mV
After ~10 ms: cell can fire again (relative refractory → absolute refractory → recovered)

PYTHON CODE (scipy ODE solver):
from scipy.integrate import odeint
import numpy as np

def hh_odes(y, t, I_ext=0):
    V, m, h, n = y
    # rate functions
    am = 0.1*(V+40)/(1-np.exp(-(V+40)/10)) if abs(V+40)>1e-7 else 1.0
    bm = 4*np.exp(-(V+65)/18)
    ah = 0.07*np.exp(-(V+65)/20)
    bh = 1/(np.exp(-(V+35)/10)+1)
    an = 0.01*(V+55)/(1-np.exp(-(V+55)/10)) if abs(V+55)>1e-7 else 0.1
    bn = 0.125*np.exp(-(V+65)/80)
    # conductances
    gNa, gK, gL = 120, 36, 0.3
    ENa, EK, EL = 55, -77, -54.4
    Cm = 1.0
    # currents
    INa = gNa * m**3 * h * (V - ENa)
    IK  = gK  * n**4     * (V - EK)
    IL  = gL             * (V - EL)
    # ODEs
    dV = (I_ext - INa - IK - IL) / Cm
    dm = am*(1-m) - bm*m
    dh = ah*(1-h) - bh*h
    dn = an*(1-n) - bn*n
    return [dV, dm, dh, dn]

t = np.arange(0, 50, 0.01)   # 50 ms, 0.01 ms steps
y0 = [-65, 0.05, 0.60, 0.32] # initial conditions
sol = odeint(hh_odes, y0, t, args=(10,))  # I_ext=10 µA/cm²
V = sol[:, 0]  # membrane voltage trace
This is essentially what the Day 1 notebook implements.""")

    # --- 5. Channel Extensions: A-current ---
    n += 1
    make_content_slide(prs,
        "Extension 1: The A-Current (Kv4, I_A)",
        ["The original HH model has: I_Na (Nav), I_K (Kv), I_L (leak)",
         "Real neurons have many more channel types. Extensions ADD currents.",
         "",
         "I_A (A-type K+ current, Kv4.1/4.2/4.3):",
         "  I_A = g_A × a³ × b × (V - E_K)",
         "  a = activation gate (fast, like m)",
         "  b = inactivation gate (fast — completely inactivates within ~50 ms)",
         "",
         "Biophysical role:",
         "  Activates near rest (-70 to -60 mV) — low threshold",
         "  Transiently delays the FIRST spike after hyperpolarisation",
         "  Controls inter-spike interval and firing rate",
         "  'Anti-burst' function: prevents rapid re-firing",
         "",
         "Drug relevance: Kv4 blockers (4-AP, fampridine) slow K+ repolarisation",
         "  Clinical: fampridine (Ampyra) for MS — improves walking",
         "  Problem: 4-AP also blocks hERG (cardiac risk!)"],
        n,
        notes="""THE A-CURRENT — WHAT IT IS AND WHY IT MATTERS:

The A-current was first described by Connor & Stevens (1971) in molluscan neurons. It is a TRANSIENT outward K+ current — it turns on quickly when the membrane depolarizes, then inactivates within 10-100 ms, even if the membrane stays depolarized.

GATING VARIABLES (similar to m/h in HH):
a (activation): fast, voltage-dependent. Half-activated at ~-45 mV. τa ~ 1-5 ms.
b (inactivation): fast. Half-inactivated at ~-80 mV. τb ~ 10-100 ms.
At rest (-70 mV): b is near 1 (not inactivated). a is near 0 (not activated).

THE CRITICAL FEATURE:
After hyperpolarization (e.g., during the AHP at -80 mV):
- b recovers fully (b → 1)
- When the neuron is subsequently depolarized, a activates quickly at low voltages
- This creates a transient outward K+ current that slows the rate of depolarization
- The neuron takes longer to reach threshold → delayed first spike

This mechanism controls how often a neuron fires. In neurons where I_A is strong, repeated stimuli generate spikes at a slow, regular rate. In neurons where I_A is weak, they burst fire (rapid succession of spikes).

CABLE THEORY NOTE:
I_A is particularly important in dendrites (the input branches of neurons). Kv4 channels are densely expressed in hippocampal and cortical pyramidal cell dendrites. They regulate how much of a synaptic input reaches the cell body (somatic depolarization). Drugs that block Kv4 in dendrites → larger EPSPs → more excitable neurons.

CLINICAL RELEVANCE:
Fampridine (4-aminopyridine, Ampyra): FDA-approved for multiple sclerosis. MS causes demyelination (loss of myelin sheath). Demyelinated axons conduct action potentials poorly — the AP amplitude decays as it travels. By blocking K+ channels (including A-type), fampridine prolongs the AP → more Na+ enters → AP is larger and travels further. Clinically: improves walking speed in ~35% of MS patients.

DRUG DESIGN CHALLENGE:
4-AP also blocks hERG (Kv11.1, cardiac K+ channel). hERG block → cardiac arrhythmia. For fampridine, the dose is low enough that cardiac risk is manageable. But designing SELECTIVE Kv4 blockers that don't touch hERG is a major medicinal chemistry challenge (and a potential QSAR project!).

HH EXTENSION CODE:
# Add I_A to the HH model
def ia_current(V, a, b, gA=20, EK=-77):
    return gA * a**3 * b * (V - EK)

def a_inf(V): return 1/(1+np.exp(-(V+50)/20))
def tau_a(V): return 0.5 + 2/(1+np.exp((V+50)/15))
def b_inf(V): return 1/(1+np.exp((V+80)/6))
def tau_b(V): return 10 + 80/(1+np.exp(-(V+70)/10))

# Add da/dt and db/dt to the ODE system
da = (a_inf(V) - a) / tau_a(V)
db = (b_inf(V) - b) / tau_b(V)""")

    # --- 6. Calcium channels ---
    n += 1
    make_content_slide(prs,
        "Extension 2: Voltage-Gated Calcium Channels (Cav)",
        ["Ca2+ is the universal second messenger in the nervous system",
         "  Triggers: NT release, muscle contraction, gene expression",
         "  Intracellular Ca2+ at rest: ~100 nM | Extracellular: ~2 mM",
         "  → 20,000x gradient PLUS electrical gradient → huge driving force",
         "",
         "HH extension: I_Ca = g_Ca × m² × h_Ca × (V - E_Ca)",
         "  E_Ca ≈ +140 mV (high because of the extreme concentration gradient)",
         "",
         "Cav subtypes and drug targets:",
         "  Cav1 (L-type): heart, muscle, dendrites",
         "    → Ca2+ channel blockers (nifedipine, verapamil) for hypertension/angina",
         "  Cav2.1 (P/Q-type): presynaptic NT release",
         "    → Omega-agatoxin (spider toxin) research tool",
         "  Cav2.2 (N-type): presynaptic, pain pathways",
         "    → Ziconotide (Prialt) for severe pain (cone snail venom)",
         "  Cav3 (T-type): pacemaking, low-threshold oscillations",
         "    → Ethosuximide for absence epilepsy"],
        n,
        notes="""CALCIUM AS A SECOND MESSENGER:
The term 'second messenger' refers to molecules inside cells that relay and amplify signals from the 'first messengers' (hormones, neurotransmitters). Ca2+ is used as a second messenger in virtually every cell type.

WHY IS Ca2+ SPECIAL?
At rest, cytoplasmic Ca2+ = ~100 nM (nanomolar). Extracellular Ca2+ = ~2 mM = 2,000,000 nM. The gradient is 20,000-fold. Any time a Ca2+ channel opens, a large Ca2+ current flows in. This large amplitude makes Ca2+ an excellent signal molecule — even a brief channel opening causes a detectable local concentration rise.

CALCIUM-DEPENDENT PROCESSES:
1. NT release: synaptotagmin detects Ca2+ → SNARE fusion → vesicle exocytosis
2. Muscle contraction: Ca2+ binds troponin C → exposes actin binding sites for myosin
3. Gene expression: Ca2+/calmodulin → CaM kinase II → CREB phosphorylation → gene transcription (long-term memory consolidation!)
4. Cell death: excessive Ca2+ influx → mitochondrial damage → apoptosis or necrosis (relevant for stroke, neurodegeneration)

CAV SUBTYPES:
L-type (Cav1.1-1.4): Long-lasting, Large conductance. In cardiac muscle: drives the plateau phase of cardiac AP. In smooth muscle: causes vasoconstriction (blood pressure control). In neurons: somatodendritic; regulates gene expression.
- Blockers: dihydropyridines (nifedipine, amlodipine) = arterial vasodilators = antihypertensives/antianginals. Phenylalkylamines (verapamil) = cardiac (slow heart rate, antiarrhythmic). Benzothiazepines (diltiazem) = mixed.

P/Q-type (Cav2.1): Present and tightly coupled to SNARE proteins at most central synapses. Required for fast NT release.
- Toxin: ω-agatoxin IVA (funnel-web spider) blocks Cav2.1. Research tool.
- Clinical: mutations in Cav2.1 cause familial hemiplegic migraine (FHM) and spinocerebellar ataxia type 6 (SCA6).

N-type (Cav2.2): Present at presynaptic terminals especially in dorsal horn (pain processing). Also in sympathetic neurons.
- Blocker: ziconotide (Prialt, ω-conotoxin MVIIA from cone snail Conus magus). FDA-approved for severe chronic pain (intrathecal infusion). Very potent (IC50 ~ pM). Cannot be given IV (blocks neuromuscular junction globally). One of the most potent analgesics known.

T-type (Cav3.1, 3.2, 3.3): Transient, Tiny, Low-threshold. Activates at low voltages (~-60 mV), inactivates quickly, contributes to rhythmic oscillations (pacemaking) in thalamic neurons, SA node of heart.
- Blocker: ethosuximide (Zarontin) = first-line for absence epilepsy. Absence seizures involve pathological thalamocortical oscillations; T-type Cav in thalamic neurons drives these oscillations. Block T-type → reduce oscillations → fewer absence seizures.

HH EXTENSION CODE:
# L-type Ca2+ channel (simplified, after Huguenard 1996)
def ica_current(V, mCa, hCa, gCa=0.5, ECa=140):
    return gCa * mCa**2 * hCa * (V - ECa)

# Note: ECa is dynamic because Ca2+ is a second messenger
# Full model uses Ca2+ concentration ODE:
# d[Ca]_i/dt = -I_Ca / (2*F*vol) - ([Ca]_i - [Ca]_rest) / tau_Ca""")

    # --- 7. Ih (HCN channels) ---
    n += 1
    make_content_slide(prs,
        "Extension 3: The Ih Current (HCN Channels)",
        ["HCN = Hyperpolarization-activated Cyclic Nucleotide-gated channel",
         "  Carries: mixed Na+/K+ current (both flow, net reversal ~-30 mV)",
         "  ACTIVATED by HYPERPOLARIZATION (opposite of most channels!)",
         "",
         "HH extension: I_h = g_h × q × (V - E_h)",
         "  q = activation gate, opens when V < -60 mV (slow, τq ~ 50-500 ms)",
         "  E_h ≈ -30 mV (depolarizing at rest)",
         "  q∞(V): activation at -60 to -120 mV",
         "",
         "Functions in neurons:",
         "  'Pacemaker current': drives rhythmic firing in SA node, thalamus, DRG",
         "  'Sag potential': after hyperpolarization, Ih activates → depolarising 'sag'",
         "  Stabilises resting potential (prevents runaway hyperpolarization)",
         "  Modulated by cAMP: β-adrenergic stimulation → faster heart rate",
         "",
         "Drug relevance: Ivabradine (Coralan) = selective If (cardiac HCN) blocker",
         "  Reduces heart rate WITHOUT reducing contractility",
         "  FDA-approved for heart failure and stable angina"],
        n,
        notes="""HCN CHANNELS — THE BACKWARDS CHANNEL:

Most ion channels open when the membrane DEPOLARIZES. HCN channels are unique: they open when the membrane HYPERPOLARIZES. After a hyperpolarising input or the AHP following a spike, V dips below -60 mV → HCN opens → Na+ and K+ both flow → net current is slightly inward (depolarising) → pushes the membrane back toward rest.

THE FOUR HCN SUBTYPES:
HCN1: fast kinetics, expressed in cortex, hippocampus (CA1 dendrites), DRG. τ ~ 50-100 ms at -100 mV.
HCN2: slower, widely expressed, strongly modulated by cAMP. τ ~ 200-500 ms.
HCN3: slowest, limited expression.
HCN4: dominant in SA node of heart, very slow. τ ~ 500-5000 ms. The cardiac pacemaker current (If = 'funny current').

CYCLIC NUCLEOTIDE MODULATION:
HCN channels have a cyclic nucleotide binding domain (CNBD) on the intracellular C-terminus. When cAMP binds, it SHIFTS the activation curve to more positive voltages (channels open more easily). This is how the sympathetic nervous system speeds up the heart: norepinephrine → β1-adrenergic receptor → Gs protein → adenylyl cyclase → more cAMP → HCN4 activates more → faster pacemaking → heart rate up.

PARASYMPATHETIC (vagus): acetylcholine → M2 muscarinic → Gi → less cAMP → HCN4 activates less → slower heart rate.

IVABRADINE (Coralan, Procoralan):
The first selective If blocker. Mechanism: binds inside the HCN4 pore from the cytoplasmic side when the channel is open. Clinically: reduces resting heart rate by ~10-15 bpm without reducing contractility (unlike beta-blockers) and without affecting blood pressure.
FDA approvals: (1) Heart failure with reduced ejection fraction (HFrEF) when heart rate ≥ 70 bpm and patient can't tolerate beta-blockers. (2) Stable angina (European approval).
Side effects: visual phosphenes (brief light flashes) because HCN1 in retinal photoreceptors is also blocked.

NEUROLOGICAL RELEVANCE:
HCN2 mutations cause epilepsy. HCN1 mutations cause early-onset epileptic encephalopathy. Lamotrigine, gabapentin, and some other anticonvulsants may partially modulate HCN.

SAG POTENTIAL (signature of HCN activity):
If you inject a hyperpolarizing current pulse into a neuron with Ih: V first hyperpolarizes, then gradually sags back toward rest (even while the pulse is still on). This 'sag' is the signature of active Ih. When the pulse ends, V briefly depolarizes above rest (rebound depolarization) as Ih turns off. This can trigger a rebound spike (used in thalamic relay neurons for generating sleep spindles).

HH EXTENSION CODE:
def ih_current(V, q, gh=0.02, Eh=-30):
    return gh * q * (V - Eh)

def q_inf(V):   return 1/(1+np.exp((V+75)/5.5))   # opens at -60 to -100 mV
def tau_q(V):   return 200 + 300/(1+np.exp((V+72)/4))  # very slow, 200-500 ms""")

    # --- 8. Persistent Nav ---
    n += 1
    make_content_slide(prs,
        "Extension 4: Persistent Na+ Current (I_NaP)",
        ["A small fraction of Nav channels never inactivate:",
         "  I_NaP = g_NaP × p∞(V) × (V - E_Na)",
         "  p = activation (no inactivation gate)",
         "  g_NaP ~ 0.1 - 1% of g_Na (small but crucial!)",
         "  Activates at subthreshold voltages: -70 to -55 mV",
         "",
         "Why it matters despite being small:",
         "  Amplifies subthreshold depolarizations",
         "  Creates bistability (two stable states: rest and depolarised plateau)",
         "  Enables plateau potentials and persistent depolarised states",
         "  Critical for: motor pattern generation, intrinsic oscillations",
         "",
         "Pathological roles:",
         "  Excessive I_NaP → hyperexcitability → epilepsy",
         "  SCN8A gain-of-function mutations → high I_NaP → severe epilepsy",
         "  ALS: I_NaP amplifies inputs to motor neurons",
         "",
         "Drug target: riluzole (Rilutek) — I_NaP blocker",
         "  FDA-approved for ALS — extends life by ~3 months",
         "  Reduces I_NaP → less motor neuron hyperexcitability"],
        n,
        notes="""PERSISTENT SODIUM CURRENT — SMALL BUT MIGHTY:

The original HH model assumes ALL Nav channels completely inactivate (h → 0 on sustained depolarization). In reality, a small fraction (~0.5-2%) of Nav channels in most neurons never fully inactivate — they remain open (or rapidly reopen) even during sustained depolarization.

MECHANISTIC BASIS:
The persistent component comes from:
1. Late openings of the transient Nav channel (the inactivation gate occasionally fails to close → 'window current' within a voltage range where m∞(V) × h∞(V) > 0, typically -70 to -55 mV)
2. A distinct isoform (Nav1.6 has prominent I_NaP in many neurons)
3. Phosphorylation-dependent modulation of Nav inactivation

THE WINDOW CURRENT:
In the HH model, the 'window current' is the steady-state Na+ conductance at intermediate voltages where both m∞ and h∞ are nonzero. At -60 mV: m∞ ≈ 0.08, h∞ ≈ 0.55 → m³h ≈ 0.0003. Small but nonzero. This contributes a small inward current at subthreshold voltages.

WHY A SMALL CURRENT MAKES A BIG DIFFERENCE:
At subthreshold voltages (-70 to -55 mV), the net membrane resistance is very high (few channels open). A tiny current has a large voltage effect (V = I × R; high R means high ΔV for same I). I_NaP at -65 mV might be only 0.05 µA/cm² but the input resistance is ~50 MΩ → this creates mV-scale depolarization that amplifies synaptic inputs.

BISTABILITY:
With sufficient I_NaP, the neuron can have two stable resting states:
1. Normal rest: -65 mV (I_NaP small)
2. Depolarised plateau: -40 to -30 mV (I_NaP large, balanced by K+ currents)
Transitions between these states are triggered by brief stimuli. This underlies 'plateau potentials' in spinal motor neurons (relevant to locomotion) and some pathological depolarized states.

RILUZOLE (ALS DRUG):
Riluzole was the first FDA-approved drug for ALS (1995). Mechanism: blocks I_NaP (and to a lesser extent, the transient Nav current). ALS motor neurons have increased I_NaP → excessive activity → glutamate excitotoxicity → cell death. By reducing I_NaP, riluzole reduces motor neuron hyperexcitability and glutamate release.
Effect: extends survival by ~2-3 months on average. Not a cure, but the first disease-modifying treatment for ALS.

RILUZOLE QSAR:
Riluzole is a benzothiazole compound. SMILES: Nc1nc2c(s1)cc(OCC(F)(F)F)cc2.
The 2-amino group and the thiazole ring are critical for Nav block. The CF3-OCH2 group provides metabolic stability and appropriate lipophilicity for CNS penetration (LogP ≈ 1.5, MW = 234 Da — fits all CNS rules).

HH EXTENSION CODE:
def inap_current(V, gNaP=0.1, ENa=55):
    p_inf = 1/(1+np.exp(-(V+50)/3))  # activates ~-60 mV, no inactivation
    return gNaP * p_inf * (V - ENa)""")

    # --- 9. Synaptic conductances ---
    n += 1
    make_content_slide(prs,
        "Extension 5: Synaptic Conductances",
        ["Synaptic inputs = transient conductance changes driven by NT release",
         "  I_syn = g_syn(t) × (V - E_syn)",
         "  g_syn(t) follows an alpha-function or double-exponential shape",
         "",
         "Key synaptic conductances (each is a separate HH-style current):",
         "  AMPA (glutamate, excitatory): E_syn ≈ 0 mV, fast (τ ~ 5 ms)",
         "    → Fast excitatory EPSPs. AMPA antagonists: CNQX, perampanel (anti-epileptic)",
         "  NMDA (glutamate, excit.): E_syn ≈ 0 mV, slow (τ ~ 100 ms) + Mg2+ block",
         "    → Mg2+ block removed at depolarised V — Hebbian learning/LTP",
         "    → NMDA antagonists: ketamine (anesthetic, antidepressant), memantine (Alzheimer's)",
         "  GABA-A (inhibitory): E_syn ≈ -70 mV, fast (τ ~ 10 ms)",
         "    → Fast inhibition. Modulated by BZDs, alcohol, propofol",
         "  GABA-B (inhibitory, GPCR): E_syn ≈ -90 mV, slow (τ ~ 200 ms)",
         "    → Baclofen agonist (spasticity treatment)",
         "",
         "Double-exponential: g(t) = g_max × (exp(-t/τ1) - exp(-t/τ2)) / normalisation"],
        n,
        notes="""SYNAPTIC CONDUCTANCES IN THE HH FRAMEWORK:

When a presynaptic neuron fires and releases NT, the postsynaptic neuron experiences a transient change in conductance. This is added to the HH equations as an additional current term:

I_total = I_Na + I_K + I_L + I_A + I_Ca + I_h + I_NaP + I_AMPA + I_NMDA + I_GABAA + I_GABAB + ...

Each synaptic current has the form: I_syn = g_syn(t) × (V - E_syn)
where g_syn(t) is a time-dependent conductance that rises and decays (modelling the NT binding/unbinding kinetics).

THE ALPHA FUNCTION (HH style): g(t) = g_max × (t/τ) × exp(1-t/τ) for t > 0
Simple, one time constant. Peak at t = τ.

DOUBLE EXPONENTIAL (more biophysical): g(t) = g_max × (exp(-t/τ_decay) - exp(-t/τ_rise)) / norm
Rise time constant τ_rise (fast, NT binding). Decay time constant τ_decay (slow, NT unbinding + channel gating).

AMPA RECEPTORS:
AMPA (α-amino-3-hydroxy-5-methyl-4-isoxazolepropionic acid receptor): fast excitatory. Opens Na+/K+ channel (E_rev ≈ 0 mV). τ_rise ~ 0.5 ms, τ_decay ~ 3-5 ms.
How CS students can think about it: AMPA provides the fast, transient excitatory input. Like a brief clock pulse to a flip-flop.
Drug targets:
- CNQX (6-cyano-7-nitroquinoxaline-2,3-dione): competitive AMPA antagonist. Research tool.
- Perampanel (Fycompa): non-competitive AMPA antagonist. FDA-approved anticonvulsant (focal seizures). Reduces fast excitation in hyperactive circuits.

NMDA RECEPTORS:
NMDA (N-methyl-D-aspartate receptor): slow excitatory. Opens Ca2+/Na+/K+ channel (E_rev ≈ 0 mV). But: requires TWO conditions simultaneously to open: (1) glutamate must be bound AND (2) membrane must be depolarized. Why? At rest, Mg2+ blocks the pore from the outside. When V is depolarized, Mg2+ is expelled. This voltage-dependent block makes NMDA a COINCIDENCE DETECTOR: input (glutamate) AND depolarization must occur together.

Mg2+ block modelling: g_NMDA = g_max × (exp(-t/τ) - ...) × 1/(1 + [Mg]/3.57 × exp(-V/16.13))

SIGNIFICANCE FOR LEARNING AND MEMORY:
LTP (long-term potentiation) = the cellular mechanism of learning. When both pre AND postsynaptic neurons fire together, NMDA opens, Ca2+ enters, activates CaM kinase II → AMPA receptors inserted → synapse strengthened → Hebbian learning ('fire together, wire together').

DRUG TARGETS OF NMDA:
Ketamine: non-competitive open-channel NMDA blocker. At low doses: anesthetic. At sub-anesthetic doses: rapid antidepressant (mechanism still debated, but NMDA block → disinhibition → BDNF release → synaptogenesis). IV ketamine and intranasal esketamine (Spravato) FDA-approved for treatment-resistant depression.
Memantine: low-affinity NMDA blocker. FDA-approved for moderate-to-severe Alzheimer's. Hypothesis: in Alzheimer's, excessive Glu release → chronic NMDA activation → Ca2+ excitotoxicity → neuron death. Memantine blocks this while allowing normal synaptic signaling.
Phencyclidine (PCP, 'angel dust'): potent NMDA blocker. High abuse potential, causes psychosis. Historically: first demonstration that NMDA block causes schizophrenia-like symptoms → 'glutamate hypothesis of schizophrenia'.

GABA-A (ionotropic, fast inhibition):
E_syn = -70 to -75 mV (near resting potential). Fast: τ_rise ~ 0.5 ms, τ_decay ~ 5-25 ms. Cl- flows in → hyperpolarizes OR shunts excitatory currents.

GABA-B (metabotropic, slow inhibition):
GPCR (Gi-coupled). Activates GIRK (G-protein-inwardly-rectifying K+) channels. K+ flows out → hyperpolarization near E_K (~-90 mV). Very slow: onset 50-100 ms, duration 100-500 ms. 
Baclofen: GABA-B agonist. Used for: muscle spasticity (MS, spinal cord injury), alcohol withdrawal, trigeminal neuralgia. Intrathecal baclofen pump for severe spasticity.

HH EXTENSION CODE:
def ampa_current(V, t_syn, gAMPA=0.5, E_AMPA=0, tau_r=0.5, tau_d=5):
    if t_syn < 0:
        return 0
    norm = tau_d*tau_r/(tau_d-tau_r)
    g = gAMPA * norm * (np.exp(-t_syn/tau_d) - np.exp(-t_syn/tau_r))
    return max(0, g) * (V - E_AMPA)""")

    # --- 10. Multicompartment and cable theory ---
    n += 1
    make_content_slide(prs,
        "Extension 6: Cable Theory & Multicompartment Models",
        ["The single-compartment HH assumes isopotential membrane — all one voltage",
         "Real neurons have dendrites (up to 1 mm long!) — voltage DECAYS with distance",
         "",
         "Cable equation: Cm × dV/dt = λ² × d²V/dx² - V/τm + I_ext/rm",
         "  λ = space constant = sqrt(rm × ri / ra) — how far signal travels",
         "  τm = Cm × rm = membrane time constant",
         "  ri = internal resistance, ra = axial resistance",
         "",
         "Multicompartment models: discretise the dendrite into N compartments",
         "  Each compartment is an HH model connected to neighbours by resistors",
         "  Large models: >1000 compartments (Purkinje cell, layer 5 pyramidal)",
         "  Software: NEURON simulator, Brian2, NetPyNE",
         "",
         "Drug effects in multicompartment context:",
         "  Nav block in axon hillock → no spike propagation",
         "  Kv4 block in dendrites → larger distal EPSPs (synaptic amplification)",
         "  NMDA-dependent Ca2+ in dendrites → local spike initiation"],
        n,
        notes="""WHY SINGLE-COMPARTMENT HH IS NOT ENOUGH FOR REAL NEURONS:

The HH equations (as presented) treat the entire neuron as a single point — all membrane sees the same voltage at the same time. This is the 'isopotential' approximation. It works for:
- The squid giant axon cross-section (the axon is very long but uniform, and they recorded from one spot)
- Small neurons (spinal interneurons, striatal medium spiny neurons)
- Simplified network models where exact dendritic morphology doesn't matter

But for PYRAMIDAL NEURONS (the main excitatory neurons in cortex and hippocampus), the dendritic tree extends up to 1 mm. A synaptic input to the distal end of a dendrite must travel 1 mm to reach the soma (cell body) where the spike is initiated. Voltage DECAYS exponentially along the dendrite.

THE CABLE EQUATION:
Hodgkin & Huxley also wrote the cable equation (but it's not usually called 'the HH equation' — it's attributed to Kelvin from telegraph cable theory, adapted by Wilfrid Rall):
Cm × ∂V/∂t = (d/(4Ra)) × ∂²V/∂x² - I_ion

where:
- d = dendrite diameter (µm)
- Ra = axial resistivity (~100-150 Ω·cm for cytoplasm)
- x = distance along dendrite

KEY PARAMETERS:
Space constant λ = sqrt(Rm × d / (4Ra)) where Rm is specific membrane resistance (Ω·cm²)
- λ = distance over which voltage decays to 1/e (~37%) of its source amplitude
- Typical λ for a dendrite: 0.1-0.5 mm
- For a 1 mm dendrite: signal at the far end = e^(-1/0.3) = e^(-3.3) = only 4% of source amplitude!

Time constant τm = Rm × Cm
- Typical τm = 10-50 ms
- Determines how fast the membrane responds to injected current (RC time constant)

MULTICOMPARTMENT MODELS:
Discretise the cable into N segments, each with its own HH equations plus current flow from neighbours:
For compartment i: Cm × dVi/dt = I_ion_i + (V_{i-1} - Vi) / Ra_i + (V_{i+1} - Vi) / Ra_i+1

Famous models:
- Mainen-Sejnowski pyramidal cell (1996): first realistic compartmental model of a cortical L5 pyramidal neuron
- Purkinje cell model (De Schutter & Bower 1994): 1600 compartments, shows complex Ca2+-dependent bursting
- NEURON simulator (Carnevale & Hines 2006): standard software for multicompartment simulations

DRUG EFFECTS IN MULTICOMPARTMENT CONTEXT:
1. Axon hillock Nav block (local anesthetic): if Nav at the spike initiation zone is blocked, no AP is generated regardless of what happens in dendrites. This is the mechanism of local anesthesia.
2. Dendritic Kv4 block (4-AP): reduces shunting inhibition in dendrites → EPSPs are larger when they reach the soma → neuron more excitable. This is one mechanism of 4-AP's pro-convulsant and anti-MS effects.
3. Dendritic NMDA spikes: in distal dendrites, if AMPA + NMDA are activated strongly enough, a local 'dendritic spike' can be initiated even before the signal reaches the soma. NMDA receptor blockers (ketamine) prevent this → alter dendritic integration profoundly.

FOR CS STUDENTS:
A multicompartment model is essentially a graph neural network where each node is an HH model and each edge is a resistor (coupling adjacent compartments). The message passing is: currents flow between compartments according to Ohm's law. This is literally the same architecture as a GNN (Day 3)! The difference: in GNNs, we LEARN the message functions. In cable theory, the message function is Ohm's law (physics-based).""")

    # --- 11. Drug targets summary ---
    n += 1
    make_content_slide(prs,
        "HH Model Extensions: Drug Target Map",
        ["Channel         | Gene family  | Key Drug Target(s)",
         "─────────────────────────────────────────────────────",
         "Nav (I_Na)      | SCN1A-11A   | Lidocaine, carbamazepine, TTX",
         "Kv (I_K)        | KCNQ, Kv2   | 4-AP, TEA, fampridine",
         "Kv (I_A)        | KCND (Kv4)  | 4-AP (Ampyra for MS)",
         "Kv11.1 hERG     | KCNH2       | AVOID block (cardiac arrhythmia!)",
         "Cav L-type      | CACNA1C/D   | Nifedipine, verapamil, diltiazem",
         "Cav N-type      | CACNA1B     | Ziconotide (Prialt) for pain",
         "Cav T-type      | CACNA1G/H   | Ethosuximide (absence epilepsy)",
         "HCN (Ih)        | HCN1-4      | Ivabradine (Coralan) for HF",
         "Nav persistent  | SCN8A/2A    | Riluzole (ALS), lacosamide",
         "AMPA            | GRIA1-4     | Perampanel (epilepsy)",
         "NMDA            | GRIN1/2A-D  | Ketamine, memantine (Alzheimer's)",
         "GABA-A          | GABRA/B/G   | Diazepam, propofol, alcohol",
         "GABA-B          | GABBR1/2    | Baclofen (spasticity)"],
        n,
        notes="""DRUG TARGET MAP — COMPLETE GUIDE:

This slide shows the direct connection between the HH model components and real drugs. Every row represents a channel that is part of an HH-style model, the gene family that encodes it, and example drugs.

HOW TO READ THIS TABLE:
Column 1 (Channel/Current): This is the HH model component. All of these are ion channels that contribute current terms to the HH ODE.
Column 2 (Gene family): The gene(s) encoding the channel. QSAR models are trained on experimental data (IC50 values) for each target. Each gene = one QSAR model.
Column 3 (Drug targets): Approved drugs or important pharmacological tools that interact with this channel.

KEY PATTERNS:
1. ALMOST EVERY DRUG USED IN NEUROLOGY/PSYCHIATRY targets something in this table. This is because neurological diseases are diseases of electrical excitability.

2. SELECTIVITY is the challenge. Lidocaine blocks Nav1.7 (pain), but also blocks Nav1.1 (heart), Nav1.5 (skeletal muscle). Selective Nav1.7 blockers have been sought for 20+ years for pain treatment (would avoid cardiac side effects). This is a major QSAR challenge.

3. hERG (KCNH2) is the most dangerous off-target. Every drug candidate must be tested for hERG block. A dedicated QSAR model for hERG IC50 is one of the first models any pharmaceutical company builds.

4. AI APPLICATIONS IN THIS TABLE:
- Build QSAR models for each channel target
- Predict which drugs have which off-target activities (polypharmacology)
- Design selective compounds that hit the desired target but not others
- Predict which patients will respond to which drug (pharmacogenomics — do they have the right Nav1.2 variant?)

5. LACOSAMIDE (Vimpat): a newer anticonvulsant (Na channel modulator). Enhances slow inactivation of Nav channels. This is a distinct mechanism from the use-dependent fast inactivation stabilisation of lidocaine. Lacosamide has a unique interaction with a binding site at the inactivated Nav state.

6. PERSONALIZED MEDICINE EXAMPLE: SCN1A mutations cause Dravet syndrome (severe childhood epilepsy). Loss-of-function in SCN1A (Nav1.1) reduces inhibitory neuron firing → network excitability → seizures. Treatment: AVOID sodium channel blockers (they would further reduce Nav1.1) → use valproate, clobazam, stiripentol. An AI model that predicts drug-Nav1.1 interaction AND knows the patient's SCN1A variant could guide personalized treatment choice.""")

    # --- 12. Python simulation reference ---
    n += 1
    make_content_slide(prs,
        "Python Reference: Simulating the HH Model",
        ["Step 1 — Define rate functions:",
         "  alpha_m(V): opening rate for m gate",
         "  beta_m(V): closing rate for m gate",
         "  (same pattern for h, n)",
         "",
         "Step 2 — Define ODE system:",
         "  [dV/dt, dm/dt, dh/dt, dn/dt] = f(V, m, h, n, I_ext)",
         "",
         "Step 3 — Set initial conditions:",
         "  V0=-65, m0=0.05, h0=0.60, n0=0.32",
         "",
         "Step 4 — Integrate:",
         "  scipy.integrate.odeint(system, y0, t, args=(I_ext,))",
         "",
         "Step 5 — Plot V(t), m(t), h(t), n(t)",
         "",
         "Extensions: add I_A, I_Ca, I_h by appending gate ODEs and current terms",
         "Packages: scipy (ODE solver), numpy (arrays), matplotlib (plots)",
         "Full simulation in: day1_introduction/Day1_Practical_HH.ipynb"],
        n,
        notes="""FULL COMMENTED PYTHON CODE FOR THE HH MODEL:

import numpy as np
from scipy.integrate import odeint
import matplotlib.pyplot as plt

# ---- Rate functions ----
def alpha_m(V):
    # Opening rate for m gate (Na+ activation). Units: 1/ms
    dV = V + 40  # shift by 40 (V=-40 is a singularity)
    if abs(dV) < 1e-7:
        return 1.0  # L'Hopital limit at the singularity
    return 0.1 * dV / (1 - np.exp(-dV / 10))

def beta_m(V):
    # Closing rate for m gate. Units: 1/ms
    return 4 * np.exp(-(V + 65) / 18)

def alpha_h(V):
    # Opening rate for h gate (Na+ inactivation). Units: 1/ms
    return 0.07 * np.exp(-(V + 65) / 20)

def beta_h(V):
    # Closing rate for h gate. Units: 1/ms
    return 1 / (np.exp(-(V + 35) / 10) + 1)

def alpha_n(V):
    # Opening rate for n gate (K+ activation). Units: 1/ms
    dV = V + 55
    if abs(dV) < 1e-7:
        return 0.1
    return 0.01 * dV / (1 - np.exp(-dV / 10))

def beta_n(V):
    # Closing rate for n gate. Units: 1/ms
    return 0.125 * np.exp(-(V + 65) / 80)

# ---- Channel parameters ----
gNa = 120.0   # Maximum Na+ conductance (mS/cm²)
gK  =  36.0   # Maximum K+ conductance (mS/cm²)
gL  =   0.3   # Leak conductance (mS/cm²)
ENa =  55.0   # Na+ reversal potential (mV)
EK  = -77.0   # K+ reversal potential (mV)
EL  = -54.4   # Leak reversal potential (mV)
Cm  =   1.0   # Membrane capacitance (µF/cm²)

# ---- ODE system ----
def hodgkin_huxley(y, t, I_ext):
    V, m, h, n = y
    
    # Ion currents
    I_Na = gNa * m**3 * h * (V - ENa)   # Inward (typically negative in convention)
    I_K  = gK  * n**4     * (V - EK)    # Outward
    I_L  = gL             * (V - EL)    # Leak
    
    # Membrane voltage ODE
    dV = (I_ext - I_Na - I_K - I_L) / Cm
    
    # Gating variable ODEs
    dm = alpha_m(V) * (1 - m) - beta_m(V) * m
    dh = alpha_h(V) * (1 - h) - beta_h(V) * h
    dn = alpha_n(V) * (1 - n) - beta_n(V) * n
    
    return [dV, dm, dh, dn]

# ---- Simulation ----
t = np.arange(0, 100, 0.01)            # 100 ms, 0.01 ms step
y0 = [-65.0, 0.05, 0.60, 0.32]         # Initial conditions at rest
I_ext = 10.0                           # Stimulus current (µA/cm²)

solution = odeint(hodgkin_huxley, y0, t, args=(I_ext,))
V = solution[:, 0]
m = solution[:, 1]
h = solution[:, 2]
n = solution[:, 3]

# ---- Plot ----
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6))
ax1.plot(t, V, 'k'); ax1.set_ylabel('Voltage (mV)'); ax1.set_title('HH Action Potential')
ax2.plot(t, m, label='m (Na act)'); ax2.plot(t, h, label='h (Na inact)')
ax2.plot(t, n, label='n (K act)'); ax2.legend(); ax2.set_ylabel('Gate probability')
ax2.set_xlabel('Time (ms)')
plt.tight_layout()
plt.savefig('hh_simulation.png', dpi=150)

TO ADD DRUG EFFECTS:
# Simulate lidocaine (50% Nav block): set gNa = 60 (instead of 120)
# Simulate 4-AP (K+ channel block): set gK = 18 (instead of 36)
# Simulate GABA-A drug: add I_GABAA current with appropriate E_syn and g(t)

The Day 1 notebook (day1_introduction/Day1_Practical_HH.ipynb) has a fully interactive version.""")

    # --- 13. Connection to QSAR / AI drug discovery ---
    n += 1
    make_content_slide(prs,
        "HH Model → QSAR → AI Drug Discovery: The Full Pipeline",
        ["Understanding HH helps you understand WHAT your QSAR model is predicting:",
         "",
         "QSAR predicts: IC50 for Nav block (e.g., 50 nM for compound X)",
         "HH simulation asks: IF Nav is blocked 80%, does the neuron still fire?",
         "SpikerBox validates: does compound X change neural firing in a real prep?",
         "",
         "The channel map connects QSAR targets to HH currents:",
         "  Nav1.2 IC50 → changes g_Na in HH → changes spike threshold/amplitude",
         "  hERG IC50 → changes g_K in HH → changes AP duration → QT interval",
         "  GABA-A IC50 → changes synaptic I_GABAA → changes network inhibition",
         "",
         "GNNs (Day 3) predict the same quantities but more accurately:",
         "  Molecular graph → GNN → predicted IC50 for any channel",
         "  AlphaFold → 3D structure of channel → docking → predicted pose",
         "  Combine: structure + QSAR → multi-target prediction",
         "",
         "Your project: you are predicting these IC50 values, designing molecules",
         "that hit the RIGHT channels and avoid the WRONG ones (hERG!)"],
        n,
        notes="""THE COMPLETE PIPELINE — CONNECTING HH TO AI DRUG DISCOVERY:

STEP 1 — TARGET IDENTIFICATION:
Which ion channel / receptor is dysfunctional in the disease?
Example: SCN1A (Nav1.1) LOF mutation → Dravet syndrome (severe epilepsy)
Example: CACNA1A (Cav2.1) GOF mutation → familial hemiplegic migraine
Example: KCNH2 (hERG) drug block → drug-induced arrhythmia (MUST AVOID)

STEP 2 — QSAR MODELING (today's lecture):
Collect experimental IC50 data for the target channel from ChEMBL / internal databases.
Train a QSAR model: molecular features (ECFP4 + descriptors) → predicted IC50.
Evaluate with scaffold split.
Interpret with SHAP → understand which molecular features drive potency.

STEP 3 — HH SIMULATION (physiological impact prediction):
Take the predicted IC50. Calculate what fraction of channels are blocked at a therapeutic concentration (using IC50 and plasma concentration estimates). Set g_channel = g_max × (1 - fraction_blocked) in the HH model. Simulate: does the neuron behave differently? Better (therapeutic)? Worse (side effect)?

STEP 4 — SPIKERBOX VALIDATION (in vitro):
Synthesise the top predicted compounds. Test on SpikerBox: add compound to preparation, observe neural firing change. Compare to HH prediction.

STEP 5 — ANIMAL MODELS (in vivo):
Test in rodent models of the disease (epilepsy: pentylenetetrazol seizure model; pain: von Frey test; arrhythmia: ECG QT measurement). Does the drug work? Are there side effects?

STEP 6 — CLINICAL TRIALS (human):
Phase I: safety in healthy volunteers. Phase II: dose-finding in patients. Phase III: efficacy vs. placebo in large trial. FDA/EMA approval.

THE QSAR MODEL'S ROLE IN THIS PIPELINE:
QSAR screens billions of virtual compounds at Step 2, filtering down to hundreds for synthesis and testing at Step 3-4. Without AI, Steps 2-3 would require synthesizing and testing every candidate — prohibitively expensive. With AI: test 10 × cheaper candidates, 10 × faster.

FOR STUDENT PROJECTS:
Project 1 (SpikerBot): You'll use the SpikerBox to validate QSAR predictions for Nav blockers in cockroach preparation.
Project 2 (GABA-A): Train a QSAR model for GABA-A modulator IC50. Predict selectivity over other Cys-loop channels.
Project 3 (Drug repurposing): Use HH simulation to predict which existing drugs could treat a neural channelopathy.
Project 4 (Eye-tracking): Use the pupillary light reflex (controlled by M-current, Ih, Nav in EW neurons) as an in vivo readout of drug effects on these HH channels.""")

    # Closing
    n += 1
    make_section_divider(prs, "You Now Understand the Mathematics of Neural Excitability 🧮🧠", n,
        notes="End of the HH Explainer deck. Students (or instructors) who have worked through this deck understand: (1) the 4 HH equations and what each term means physically, (2) the gating variables m, h, n and their biophysical basis, (3) extensions: A-current, Cav, Ih, I_NaP, synaptic conductances, cable theory, (4) drug targets at every level, (5) how HH connects to QSAR and AI drug discovery. This deck provides the theoretical foundation for interpreting all computational neuroscience results in the course.",
        subtitle="HH equations, gating variables, channel extensions, drug targets, Python code, AI connection — all covered!")

    os.makedirs("hh_explainer", exist_ok=True)
    prs.save("hh_explainer/slides.pptx")
    print(f"  HH Explainer: {n} slides -> hh_explainer/slides.pptx")



# =====================================================================
# MAIN
# =====================================================================
if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)) if os.path.dirname(os.path.abspath(__file__)) else ".")
    print("=" * 60)
    print("  AI for Drug Discovery — Slide Generator")
    print("  4 Teaching Days + Exam + 3 Quizzes + Annotated Day 2 + HH Explainer")
    print("  Quiz 1:          Introductory Drug Discovery Quiz")
    print("  Quiz 2:          Abbreviations & Biology Quiz")
    print("  Quiz 3:          Physiology & Neuroscience Quiz")
    print("  Day 1 (15 Apr):  Introduction, Pipeline, Projects")
    print("  Day 2 (22 Apr):  Molecular ML, QSAR, Evaluation")
    print("  Day 2 annotated: Full presenter notes for everything")
    print("  Day 3 (5 May):   Deep Learning, GNNs, AlphaFold")
    print("  Day 4 (19 May):  Ethics, Physiology, Workshop")
    print("  HH Explainer:    Hodgkin-Huxley deep-dive lecture deck")
    print("=" * 60)
    generate_quiz()
    generate_quiz_abbreviations()
    generate_quiz_physiology()
    generate_day1()
    generate_day2()
    generate_day2_annotated()
    generate_day3()
    generate_day4()
    generate_hh_explainer()
    print("=" * 60)
    print("  Done! All slides generated.")
    print("=" * 60)
