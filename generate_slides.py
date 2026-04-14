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

    # --- Slide 2: Morning Agenda ---
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


# ===== MAIN =====
if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)) if os.path.dirname(os.path.abspath(__file__)) else ".")
    print("=" * 60)
    print("  AI for Drug Discovery — Slide Generator")
    print("  4 Teaching Days + Exam + Quiz")
    print("  Quiz:           Introductory Drug Discovery Quiz")
    print("  Day 1 (15 Apr): Introduction, Pipeline, Projects — 6 units")
    print("  Day 2 (22 Apr): Molecular ML, QSAR, Evaluation — 8 units")
    print("  Day 3 (5 May):  Deep Learning, GNNs, AlphaFold — 8 units")
    print("  Day 4 (19 May): Ethics, Physiology, Workshop — 4 units")
    print("  Exam  (27 May): Presentations / Posters")
    print("=" * 60)
    generate_quiz()
    generate_day1()
    generate_day2()
    generate_day3()
    generate_day4()
    print("=" * 60)
    print("  Done! All slides generated.")
    print("=" * 60)
