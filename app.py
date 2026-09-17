import streamlit as st
from core.reference_loader import (
    docx_text, load_orientation_command, load_unified_short_example,
    simple_docx_format_issues,
)
from core.prompts import (
    CAREER_CONSISTENCY_RULE_EL, CAREER_CONSISTENCY_RULE_EN,
    build_orientation_source, build_orientation_prompt, split_orientation_response,
)
from core.docx_builder import build_orientation_docx, build_orientation_client_docx
from core.generator import generate_analysis
from core.validator import validate_orientation
from core.case_state import reset_case_state, handle_pdf_upload

# ---------------------------------------------------------------------------
# UI text in both languages. Internal identifiers used by core/validator.py
# and core/prompts.py (presentation, "Ναι"/"Όχι", the "Όνομα" key)
# stay as the original Greek literals no matter which UI language is shown —
# only what the person reads/labels-for changes. This mirrors the idea from
# the reviewed prototype (site-wide EL/EN toggle) without touching the
# validated core.
# ---------------------------------------------------------------------------
TR = {
    "el": {
        "hero_title": "AstroCheck Career",
        "hero_subtitle": "Ανάδειξη ταλέντων και επαγγελματικός προσανατολισμός για παιδί, έφηβο ή ενήλικα — από ένα Astrodienst PDF.",
        "sidebar_header": "AstroCheck Career",
        "sidebar_privacy": "Τα δεδομένα επεξεργάζονται στη συνεδρία και δεν αποθηκεύονται από την εφαρμογή.",
        "sidebar_reset": "🔄 Νέα ανάλυση (καθαρισμός όλων)",
        "sidebar_reset_help": "Καθαρίζει χάρτη και επιλογές, ώστε να ξεκινήσεις καθαρά με το επόμενο άτομο.",
        "step1_title": "Βήμα 1 · Ανέβασε το Astrodienst PDF",
        "pdf_uploader": "Astrodienst Data Sheet (.pdf)",
        "read_pdf": "Ανάγνωση PDF",
        "reading_spinner": "Διαβάζεται το PDF…",
        "pdf_error": "Η ανάγνωση σταμάτησε με ασφάλεια — το PDF μπορεί να μην είναι το σωστό Astrodienst Data Sheet, ή η μορφή του διαφέρει.",
        "technical_detail": "Τεχνική λεπτομέρεια",
        "chart_loaded": "✓ Ελεγμένος χάρτης στη συνεδρία: **{name}**. Ανέβασε νέο PDF μόνο αν θέλεις να τον αντικαταστήσεις, ή πάτα «🔄 Νέα ανάλυση» στο πλάι.",
        "step2_title": "Βήμα 2 · Όνομα",
        "name_label": "Όνομα για το τελικό έγγραφο",
        "presentation_simple": "Απλή και πρακτική",
        "success_simple": "Θα ζητηθεί κείμενο 4–5 σελίδων, με καθημερινή γλώσσα, 4–6 ταλέντα, 4–5 επαγγελματικούς τομείς και σύντομη τελική σύνθεση.",
        "no_extra_data": "Δεν ζητούνται πρόσθετα προσωπικά, ψυχολογικά, σχολικά ή οικονομικά δεδομένα. Η υπηρεσία παρουσιάζει μόνο συμβολικές πιθανότητες προς διερεύνηση από τον ελεγμένο χάρτη.",
        "step3_title": "Βήμα 3 · Λήψη εντολής για ChatGPT/Claude",
        "download_command": "⬇️ Λήψη εντολής προσανατολισμού για ChatGPT/Claude",
        "upload_hint_simple": "Ανέβασε το Word στο ChatGPT ή στο Claude. Ζήτησε δύο αρχεία: το καθαρό Word του πελάτη και το εσωτερικό τεχνικό δελτίο ελέγχου.",
        "paste_expander": "Έτοιμο μήνυμα για επικόλληση στο ChatGPT/Claude",
        "auto_title": "#### 🤖 Εναλλακτικά: αυτόματη δημιουργία",
        "auto_caption": "Χρειάζεται δικό σου OpenAI API key (δεν αποθηκεύεται πουθενά). Παρακάμπτει τα βήματα αντιγραφής προς/από ChatGPT ή Claude — το αποτέλεσμα περνάει αυτόματα από τον ίδιο μηχανικό έλεγχο.",
        "auto_button": "Αυτόματη δημιουργία επαγγελματικού προσανατολισμού",
        "auto_spinner": "Δημιουργείται ο επαγγελματικός προσανατολισμός…",
        "auto_ok": "✓ Πέρασε τον έλεγχο πληρότητας. Δες το αποτέλεσμα παρακάτω.",
        "auto_error": "Η αυτόματη δημιουργία απέτυχε.",
        "auto_need_audit_caption": "Σε απλή παρουσίαση, το μοντέλο παράγει σε μία κλήση και τα δύο: το καθαρό παραδοτέο ΚΑΙ το εσωτερικό τεχνικό δελτίο, ξεχωριστά.",
        "auto_attempt_spinner": "Προσπάθεια {n}/{max}: δημιουργείται ο επαγγελματικός προσανατολισμός…",
        "auto_correcting_spinner": "Προσπάθεια {n}/{max}: το μοντέλο διορθώνει {count} πρόβλημα/προβλήματα που εντόπισε ο έλεγχος…",
        "auto_ok_after_retries": "✓ Πέρασε τον έλεγχο πληρότητας μετά από {n} προσπάθεια/ες αυτοδιόρθωσης. Δες το αποτέλεσμα παρακάτω.",
        "auto_exhausted": "Δοκιμάστηκαν {max} αυτόματες προσπάθειες αυτοδιόρθωσης, χωρίς επιτυχία. Χρειάζεται πλέον ανθρώπινος έλεγχος -- δες τις λεπτομέρειες παρακάτω.",
        "step4_title": "Βήμα 4 · Έλεγχος αποτελέσματος",
        "result_uploader": "Αποτέλεσμα επαγγελματικού προσανατολισμού (.docx)",
        "audit_uploader": "Εσωτερικό τεχνικό δελτίο ελέγχου (.docx) — δεν παραδίδεται στον πελάτη",
        "check_button": "Έλεγχος επαγγελματικού προσανατολισμού",
        "download_checked": "⬇️ Λήψη ελεγμένου επαγγελματικού προσανατολισμού",
        "details_expander": "Λεπτομέρειες",
        "ctx_name": "Όνομα",
        "ctx_presentation": "Τρόπος παρουσίασης",
    },
    "en": {
        "hero_title": "AstroCheck Career",
        "hero_subtitle": "Talent mapping and career orientation for a child, teen, or adult — from a single Astrodienst PDF.",
        "sidebar_header": "AstroCheck Career",
        "sidebar_privacy": "Data is processed only for this session and is not stored by the app.",
        "sidebar_reset": "🔄 Start over (clear everything)",
        "sidebar_reset_help": "Clears the chart and choices, so you can start clean with the next person.",
        "step1_title": "Step 1 · Upload the Astrodienst PDF",
        "pdf_uploader": "Astrodienst Data Sheet (.pdf)",
        "read_pdf": "Read PDF",
        "reading_spinner": "Reading the PDF…",
        "pdf_error": "Reading stopped safely — the PDF might not be the correct Astrodienst Data Sheet, or its layout differs.",
        "technical_detail": "Technical detail",
        "chart_loaded": "✓ Chart already checked in this session: **{name}**. Upload a new PDF only to replace it, or press “🔄 Start over” in the sidebar.",
        "step2_title": "Step 2 · Name",
        "name_label": "Name for the final document",
        "presentation_simple": "Simple & practical",
        "success_simple": "A 4–5 page text will be requested, in everyday language, with 4–6 talents, 4–5 career fields, and a brief final synthesis.",
        "no_extra_data": "No additional personal, psychological, school, or financial data is requested. This service only presents symbolic possibilities to explore, drawn from the checked chart.",
        "step3_title": "Step 3 · Download the command for ChatGPT/Claude",
        "download_command": "⬇️ Download the orientation command for ChatGPT/Claude",
        "upload_hint_simple": "Upload the Word file to ChatGPT or Claude. Ask for two files: the clean client document and the internal technical audit sheet.",
        "paste_expander": "Ready-to-paste message for ChatGPT/Claude",
        "auto_title": "#### 🤖 Alternative: automatic generation",
        "auto_caption": "Requires your own OpenAI API key (never stored). Skips the copy/paste steps to and from ChatGPT or Claude — the result is checked automatically by the same validator.",
        "auto_button": "Automatically generate the career orientation",
        "auto_spinner": "Generating the career orientation…",
        "auto_ok": "✓ Passed the completeness check. See the result below.",
        "auto_error": "Automatic generation failed.",
        "auto_need_audit_caption": "In simple presentation, the model produces both the clean deliverable AND the internal technical audit sheet in one call, kept separate.",
        "auto_attempt_spinner": "Attempt {n}/{max}: generating the career orientation…",
        "auto_correcting_spinner": "Attempt {n}/{max}: the model is fixing {count} issue(s) found by the check…",
        "auto_ok_after_retries": "✓ Passed the completeness check after {n} self-correction attempt(s). See the result below.",
        "auto_exhausted": "Tried {max} automatic self-correction attempts without success. Human review is now needed -- see the details below.",
        "step4_title": "Step 4 · Check the result",
        "result_uploader": "Career orientation result (.docx)",
        "audit_uploader": "Internal technical audit sheet (.docx) — not delivered to the client",
        "check_button": "Check career orientation result",
        "download_checked": "⬇️ Download the checked career orientation result",
        "details_expander": "Details",
        "ctx_name": "Name",
        "ctx_presentation": "Presentation mode",
    },
}

st.set_page_config(page_title="AstroCheck Career", page_icon="✦", layout="wide")
st.markdown("""<style>
.stApp{background:#f6f2fa}.block-container{max-width:1100px;padding-top:2rem}.hero{background:#3d2350;color:white;border-radius:22px;padding:30px 34px;margin-bottom:18px}.hero h1{margin:0 0 8px;font-family:Georgia;font-size:42px}.hero p{color:#e6d9f2}.ok{padding:14px 16px;background:#e5f2e7;border-left:5px solid #39704c;border-radius:8px}.warn{padding:14px 16px;background:#fff1dd;border-left:5px solid #b7791f;border-radius:8px}</style>""",unsafe_allow_html=True)

_, lang_col = st.columns([6, 1.6])
with lang_col:
    site_language = st.selectbox("Γλώσσα / Language", ["Ελληνικά", "English"], key="site_language")
lang = "el" if site_language == "Ελληνικά" else "en"
t = TR[lang]

st.markdown(f'<div class="hero"><h1>{t["hero_title"]}</h1><p>{t["hero_subtitle"]}</p></div>', unsafe_allow_html=True)

if 'chart' not in st.session_state: st.session_state.chart = None
if 'uploader_gen' not in st.session_state: st.session_state.uploader_gen = 0

with st.sidebar:
    st.header(t["sidebar_header"])
    st.caption(t["sidebar_privacy"])
    if st.button(t["sidebar_reset"], use_container_width=True, help=t["sidebar_reset_help"]):
        st.session_state.chart = None
        st.session_state.uploader_gen += 1
        reset_case_state()
        st.rerun()

chart = st.session_state.chart

if not chart:
    st.subheader(t["step1_title"])
    pdf = st.file_uploader(t["pdf_uploader"], type=['pdf'], key=f"pdf_{st.session_state.uploader_gen}")
    if pdf and st.button(t["read_pdf"], type="primary", use_container_width=True):
        with st.spinner(t["reading_spinner"]):
            ok, new_chart, err = handle_pdf_upload(pdf.getvalue(), pdf.name)
            if ok:
                st.rerun()
            else:
                st.error(t["pdf_error"])
                with st.expander(t["technical_detail"]): st.code(str(err))
    st.stop()

st.success(t["chart_loaded"].format(name=chart.name))

st.subheader(t["step2_title"])
name_override = st.text_input(t["name_label"], value=chart.name, key='name_override')

SERVICE_TITLE = "Ανάδειξη Ταλέντων και Διερεύνηση Επαγγελματικών Επιλογών"
OUTPUT_NAME = "AstroCheck_Anadeixi_Talenton.docx"
# Η "Αναλυτική με αστρολογική τεκμηρίωση" παρουσίαση αφαιρέθηκε -- στην πράξη
# χρησιμοποιείται πάντα η "Απλή και πρακτική", οπότε κλειδώνεται σταθερά εδώ.
presentation_label = t["presentation_simple"]
presentation = "Απλή και πρακτική"

st.success(t["success_simple"])
st.caption(t["no_extra_data"])

st.subheader(t["step3_title"])
context = {
    t["ctx_name"]: name_override or chart.name,
    t["ctx_presentation"]: presentation_label,
}

command_text = load_orientation_command()
orientation_source = build_orientation_source(chart)
style_example_text = load_unified_short_example()
orientation_doc = build_orientation_docx(
    name_override or chart.name, SERVICE_TITLE, context, command_text,
    orientation_source, style_example_text=style_example_text,
)
st.download_button(t["download_command"], orientation_doc, file_name=OUTPUT_NAME, type="primary", use_container_width=True)

# The single top language toggle now decides the deliverable's language too
# (instead of asking again in a separate control) — this was the strongest
# idea worth adopting from the reviewed prototype.
language_clause = (
    "Γράψε το καθαρό παραδοτέο του πελάτη (και, αν ζητηθεί, το εσωτερικό τεχνικό δελτίο) εξ ολοκλήρου στα Ελληνικά."
    if lang == "el" else
    "Write the clean client deliverable (and, if requested, the internal technical audit sheet) entirely in English."
)

st.caption(t["upload_hint_simple"])
if lang == "el":
    paste_message = """Ακολούθησε πιστά τη δεσμευτική εντολή που περιλαμβάνεται στο έγγραφο και χρησιμοποίησε αποκλειστικά τα ελεγμένα τεχνικά δεδομένα που περιέχει. Μην επινοήσεις προσωπικά, επαγγελματικά ή ψυχολογικά στοιχεία.

Η επιλεγμένη παρουσίαση είναι «Απλή και πρακτική». Παράδωσε δύο χωριστά, ολοκληρωμένα αρχεία Word:
1. Το καθαρό παραδοτέο του πελάτη, χωρίς πλανήτες, Οίκους, όψεις, orb ή κατηγορίες βαρύτητας.
2. Το εσωτερικό τεχνικό δελτίο ελέγχου, με την πλήρη τεκμηρίωση που απαιτεί η δεσμευτική εντολή. Το δεύτερο αρχείο δεν παραδίδεται στον πελάτη.

Εφάρμοσε υποχρεωτικά τον Κανόνα 0Γ («Ενοποιημένος κανόνας σύντομης και απλής έκδοσης»): το καθαρό παραδοτέο να είναι 4–5 σελίδες, σε καθημερινή γλώσσα, και χωρίς τους αναλυτικούς πίνακες, το εργασιακό περιβάλλον, τα επόμενα βήματα, το σχέδιο 8–12 εβδομάδων ή τις επαναλαμβανόμενες ενότητες της πλήρους έκδοσης. Το εσωτερικό τεχνικό δελτίο παραμένει αναλυτικό.

Το ανώνυμο πρότυπο σύντομης έκδοσης περιλαμβάνεται ήδη μέσα στο έγγραφο. Χρησιμοποίησέ το αποκλειστικά για τη δομή, το μήκος και την απλή γλώσσα· μην αντιγράψεις από αυτό περιεχόμενο ή συμπεράσματα.

Το καθαρό Word πρέπει να έχει λευκό φόντο και μαύρο κείμενο, όπως το πρότυπο. Μην χρησιμοποιήσεις highlight, χρωματιστό φόντο, σκιάσεις, έγχρωμα πλαίσια ή χρωματιστές λωρίδες. Η έμφαση να γίνεται μόνο με τίτλους, κουκκίδες και περιορισμένη έντονη γραφή.

""" + language_clause + """

Κάνε προσεκτικό αυτοέλεγχο πριν από την παράδοση. Ο πραγματικός validator θα εκτελεστεί στη συνέχεια μέσα στο AstroCheck Career."""
else:
    paste_message = """Follow the binding command included in the document precisely, and use only the checked technical data it contains. Do not invent personal, professional, or psychological details.

The chosen presentation is "Simple & practical". Deliver two separate, complete Word files:
1. The clean client deliverable, without planets, Houses, aspects, orb, or weight categories.
2. The internal technical audit sheet, with the full documentation the binding command requires. This second file is not delivered to the client.

Apply Rule 0Γ ("Unified simple-version rule") mandatorily: the clean deliverable must be 4–5 pages, in everyday language, and must NOT include the detailed tables, work environment, next steps, the 8–12 week plan, or the repeated sections of the full version. The internal technical sheet stays detailed.

The short-version anonymous template is already included inside the document. Use it only for structure, length, and plain language — do not copy content or conclusions from it.

The clean Word file must have a white background and black text, like the template. Do not use highlighting, colored backgrounds, shading, colored boxes, or colored bars. Emphasis should only come from headings, bullet points, and limited bold text.

""" + language_clause + """

Do a careful self-check before delivering. The real validator will run afterwards inside AstroCheck Career."""

# Ενισχυτικές οδηγίες (ισχύουν πάντα) -- προστέθηκαν μετά από πραγματικά
# περιστατικά όπου το μοντέλο έγραφε το όνομα με λατινικούς χαρακτήρες ή τον
# τίτλο ολόκληρο σε κεφαλαία. Χτίζονται σε ΞΕΧΩΡΙΣΤΗ μεταβλητή (όχι απευθείας
# μέσα στο paste_message) ώστε να περνούν ΚΑΙ στο κουμπί αυτόματης δημιουργίας
# (μέσω build_orientation_prompt), όχι μόνο στο κείμενο αντιγραφής για
# ChatGPT/Claude -- αλλιώς τα δύο μονοπάτια θα έδιναν διαφορετικό αποτέλεσμα.
reinforcement_instructions = (
    f"""Γράψε το όνομα «{name_override or chart.name}» ακριβώς όπως δόθηκε, με ελληνικούς χαρακτήρες -- μην το μεταγράψεις σε λατινικό αλφάβητο (π.χ. όχι "GAVRIELA"). Ο κύριος τίτλος του εγγράφου να είναι σε κανονική μορφή πεζών/κεφαλαίων (π.χ. «Ανάδειξη Ταλέντων και Διερεύνηση Επαγγελματικών Επιλογών»), όχι ολόκληρος σε κεφαλαία."""
    if lang == "el" else
    f"""Write the name "{name_override or chart.name}" exactly as given, in Greek characters -- do not transliterate it into the Latin alphabet (e.g. not "GAVRIELA"). The document's main title should use normal sentence/title case, not ALL CAPS."""
)

paste_message += "\n\n" + reinforcement_instructions
paste_message += "\n\n" + (CAREER_CONSISTENCY_RULE_EL if lang == "el" else CAREER_CONSISTENCY_RULE_EN)

with st.expander(t["paste_expander"], expanded=True):
    st.code(paste_message, language=None)

st.divider()
with st.container(border=True):
    st.markdown(t["auto_title"])
    st.caption(t["auto_caption"])
    career_api_key = st.text_input("OpenAI API key", type="password", key="career_api_key", placeholder="sk-...", label_visibility="collapsed")
    if st.button(t["auto_button"], type="primary", disabled=not career_api_key, use_container_width=True):
        with st.spinner(t["auto_spinner"]):
            try:
                auto_prompt = build_orientation_prompt(
                    context, command_text, orientation_source, style_example_text,
                    language_clause=language_clause, need_audit=True,
                    extra_instructions=reinforcement_instructions,
                )
                MAX_AUTO_ATTEMPTS = 3
                current_prompt = auto_prompt
                previous_raw = None
                check = None
                client_text = audit_text = None
                for attempt in range(1, MAX_AUTO_ATTEMPTS + 1):
                    if attempt == 1:
                        spinner_msg = t["auto_attempt_spinner"].format(n=attempt, max=MAX_AUTO_ATTEMPTS)
                    else:
                        spinner_msg = t["auto_correcting_spinner"].format(
                            n=attempt, max=MAX_AUTO_ATTEMPTS, count=len(check.details_lines())
                        )
                    with st.spinner(spinner_msg):
                        raw = generate_analysis(career_api_key, current_prompt)
                    client_text, audit_text = split_orientation_response(raw)
                    orientation_personal = {"Όνομα": name_override or chart.name}
                    check = validate_orientation(
                        chart, client_text, orientation_personal,
                        presentation_mode=presentation, audit_text=audit_text,
                        format_issues=[],
                    )
                    if check.ok or attempt == MAX_AUTO_ATTEMPTS:
                        break
                    # Ο επόμενος γύρος δεν ξαναγράφει τα πάντα από την αρχή --
                    # στέλνει πίσω στο μοντέλο ό,τι μόλις έγραψε, μαζί με τη
                    # συγκεκριμένη λίστα προβλημάτων που εντόπισε ο πραγματικός
                    # validator, και του ζητά να διορθώσει ΜΟΝΟ αυτά. Αυτό είναι
                    # που πλησιάζει τον στόχο "χωρίς ανθρώπινη παρέμβαση".
                    issues_list = "\n".join(f"- {line}" for line in check.details_lines())
                    correction_note = (
                        f"\n\nΗ ΠΡΟΗΓΟΥΜΕΝΗ ΣΟΥ ΑΠΑΝΤΗΣΗ (προς διόρθωση):\n{raw}\n\n"
                        f"Ο ΠΡΑΓΜΑΤΙΚΟΣ ΜΗΧΑΝΙΚΟΣ ΕΛΕΓΧΟΣ ΒΡΗΚΕ ΤΑ ΕΞΗΣ ΠΡΟΒΛΗΜΑΤΑ:\n{issues_list}\n\n"
                        "Ξαναγράψε ολόκληρη την απάντηση (με την ίδια ακριβώς μορφή/δείκτες όπως πριν), "
                        "διορθώνοντας ΜΟΝΟ αυτά τα συγκεκριμένα προβλήματα. Μην αλλάξεις τίποτα άλλο που "
                        "ήδη ήταν σωστό."
                        if lang == "el" else
                        f"\n\nYOUR PREVIOUS ANSWER (to be corrected):\n{raw}\n\n"
                        f"THE REAL ENGINEERING CHECK FOUND THESE ISSUES:\n{issues_list}\n\n"
                        "Rewrite the entire answer (using the exact same format/markers as before), "
                        "fixing ONLY these specific issues. Do not change anything else that was already correct."
                    )
                    current_prompt = auto_prompt + correction_note

                st.session_state.orientation_validation = check
                st.session_state.orientation_docx_bytes = build_orientation_client_docx(SERVICE_TITLE, name_override or chart.name, client_text)
                st.session_state.orientation_docx_name = OUTPUT_NAME
                st.session_state.orientation_audit_docx_bytes = audit_text.encode("utf-8") if audit_text else None
                if check.ok:
                    if attempt == 1:
                        st.success(t["auto_ok"])
                    else:
                        st.success(t["auto_ok_after_retries"].format(n=attempt))
                else:
                    st.error(t["auto_exhausted"].format(max=MAX_AUTO_ATTEMPTS))
            except Exception as e:
                st.error(t["auto_error"])
                with st.expander(t["technical_detail"]): st.code(str(e))
    st.caption(t["auto_need_audit_caption"])

st.divider()
st.subheader(t["step4_title"])
orientation_result = st.file_uploader(t["result_uploader"], type=['docx'], key=f"orientation_result_{st.session_state.uploader_gen}")
orientation_audit = st.file_uploader(t["audit_uploader"], type=['docx'], key=f"orientation_audit_{st.session_state.uploader_gen}")
ready_to_check = bool(orientation_result) and bool(orientation_audit)
if st.button(t["check_button"], disabled=not ready_to_check, use_container_width=True):
    result_bytes = orientation_result.getvalue()
    result_text = docx_text(result_bytes)
    result_format_issues = simple_docx_format_issues(result_bytes)
    audit_bytes = orientation_audit.getvalue() if orientation_audit else None
    audit_text = docx_text(audit_bytes) if audit_bytes else None
    orientation_personal = {"Όνομα": name_override or chart.name}
    check = validate_orientation(
        chart, result_text, orientation_personal,
        presentation_mode=presentation, audit_text=audit_text,
        format_issues=result_format_issues,
    )
    st.session_state.orientation_validation = check
    st.session_state.orientation_docx_bytes = result_bytes
    st.session_state.orientation_docx_name = orientation_result.name
    st.session_state.orientation_audit_docx_bytes = audit_bytes
check = st.session_state.get('orientation_validation')
if check:
    if check.ok:
        st.markdown(f'<div class="ok">{check.summary()}</div>', unsafe_allow_html=True)
        st.download_button(t["download_checked"], st.session_state.orientation_docx_bytes, file_name=st.session_state.orientation_docx_name, use_container_width=True)
    else:
        st.markdown(f'<div class="warn">⚠ {check.summary()}</div>', unsafe_allow_html=True)
        with st.expander(t["details_expander"], expanded=True):
            for line in check.details_lines(): st.write("•", line)
