from .astrology import RULERS, angular_distance, axis_activation_note, degree_theory, movement_text, orb_to_text

def fmt(p):
    rx = " ανάδρομος" if p.retrograde else ""
    return f"{p.sign} {p.degree}°{p.minute:02d}′{p.second:02d}″{rx}"


ORIENTATION_RESPONSE_DELIMITER = "===TECHNICAL_AUDIT==="
ORIENTATION_CLIENT_MARKER = "===CLIENT_DELIVERABLE==="


CAREER_CONSISTENCY_RULE_EL = """ΥΠΟΧΡΕΩΤΙΚΟΣ ΚΑΝΟΝΑΣ ΣΥΝΕΠΕΙΑΣ ΤΑΛΕΝΤΩΝ ΚΑΙ ΕΠΑΓΓΕΛΜΑΤΩΝ
Πρώτα καθόρισε στο εσωτερικό τεχνικό δελτίο τα τεκμηριωμένα ταλέντα, έπειτα τους επαγγελματικούς τομείς που συνδέονται ρητά με τουλάχιστον δύο από αυτά, και τέλος τα επιτρεπόμενα ενδεικτικά επαγγέλματα κάθε τομέα. Το καθαρό παραδοτέο επιτρέπεται μόνο να απλοποιήσει αυτά τα ήδη εγκεκριμένα στοιχεία· απαγορεύεται να δημιουργήσει νέο ταλέντο, τομέα ή επάγγελμα κατά τη συγγραφή του.

Στο τεχνικό δελτίο πρόσθεσε υποχρεωτικά τις δύο αναγνωρίσιμες ενότητες «ΕΓΚΕΚΡΙΜΕΝΟΙ ΕΠΑΓΓΕΛΜΑΤΙΚΟΙ ΤΟΜΕΙΣ» και «ΕΓΚΕΚΡΙΜΕΝΑ ΕΠΑΓΓΕΛΜΑΤΑ», ώστε κάθε τομέας και κάθε επάγγελμα του καθαρού παραδοτέου να εμφανίζεται επίσης εκεί.

Απαγορεύεται να δημιουργήσεις τομέα υγείας ή επάγγελμα υγείας μόνο από γενικές έννοιες όπως «φροντίδα», «ενσυναίσθηση», «θεραπευτική ποιότητα», «ακρίβεια», «βελτίωση» ή «βοήθεια». Επαγγέλματα όπως ιατρός/γιατρός, νοσηλευτής, φαρμακοποιός, ψυχολόγος, εργοθεραπευτής ή διατροφολόγος επιτρέπονται μόνο όταν το τεχνικό δελτίο περιέχει χωριστή ενότητα «ΡΗΤΗ ΤΕΚΜΗΡΙΩΣΗ ΤΟΜΕΑ ΥΓΕΙΑΣ» και καταγράφει τουλάχιστον δύο διακριτούς λειτουργικούς δείκτες που στηρίζουν ειδικά τον κλάδο της υγείας. Αν υπάρχει μόνο ένδειξη ακρίβειας, πρακτικής βελτίωσης ή βοήθειας, χρησιμοποίησε αυτές τις ουδέτερες έννοιες χωρίς να τις μετατρέψεις σε επάγγελμα υγείας."""

CAREER_CONSISTENCY_RULE_EN = """MANDATORY TALENT-TO-CAREER CONSISTENCY RULE
First define the documented talents in the internal technical audit, then career fields explicitly linked to at least two of those talents, and finally the allowed example jobs for each field. The clean client deliverable may only simplify those already-approved items; it must not create a new talent, field, or job while drafting.

The technical audit must include the exact Greek section headings «ΕΓΚΕΚΡΙΜΕΝΟΙ ΕΠΑΓΓΕΛΜΑΤΙΚΟΙ ΤΟΜΕΙΣ» and «ΕΓΚΕΚΡΙΜΕΝΑ ΕΠΑΓΓΕΛΜΑΤΑ», and every career field and job used in the clean deliverable must also appear there.

Do not infer a health field or health profession merely from generic concepts such as care, empathy, healing quality, precision, improvement, or help. Doctor/physician, nurse, pharmacist, psychologist, occupational therapist, or dietitian are allowed only if the audit includes the separate heading «ΡΗΤΗ ΤΕΚΜΗΡΙΩΣΗ ΤΟΜΕΑ ΥΓΕΙΑΣ» and at least two distinct functional indicators specifically supporting the health field."""


def build_orientation_prompt(context, command_text, orientation_source,
                              style_example_text="", language_clause="",
                              need_audit=False, extra_instructions=""):
    """Καθαρό text prompt για απευθείας κλήση στο μοντέλο (χωρίς να χρειάζεται
    το ενδιάμεσο βήμα «κατέβασε Word -> επικόλλησε σε ChatGPT/Claude -> ανέβασε
    το αποτέλεσμα»). Χρησιμοποιεί ακριβώς τα ίδια υλικά με το
    docx_builder.build_orientation_docx (δεσμευτική εντολή, δηλωμένο πλαίσιο,
    ελεγμένα δεδομένα, ανώνυμο πρότυπο ύφους) -- μόνο η μορφή αλλάζει, από
    .docx σε ένα ενιαίο string κατάλληλο για το generator.generate_analysis.
    """
    context_lines = "\n".join(f"- {k}: {v}" for k, v in context.items() if str(v).strip())
    parts = [
        "ΔΕΣΜΕΥΤΙΚΗ ΑΡΧΗ\nΠρόκειται για χωριστή προαιρετική υπηρεσία. Χρησιμοποίησε αποκλειστικά την παρακάτω δεσμευτική εντολή, το δηλωμένο πλαίσιο και την ήδη ελεγμένη τεχνική ανάλυση του ίδιου ατόμου. Μην χρησιμοποιήσεις μνήμη ή προηγούμενες συνομιλίες και μην επινοήσεις προσωπικά δεδομένα.",
        f"ΔΗΛΩΜΕΝΟ ΠΛΑΙΣΙΟ ΥΠΗΡΕΣΙΑΣ\n{context_lines}",
        f"ΔΕΣΜΕΥΤΙΚΗ ΕΝΤΟΛΗ ΕΠΑΓΓΕΛΜΑΤΙΚΟΥ ΠΡΟΣΑΝΑΤΟΛΙΣΜΟΥ\n{command_text}",
        CAREER_CONSISTENCY_RULE_EN if "English" in language_clause else CAREER_CONSISTENCY_RULE_EL,
    ]
    if style_example_text.strip():
        parts.append(
            "ΑΝΩΝΥΜΟ ΠΡΟΤΥΠΟ ΣΥΝΤΟΜΗΣ ΕΚΔΟΣΗΣ — ΜΟΝΟ ΓΙΑ ΔΟΜΗ ΚΑΙ ΥΦΟΣ\n"
            "Χρησιμοποίησε το ακόλουθο υλικό μόνο ως πρότυπο μήκους, διάταξης και "
            "απλής γλώσσας. Απαγορεύεται να αντιγράψεις από αυτό ταλέντα, "
            "επαγγελματικούς τομείς, επαγγέλματα ή συμπεράσματα. Το περιεχόμενο "
            "για το νέο πρόσωπο πρέπει να προκύπτει αποκλειστικά από τα ελεγμένα "
            "δεδομένα που ακολουθούν.\n" + style_example_text
        )
    parts.append(f"ΕΛΕΓΜΕΝΗ ΤΕΧΝΙΚΗ ΑΝΑΛΥΣΗ — ΜΟΝΑΔΙΚΗ ΑΣΤΡΟΛΟΓΙΚΗ ΠΗΓΗ\n{orientation_source}")
    if language_clause.strip():
        parts.append(language_clause)
    if extra_instructions.strip():
        parts.append(extra_instructions)
    if need_audit:
        parts.append(
            "ΜΟΡΦΗ ΑΠΑΝΤΗΣΗΣ (ΥΠΟΧΡΕΩΤΙΚΗ)\n"
            "Επίστρεψε ΑΚΡΙΒΩΣ δύο ενότητες, με αυτή τη σειρά, καθεμία σε δική "
            "της γραμμή με τον ακριβή τίτλο:\n"
            f"{ORIENTATION_CLIENT_MARKER}\n"
            "(εδώ το καθαρό παραδοτέο του πελάτη, σύμφωνα με τους παραπάνω κανόνες)\n"
            f"{ORIENTATION_RESPONSE_DELIMITER}\n"
            "(εδώ το πλήρες εσωτερικό τεχνικό δελτίο ελέγχου, με όλη την τεκμηρίωση "
            "που απαιτεί η δεσμευτική εντολή)\n"
            "Μην προσθέσεις κανένα άλλο κείμενο πριν από τον πρώτο τίτλο ή μετά "
            "το τέλος του τεχνικού δελτίου."
        )
    return "\n\n".join(parts)


def split_orientation_response(raw_text):
    """Σπάει την απάντηση του μοντέλου σε (καθαρό_κείμενο_πελάτη, τεχνικό_δελτίο).

    Το δεύτερο στοιχείο είναι None όταν δεν ζητήθηκε τεχνικό δελτίο (αναλυτική
    παρουσίαση) ή όταν το μοντέλο δεν ακολούθησε τη ζητούμενη μορφή -- σε αυτή
    την περίπτωση επιστρέφεται όλο το κείμενο ως παραδοτέο πελάτη, ώστε ο
    validator να αποφασίσει (και πιθανόν να απορρίψει) αντί να χαθεί σιωπηλά
    περιεχόμενο.
    """
    text = raw_text.strip()
    if ORIENTATION_RESPONSE_DELIMITER in text:
        client_part, _, audit_part = text.partition(ORIENTATION_RESPONSE_DELIMITER)
        client_part = client_part.replace(ORIENTATION_CLIENT_MARKER, "").strip()
        return client_part, audit_part.strip()
    return text.replace(ORIENTATION_CLIENT_MARKER, "").strip(), None



def build_orientation_source(chart):
    """Καθαρή τεχνική πηγή για την προαιρετική υπηρεσία προσανατολισμού.

    Δεν περνά την αφηγηματική ανάλυση των 12 Οίκων, η οποία μπορεί να
    περιέχει προσωπικό πλαίσιο. Έτσι το μοντέλο βλέπει μόνο τα ελεγμένα
    αστρολογικά δεδομένα που χρειάζεται η ξεχωριστή υπηρεσία.
    """
    points = "\n".join(
        f"- {p.name}: {fmt(p)} | Οίκος {p.house or '—'} | {movement_text(p)}"
        for p in chart.points
    )
    cusps = "\n".join(
        f"- {number}ος Οίκος: {fmt(c)}"
        for number, c in enumerate(chart.cusps, start=1)
    )
    aspects = "\n".join(
        f"- {a.first}–{a.second} | {a.aspect} | orb {a.orb_text} | {a.weight} | πηγή: {a.source}"
        for a in chart.aspects
    ) or "- Δεν αναγνωρίστηκαν όψεις. Σταμάτησε και ζήτησε τεχνικό έλεγχο."
    return f"""ΕΛΕΓΜΕΝΑ ΑΣΤΡΟΛΟΓΙΚΑ ΔΕΔΟΜΕΝΑ

ΠΛΑΝΗΤΕΣ ΚΑΙ ΣΗΜΕΙΑ
{points}

ΑΚΜΕΣ ΟΙΚΩΝ
{cusps}

ΠΑΡΑΡΤΗΜΑ ΕΠΙΒΕΒΑΙΩΜΕΝΩΝ ΟΨΕΩΝ — ΜΟΝΑΔΙΚΗ ΠΗΓΗ ORB ΚΑΙ ΒΑΡΥΤΗΤΑΣ
{aspects}
"""

def ruler_block(chart, cusp):
    modern, traditional = RULERS[cusp.sign]
    def describe(name, label):
        ruler = next((p for p in chart.points if p.name == name), None)
        if not ruler:
            return f"{label}: {name} — δεν εντοπίστηκε στα δεδομένα."
        aspects = [a for a in chart.aspects if ruler.name in (a.first,a.second)]
        aspect_text = "; ".join(
            f"{a.first}–{a.second} {a.aspect}, orb {a.orb_text}, {a.weight}"
            for a in aspects
        ) or "καμία επιβεβαιωμένη όψη"
        return (f"{label}: {name}. Θέση: {fmt(ruler)}, {ruler.house}ος Οίκος. "
                f"Όψεις: {aspect_text}.")

    blocks = [describe(modern, "Σύγχρονος/κύριος κυβερνήτης")]
    if traditional and traditional != modern:
        blocks.append(describe(traditional, "Παραδοσιακός κυβερνήτης"))
    return "\n".join(blocks)

def house_section(chart, number):
    cusp=chart.cusps[number-1]; nxt=chart.cusps[number%12]
    planets=[p for p in chart.points if p.house==number and p.kind in ("planet","node")]
    modern_ruler, traditional_ruler = RULERS[cusp.sign]
    involved={p.name for p in planets}
    involved.add(modern_ruler)
    if traditional_ruler:
        involved.add(traditional_ruler)
    # Ο Οίκος 1 και ο Οίκος 7 μοιράζονται τον άξονα Ωροσκόπου/Δύσης· ο Οίκος 4
    # και ο Οίκος 10 τον άξονα Πυθμένα Ουρανού/Μεσουρανήματος. Καμία όψη προς
    # αυτά τα σημεία δεν πρέπει να εμφανίζεται ΜΟΝΟ στον Οίκο όπου τυχαίνει να
    # βρίσκεται ο άλλος πλανήτης -- πρέπει να εμφανίζεται και εδώ, στον Οίκο
    # που το ίδιο το σημείο ορίζει. (Έτσι διορθώνεται η περίπτωση Χείρωνα: η
    # σύνοδός του με τον Ωροσκόπο πρέπει να αναλυθεί και στον 1ο και στον 7ο
    # Οίκο, όχι μόνο στον 12ο όπου τυχαίνει να κατοικεί ο Χείρωνας.)
    if number in (1, 7):
        involved.add("Ωροσκόπος")
    if number in (4, 10):
        involved.add("Μεσουράνημα")
    aspects=[a for a in chart.aspects if a.first in involved or a.second in involved]
    hard=[a for a in aspects if a.aspect in ("Τετράγωνο","Αντίθεση")]
    other=[a for a in aspects if a not in hard]
    near=[]
    for p in planets:
        distance=(nxt.absolute-p.absolute)%360
        if distance <= 5:
            near.append(f"{p.name}: τεχνικά στον {number}ο, απόσταση {orb_to_text(distance)} από την επόμενη ακμή· μπορεί συμπληρωματικά να επηρεάζει τον {number%12+1}ο.")
    plist="\n".join(f"- {p.name}: {fmt(p)} · {degree_theory(p)}" for p in planets) or "- Κανένας πλανήτης."
    def alines(items):
        lines = []
        for a in items:
            note = axis_activation_note(a)
            suffix = f" [{note}]" if note else ""
            lines.append(f"- {a.first}–{a.second}: {a.aspect}, orb {a.orb_text}, {a.weight}, πηγή: {a.source}{suffix}")
        return "\n".join(lines) or "- Καμία."
    return f"""{number}ος ΟΙΚΟΣ
Ακμή και έκταση: {fmt(cusp)} → {fmt(nxt)}.
Πλανήτες/σημεία:
{plist}
Κύριος κυβερνήτης:
{ruler_block(chart,cusp)}
Πλανήτες κοντά σε επόμενη ακμή:
{chr(10).join('- '+x for x in near) if near else '- Κανένας σε απόσταση έως 5°.'}
ΥΠΟΧΡΕΩΤΙΚΑ τετράγωνα και αντιθέσεις πλανητών και κυβερνητών:
{alines(hard)}
Άλλες επιβεβαιωμένες κύριες όψεις:
{alines(other)}

Στη συγγραφή αυτού του Οίκου συμπερίλαβε υποχρεωτικά: «Σύνθεση με τον υπόλοιπο χάρτη», «Η πρακτική εφαρμογή» και πλαίσιο σύνοψης με Βασική δύναμη, Βασική πρόκληση, Κυβερνήτη και Τελικό συμπέρασμα."""

def build_master_prompt(chart, personal, language, instructions_text, style_text,
                 instructions_name="Ενσωματωμένες οδηγίες v5.3",
                        style_name="Ενσωματωμένος καθαρός οδηγός ύφους"):
    # Το διορθωμένο όνομα (name_override στο tab3) είναι η πηγή αλήθειας όταν
    # δίνεται· το chart.name (ωμή εξαγωγή από το PDF) είναι μόνο fallback.
    # Πριν, η επικεφαλίδα έγραφε πάντα {chart.name} ενώ το ΠΡΟΣΩΠΙΚΟ ΠΛΑΙΣΙΟ
    # παρακάτω έδειχνε το διορθωμένο -- τα δύο μπορούσαν να διαφωνούν.
    display_name = (personal.get("Όνομα") or "").strip() or chart.name
    personal_text="\n".join(f"- {k}: {v}" for k,v in personal.items() if v) or "- Δεν δόθηκαν ακόμη προσωπικές πληροφορίες. Ζήτησε τες πριν από τη συγγραφή."
    aspect_appendix="\n".join(f"- {a.first}–{a.second}: {a.aspect}, orb {a.orb_text}, {a.weight}, {a.source}" for a in chart.aspects) or "- Δεν αναγνωρίστηκαν όψεις. Σταμάτησε και ζήτησε έλεγχο."
    houses="\n\n".join(house_section(chart,i) for i in range(1,13))
    return f"""ΔΕΣΜΕΥΤΙΚΗ ΕΝΤΟΛΗ
Χρησιμοποίησε το «{instructions_name}» ως δεσμευτική προδιαγραφή και το «{style_name}» αποκλειστικά ως πρότυπο ύφους, βάθους, δομής και μορφοποίησης. Όλα τα αστρολογικά δεδομένα προέρχονται αποκλειστικά από το νέο PDF και τον παρακάτω ελεγμένο πίνακα. Μην μεταφέρεις δεδομένα ή προσωπικές πληροφορίες από το πρότυπο.

ΑΠΑΡΑΒΑΤΟ ΟΡΙΟ ΠΗΓΩΝ
Ο οδηγός ύφους δεν αποτελεί πηγή δεδομένων ή ερμηνευτικών συμπερασμάτων. Σε περίπτωση σύγκρουσης υπερισχύουν οι οδηγίες v5 και τα ελεγμένα δεδομένα του νέου χάρτη. Απαγορεύεται επίσης να χρησιμοποιήσεις μνήμη, προηγούμενες συνομιλίες ή εξωτερική γνώση για προσωπικά γεγονότα του ατόμου.

Γλώσσα τελικού Word: {language}.
Όνομα: {display_name}
Ημερομηνία: {chart.date} · Ώρα: {chart.time} · Τόπος: {chart.place}
Σύστημα Οίκων: {chart.house_system}

ΠΡΟΣΩΠΙΚΟ ΠΛΑΙΣΙΟ
{personal_text}

ΥΠΟΧΡΕΩΤΙΚΟΙ ΚΑΝΟΝΕΣ
1. Αν λείπει ή αμφισβητείται προσωπικό στοιχείο, ζήτησε επιβεβαίωση. Μην επινοήσεις εμπειρίες.
1Α. Αν στο ΠΡΟΣΩΠΙΚΟ ΠΛΑΙΣΙΟ υπάρχει μόνο όνομα, απαγορεύεται να αναφέρεις συγκεκριμένο επάγγελμα, σπουδές, οικογενειακή κατάσταση, αριθμό ή ηλικία παιδιών, έργα, στόχους, συνήθειες ή εμπειρίες, ακόμη και αν τα γνωρίζεις από προηγούμενη συνομιλία.
2. Μην παραλείψεις κανένα τετράγωνο ή αντίθεση του Astrodienst, ακόμη και όταν είναι πολύ πλατύ/δευτερεύον.
3. Η ίδια όψη πρέπει να έχει παντού το ίδιο orb και την ίδια κατηγορία.
3Α. Αντέγραψε ακριβώς από τα ελεγμένα δεδομένα και τα τρία πεδία κάθε όψης: ΤΥΠΟ ΟΨΗΣ, ORB και ΚΑΤΗΓΟΡΙΑ ΒΑΡΥΤΗΤΑΣ. Απαγορεύεται να μετατρέψεις αντίθεση σε τετράγωνο ή να αλλάξεις «Πλατιά αλλά έγκυρη» σε «Πολύ πλατιά/δευτερεύουσα», ακόμη και σε συνθετική ή τελική ενότητα.
4. Χρησιμοποίησε προσεκτική, πιθανική, μη μοιρολατρική και μη διαγνωστική γλώσσα.
5. Η Θεωρία των Μοιρών είναι μόνο συμπληρωματική και ακολουθεί ζώδιο, Οίκο, όψεις και κυβερνήτη.
6. Κάθε Οίκος να είναι συνεχές συνθετικό κείμενο και όχι ασύνδετη λίστα.
7. Όταν μια όψη αφορά τον Ωροσκόπο ή το Μεσουράνημα και έχει σημείωση σε αγκύλες [...], ενσωμάτωσε τη σημασία της -- ότι ενεργοποιείται ταυτόχρονα το απέναντι σημείο του άξονα (Δύση/Πυθμένας Ουρανού). ΜΗΝ τη γράψεις ως δεύτερη, ανεξάρτητη όψη με δικό της orb· είναι η ίδια όψη, από την άλλη άκρη του άξονα.

ΕΛΕΓΜΕΝΑ ΔΕΔΟΜΕΝΑ ΑΝΑ ΟΙΚΟ
{houses}

ΤΕΛΙΚΕΣ ΥΠΟΧΡΕΩΤΙΚΕΣ ΕΝΟΤΗΤΕΣ
- Τελική συνθετική εικόνα.
- Συμβολική κατεύθυνση εξέλιξης: Βόρειος/Νότιος Δεσμός, κυβερνήτης Βόρειου Δεσμού, επιβεβαιωμένες όψεις και πρακτική έκφραση.
- Προτάσεις προσωπικής ανάπτυξης και υποστήριξης, πρακτικές και μη διαγνωστικές.
- Παράρτημα επιβεβαιωμένων όψεων.
- Δεύτερος μαθηματικός, γλωσσικός και οπτικός έλεγχος πριν από την παράδοση.
- Στον δεύτερο έλεγχο σύγκρινε λέξη προς λέξη κάθε αναφερόμενη όψη με το Παράρτημα: ίδιο ζεύγος, ίδιος τύπος, ίδιο orb και ίδια κατηγορία βαρύτητας. Αν υπάρχει διαφορά, διόρθωσέ την πριν παραδώσεις.

ΠΑΡΑΡΤΗΜΑ ΕΠΙΒΕΒΑΙΩΜΕΝΩΝ ΟΨΕΩΝ — ΜΟΝΑΔΙΚΗ ΠΗΓΗ
{aspect_appendix}

ΜΟΡΦΟΠΟΙΗΣΗ
Παράδωσε καλαίσθητο Word με τίτλο, υπότιτλο, μεθοδολογία, βασικά δεδομένα, αρίθμηση σελίδων και κάθε Οίκο κατά προτίμηση σε νέα σελίδα. Κράτησε κάθε πλαίσιο σύνοψης ολόκληρο στην ίδια σελίδα.

================ ΠΛΗΡΕΙΣ ΔΕΣΜΕΥΤΙΚΕΣ ΟΔΗΓΙΕΣ v5 ================
{instructions_text}
================ ΤΕΛΟΣ ΟΔΗΓΙΩΝ v5 ================

================ ΚΑΘΑΡΟΣ ΟΔΗΓΟΣ ΥΦΟΥΣ — ΟΧΙ ΠΗΓΗ ΔΕΔΟΜΕΝΩΝ ================
{style_text}
================ ΤΕΛΟΣ ΟΔΗΓΟΥ ΥΦΟΥΣ ================"""
