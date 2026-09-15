from core.reference_loader import _filter_orientation_sections


_SAMPLE = """0. Υποχρεωτική επιλογή λειτουργίας
Γενικός κανόνας που ισχύει πάντα.
0Α. Υποχρεωτική μορφή παρουσίασης κάθε ταλέντου
Πίνακες, μόνο για την αναλυτική έκδοση.
0Γ. Ειδικός κανόνας σύντομης και απλής έκδοσης για ενήλικα
Κανόνες μόνο για ενήλικα.
0Δ. Ειδικός κανόνας σύντομης και απλής έκδοσης για παιδί/έφηβο
Κανόνες μόνο για παιδί/έφηβο.
1. Εξέτασε τους πέντε προσωπικούς πλανήτες
Μεθοδολογία που ισχύει και στα δύο modes.
10Α. Πλαίσιο εκπαιδευτικού συστήματος
Πλήρης αναλυτική εκδοχή, υπερισχύεται από το 0Γ/0Δ.
16. Υποχρεωτική δομή τελικού κειμένου
Δομή μόνο για την αναλυτική έκδοση.
19. Υπερισχύων κανόνας χωρίς συλλογή πρόσθετων δεδομένων
Γενικός κανόνας που ισχύει πάντα.
"""


def test_child_service_keeps_child_block_and_shared_rules():
    text = _filter_orientation_sections(_SAMPLE, "Παιδί/έφηβος")
    assert "Κανόνες μόνο για παιδί/έφηβο." in text
    assert "Μεθοδολογία που ισχύει και στα δύο modes." in text
    assert "Γενικός κανόνας που ισχύει πάντα." in text


def test_child_service_drops_adult_block():
    text = _filter_orientation_sections(_SAMPLE, "Παιδί/έφηβος")
    assert "Κανόνες μόνο για ενήλικα." not in text


def test_child_service_drops_analytical_only_sections():
    text = _filter_orientation_sections(_SAMPLE, "Παιδί/έφηβος")
    assert "Πίνακες, μόνο για την αναλυτική έκδοση." not in text
    assert "Πλήρης αναλυτική εκδοχή, υπερισχύεται από το 0Γ/0Δ." not in text
    assert "Δομή μόνο για την αναλυτική έκδοση." not in text


def test_adult_service_keeps_adult_block_and_drops_child_block():
    text = _filter_orientation_sections(_SAMPLE, "Ενήλικας σε αλλαγή επαγγελματικής πορείας")
    assert "Κανόνες μόνο για ενήλικα." in text
    assert "Κανόνες μόνο για παιδί/έφηβο." not in text


def test_filtering_meaningfully_shrinks_the_document():
    filtered = _filter_orientation_sections(_SAMPLE, "Παιδί/έφηβος")
    assert len(filtered) < len(_SAMPLE) * 0.7
