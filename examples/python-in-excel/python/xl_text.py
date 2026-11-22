"""Text with Python's standard library: fuzzy matching against a list, similarity, accents and
names. No package needed."""
import difflib
import re
import unicodedata
from asfpy import xl_func


@xl_func(name='PY.FUZZY')
def fuzzy_match(text: str, choices: list, cutoff: float = 0.6):
    """The entry of choices closest to text (typos, case and spacing aside), or empty below cutoff."""
    keys = {c.strip().lower(): c for c in choices if isinstance(c, str) and c.strip()}
    found = difflib.get_close_matches(text.strip().lower(), keys, n=1, cutoff=cutoff)
    return keys[found[0]] if found else ''


@xl_func(name='PY.SIMILARITY')
def similarity(a: str, b: str):
    """How alike two texts are, from 0 to 1, case aside."""
    return round(difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio(), 4)


@xl_func(name='PY.NOACCENTS')
def strip_accents(text: str):
    """The text without accents: José Müller becomes Jose Muller."""
    return ''.join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))


@xl_func(name='PY.PROPERNAME')
def proper_name(text: str):
    """A person's name tidied: spaces collapsed, each part capitalized, with Mc, O', de, van."""
    small = {'de', 'del', 'la', 'van', 'von', 'da', 'di'}
    parts = re.sub(r'\s+', ' ', text.strip()).split(' ')
    out = []
    for i, p in enumerate(parts):
        w = p.lower()
        if i and w in small:
            out.append(w)
        elif w.startswith('mc') and len(w) > 2:
            out.append('Mc' + w[2:].capitalize())
        else:
            out.append('-'.join("'".join(s.capitalize() for s in h.split("'")) for h in w.split('-')))
    return ' '.join(out)
