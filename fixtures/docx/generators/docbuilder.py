"""Doc: a Package plus the bookkeeping the feature modules share (ids, comments, notes, links, images)."""
from ooxml import *  # noqa: F401,F403
from ooxml import (CT, NS_DECL, Package, anchor_image, esc, font_table_xml, inline_image, numbering_xml, p, png, r,
                   settings_xml, styles_xml, theme_xml)

PEOPLE = ["Alice Archer", "Bram Keller", "Chidi Okafor", "Dana Whitfield"]


def lcg(seed):
    """Deterministic pseudo-random ints (no `random` module, so output never depends on the Python build)."""
    state = seed & 0x7FFFFFFF or 1
    while True:
        state = (state * 1103515245 + 12345) & 0x7FFFFFFF
        yield state >> 8


_SUBJ = ["The harbour committee", "A junior surveyor", "Every lighthouse keeper", "The northern ferry service", "Our records office",
         "The tide almanac", "A visiting engineer", "The grain exchange", "Local fishing crews", "The bridge authority", "A careful archivist",
         "The winter inspection team", "Several market traders", "The quay-side bakery", "An experienced pilot", "The regional council"]
_VERB = ["measured", "recorded", "reviewed", "rebuilt", "mapped", "postponed", "catalogued", "negotiated", "inspected", "repainted",
         "reorganised", "compared", "approved", "documented", "relocated", "audited"]
_OBJ = ["the eastern jetty", "last season's tonnage figures", "a ledger of unpaid berth fees", "the lamp oil inventory", "three disputed boundary stones",
        "the fog-signal schedule", "a bundle of weathered charts", "the ferry timetable", "every outstanding repair order", "the old customs register",
        "two dozen mooring rings", "the sluice gate mechanism", "a draft of the annual report", "the dredging contract", "the winter storage sheds", "an index of local pilots"]
_TAIL = ["before the first frost", "with unusual care", "as the tide turned", "according to the standing orders", "during the long spring inspection",
         "without waiting for approval", "in the presence of two witnesses", "after several rounds of discussion", "while the market was closed",
         "using only the instruments on hand", "to the satisfaction of nobody in particular", "in time for the autumn meeting", "and filed the result in triplicate"]


def sentence(g):
    return f"{_SUBJ[next(g) % len(_SUBJ)]} {_VERB[next(g) % len(_VERB)]} {_OBJ[next(g) % len(_OBJ)]} {_TAIL[next(g) % len(_TAIL)]}."


def paragraph_text(seed, n=5):
    g = lcg(seed)
    return " ".join(sentence(g) for _ in range(n))


class Doc(Package):
    def __init__(self, title, creator="Spindle fixture generator", **kw):
        super().__init__(title, creator, **kw)
        self._ids = 1000
        self.comments = []  # (cid, author_idx, day, para_id, text, parent_para, done)
        self.fn, self.en = [], []
        self._img = 0
        self.settings_extra = ""
        self.settings_mid = ""
        self.extra_numbering = ""
        self.styles = None
        self.theme = None
        self.font_extra = ""
        self.fonts = None
        self.font_rels = []  # (rid, name, key) for embedded fonts
        self.hf_n = 0
        self.mutate = None  # optional hook(doc) run just before the package is written

    def nid(self):
        self._ids += 1
        return self._ids

    # ---- links / images
    def link(self, url, text, **kw):
        rid = self.rel("hyperlink", url, external=True)
        return f'<w:hyperlink r:id="{rid}" w:history="1">{r(text, style="Hyperlink", **kw)}</w:hyperlink>'

    def image_rid(self, shape="rect", w=160, h=100, fg=(30, 90, 200), bg=(235, 235, 235)):
        self._img += 1
        return self.media(f"image{self._img}.png", png(w, h, shape, fg, bg)), self._img

    def inline(self, shape="rect", w=160, h=100, descr="", crop=None, fg=(30, 90, 200)):
        rid, n = self.image_rid(shape, w, h, fg)
        return inline_image(rid, self.nid(), w, h, descr or f"{shape} image", crop)

    def floating(self, shape="circle", w=120, h=120, descr="", **kw):
        rid, n = self.image_rid(shape, w, h, fg=kw.pop("fg", (200, 70, 40)))
        return anchor_image(rid, self.nid(), w, h, descr or f"{shape} image", **kw)

    # ---- notes
    def footnote(self, text, extra=""):
        nid = len(self.fn) + 1
        self.fn.append((nid, self.note_para("footnote", text) + extra))
        return self.note_ref("footnote", nid)

    def endnote(self, text):
        nid = len(self.en) + 1
        self.en.append((nid, self.note_para("endnote", text)))
        return self.note_ref("endnote", nid)

    # ---- comments
    def comment(self, who, day, text, parent=None, done=False):
        cid = len(self.comments)
        para = f"{0x20000000 + cid:08X}"
        self.comments.append((cid, who % len(PEOPLE), day, para, text, parent, done))
        return cid, para

    def commented(self, cids, inner):
        s = "".join(f'<w:commentRangeStart w:id="{c}"/>' for c in cids)
        e = "".join(f'<w:commentRangeEnd w:id="{c}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{c}"/></w:r>' for c in cids)
        return s + inner + e

    # ---- revisions
    def ins(self, who, text, day=5, **kw):
        return f'<w:ins w:id="{self.nid()}" w:author="{PEOPLE[who]}" w:date="2026-01-{day:02d}T10:00:00Z">{r(text, **kw)}</w:ins>'

    def dele(self, who, text, day=5):
        return (f'<w:del w:id="{self.nid()}" w:author="{PEOPLE[who]}" w:date="2026-01-{day:02d}T10:00:00Z"><w:r>'
                f'<w:delText xml:space="preserve">{esc(text)}</w:delText></w:r></w:del>')

    # ---- finish
    def finish(self, body, final_sect, path):
        if self.fn:
            self.notes("footnote", self.fn)
        if self.en:
            self.notes("endnote", self.en)
        if self.comments:
            cs = []
            ex = []
            for cid, who, day, para, text, parent, done in self.comments:
                cs.append(f'<w:comment w:id="{cid}" w:author="{PEOPLE[who]}" w:date="2026-01-{day:02d}T14:00:00Z" w:initials="{"".join(w[0] for w in PEOPLE[who].split())}">'
                          f'<w:p w14:paraId="{para}" w14:textId="77777777"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>{r(text)}</w:p></w:comment>')
                pp = f' w15:paraIdParent="{parent}"' if parent else ""
                ex.append(f'<w15:commentEx w15:paraId="{para}"{pp} w15:done="{1 if done else 0}"/>')
            self.part("word/comments.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments {NS_DECL}>{"".join(cs)}</w:comments>', CT["comments"], "comments")
            self.part("word/commentsExtended.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx {NS_DECL}>{"".join(ex)}</w15:commentsEx>',
                      CT["commentsExtended"], "http://schemas.microsoft.com/office/2011/relationships/commentsExtended")
            people = "".join(f'<w15:person w15:author="{n}"><w15:presenceInfo w15:providerId="None" w15:userId="{n}"/></w15:person>' for n in PEOPLE)
            self.part("word/people.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:people {NS_DECL}>{people}</w15:people>',
                      "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml", "http://schemas.microsoft.com/office/2011/relationships/people")
        self.standard(
            styles=self.styles or styles_xml(),
            numbering=numbering_xml(self.extra_numbering),
            settings=settings_xml(self.settings_extra, footnote_pr=bool(self.fn or self.en), mid=self.settings_mid),
            theme=self.theme or theme_xml(),
            fonts=self.fonts or font_table_xml(extra=self.font_extra),
        )
        self.document(body, final_sect)
        if self.mutate:
            self.mutate(self)
        self.save(path)
