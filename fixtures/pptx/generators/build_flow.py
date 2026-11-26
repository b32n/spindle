"""Slide flow and review: transitions, animations, notes, comments, hidden slides, sections, and right-to-left and CJK text."""
import pptxbuilder as X
from dml import body_ph, bullet_char, bullet_num, para, rgb, run, scheme, solid, sp, spacing, textbox, title

MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"


def _slide(d, text, layout=3):
    s = d.slide(layout)
    s.add(title(2, text))
    return s


def _shape(sid, name, x, y, color="accent1"):
    return sp(sid, name, x, y, 2200000, 900000, fill=solid(scheme(color)), paras=para(name, algn="ctr"))


# -- transitions and animations ---------------------------------------------------------------------------------------------

class _Ids:
    def __init__(self):
        self.n = 2

    def __call__(self):
        self.n += 1
        return self.n


def _target(spid, para_index=None):
    if para_index is None:
        return f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
    return f'<p:tgtEl><p:spTgt spid="{spid}"><p:txEl><p:pRg st="{para_index}" end="{para_index}"/></p:txEl></p:spTgt></p:tgtEl>'


def _set_visibility(nid, spid, value, delay=0, para_index=None):
    return (f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst></p:cTn>{_target(spid, para_index)}'
            f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="{value}"/></p:to></p:set>')


def effect(nid, kind, spid, para_index=None):
    """One click effect: entrance (fade), exit (fade), emphasis (grow) or a motion path."""
    cls, preset, inner = {
        "entrance": ("entr", 10, lambda: _set_visibility(nid, spid, "visible", para_index=para_index)
                     + f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="500"/>{_target(spid, para_index)}</p:cBhvr></p:animEffect>'),
        "exit": ("exit", 10, lambda: f'<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="500"/>{_target(spid)}</p:cBhvr></p:animEffect>'
                 + _set_visibility(nid, spid, "hidden", delay=499)),
        "emphasis": ("emph", 6, lambda: f'<p:animScale><p:cBhvr><p:cTn id="{nid()}" dur="500" fill="hold"/>{_target(spid)}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>'),
        "path": ("path", 0, lambda: f'<p:animMotion origin="layout" path="M 0 0 L 0.25 0.1 E" pathEditMode="relative" rAng="0" ptsTypes="AA"><p:cBhvr><p:cTn id="{nid()}" dur="2000" fill="hold"/>{_target(spid)}'
                 '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr><p:rCtr x="0.125" y="0.05"/></p:animMotion>'),
    }[kind]
    outer, mid, node = nid(), nid(), nid()
    return (f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{mid}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{node}" presetID="{preset}" presetClass="{cls}" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{inner()}'
            "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>")


def timing(effects, builds=""):
    """A p:timing tree; `effects` are (kind, shape id[, paragraph]) tuples played in order, one click each."""
    nid = _Ids()
    root, seq = 1, 2
    body = "".join(effect(nid, *e) for e in effects)
    return (f'<p:timing><p:tnLst><p:par><p:cTn id="{root}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="{seq}" dur="indefinite" nodeType="mainSeq">'
            f'<p:childTnLst>{body}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
            f"<p:bldLst>{builds}</p:bldLst></p:timing>")


def transitions_animations(path):
    d = X.Deck(title="Transitions and animations")
    s = _slide(d, "Fade transition")
    s.transition = '<p:transition spd="med"><p:fade/></p:transition>'
    s = _slide(d, "Push transition, advance after 5 seconds")
    s.transition = '<p:transition spd="slow" advClick="0" advTm="5000"><p:push dir="u"/></p:transition>'
    s = _slide(d, "Wipe transition")
    s.transition = '<p:transition spd="fast"><p:wipe dir="r"/></p:transition>'
    s = _slide(d, "PowerPoint 2010 vortex transition with a fallback")
    s.transition = (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:p14="{P14}" Requires="p14"><p:transition spd="slow" p14:dur="2000"><p14:vortex dir="r"/></p:transition></mc:Choice>'
                    '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
    s = _slide(d, "Entrance, exit, emphasis and motion path")
    for n, kind in enumerate(("entrance", "exit", "emphasis", "path")):
        s.add(_shape(3 + n, kind.title(), 700000 + n * 2800000, 1700000, ("accent1", "accent2", "accent3", "accent4")[n]))
    s.timing = timing([("entrance", 3), ("exit", 4), ("emphasis", 5), ("path", 6)], builds="".join(f'<p:bldP spid="{3 + n}" grpId="0" animBg="1"/>' for n in range(4)))
    s = _slide(d, "Build a list paragraph by paragraph", 1)
    s.add(body_ph(3, "".join(para(f"Point {n}") for n in range(1, 4))))
    s.timing = timing([("entrance", 3, n) for n in range(3)], builds='<p:bldP spid="3" grpId="0" build="p"/>')
    d.write(path)


# -- notes, comments, hidden slides, sections ----------------------------------------------------------------------------------

def notes(path):
    d = X.Deck(title="Notes", notes_master=True)
    s = _slide(d, "A slide with plain notes")
    s.notes = para("Say hello, then introduce the topic.") + para("Pause for questions.")
    s = _slide(d, "A slide with rich notes")
    s.notes = (para(run("Key point: ", b=1) + run("keep it short", i=1) + run(" and ") + run("underline the number", u="sng") + run(" in red", fill=solid(rgb("C00000"))))
               + para("First step", marL=342900, indent=-342900, ppr_children=bullet_char())
               + para("Nested step", lvl=1, marL=742950, indent=-285750, ppr_children=bullet_char("&#8211;"))
               + para("Numbered", marL=342900, indent=-342900, ppr_children=bullet_num())
               + para("Spaced paragraph", ppr_children=spacing(before=600, after=600)))
    _slide(d, "A slide without notes")
    d.write(path)


def comments_legacy(path):
    d = X.Deck(title="Legacy comments")
    d.comment_authors = [(0, "Reviewer One", "R1"), (1, "Reviewer Two", "R2")]
    _slide(d, "A slide with two comments")
    _slide(d, "A slide without comments")
    _slide(d, "A slide with one comment")
    d.slide_comments = {0: [(0, 1200, 900, "Please check this title."), (1, 3000, 1500, "Agreed, shorten it.")], 2: [(1, 800, 600, "Add a source here.")]}
    d.write(path)


def comments_modern(path):
    d = X.Deck(title="Modern comments")
    s1 = _slide(d, "A slide with a threaded comment")
    _slide(d, "A slide without comments")
    authors = (f'{X.HEAD}<p188:authorLst xmlns:a="{X.A}" xmlns:r="{X.R}" xmlns:p188="{P188}">'
               '<p188:author id="{00000000-0000-4000-8000-0000000000A1}" name="Reviewer One" initials="R1" userId="reviewer.one" providerId="None"/>'
               '<p188:author id="{00000000-0000-4000-8000-0000000000A2}" name="Reviewer Two" initials="R2" userId="reviewer.two" providerId="None"/></p188:authorLst>')
    mk = lambda text: f'<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p188:txBody>'
    comment = (f'{X.HEAD}<p188:cmLst xmlns:a="{X.A}" xmlns:r="{X.R}" xmlns:p188="{P188}"><p188:cm id="{{00000000-0000-4000-8000-0000000000C1}}" authorId="{{00000000-0000-4000-8000-0000000000A1}}" created="2026-01-15T09:00:00.000">'
               '<pc:sldMkLst xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"><pc:docMk/><pc:sldMk cId="1000" sldId="256"/></pc:sldMkLst>'
               f'{mk("Is this the final title?")}<p188:replyLst><p188:reply id="{{00000000-0000-4000-8000-0000000000C2}}" authorId="{{00000000-0000-4000-8000-0000000000A2}}" created="2026-01-15T09:00:00.000">'
               f'{mk("Yes, it is.")}</p188:reply></p188:replyLst></p188:cm></p188:cmLst>')
    s1.part("ppt/comments/modernComment_100_3E8.xml", comment, "application/vnd.ms-powerpoint.comments+xml", "http://schemas.microsoft.com/office/2018/10/relationships/comments")
    d.extra["ppt/authors.xml"] = (authors.encode(), "application/vnd.ms-powerpoint.authors+xml")
    d.pres_rels_extra.append(("http://schemas.microsoft.com/office/2018/10/relationships/authors", "authors.xml"))
    d.write(path)


def hidden_sections(path):
    d = X.Deck(title="Hidden slides and sections")
    for n, text in enumerate(("Introduction", "Hidden backup slide", "Body", "More body", "Wrap-up"), 1):
        s = _slide(d, text)
        s.hidden = n == 2
    sections = [("Intro", (256, 257)), ("Body", (258, 259)), ("Wrap-up", (260,))]
    body = "".join(f'<p14:section name="{name}" id="{{00000000-0000-4000-8000-00000000005{n}}}"><p14:sldIdLst>{"".join(f"<p14:sldId id={chr(34)}{i}{chr(34)}/>" for i in ids)}</p14:sldIdLst></p14:section>'
                   for n, (name, ids) in enumerate(sections, 1))
    d.pres_extra["ext"] = f'<p:extLst><p:ext uri="{{521415D9-36F7-43E2-AB2F-B90AF26B5E84}}"><p14:sectionLst xmlns:p14="{P14}">{body}</p14:sectionLst></p:ext></p:extLst>'
    d.pres_extra["custShows"] = ('<p:custShowLst><p:custShow name="Short version" id="0"><p:sldLst><p:sld r:id="@slide:1@"/><p:sld r:id="@slide:5@"/></p:sldLst></p:custShow>'
                                 '<p:custShow name="Without backup" id="1"><p:sldLst><p:sld r:id="@slide:1@"/><p:sld r:id="@slide:3@"/><p:sld r:id="@slide:4@"/><p:sld r:id="@slide:5@"/></p:sldLst></p:custShow></p:custShowLst>')
    d.write(path)


def rtl_cjk(path):
    d = X.Deck(title="Right-to-left and CJK text")
    s = _slide(d, "Right-to-left text")
    he = para(run("שלום עולם, זו שקופית לבדיקה", lang="he-IL", cs="Arial"), algn="r", rtl=1)
    ar = para(run("مرحبا بالعالم، هذه شريحة للاختبار", lang="ar-SA", cs="Arial"), algn="r", rtl=1)
    mixed = para(run("שלום ", lang="he-IL") + run("Spindle", lang="en-US") + run(" עולם", lang="he-IL"), algn="r", rtl=1)
    s.add(sp(3, "Hebrew and Arabic", 700000, 1500000, 8000000, 2200000, paras=he + ar + mixed, bpr='<a:bodyPr wrap="square" rtlCol="1"/>', txbox=True, fill="<a:noFill/>"))
    s = _slide(d, "Chinese, Japanese and Korean")
    cjk = (para(run("你好，世界。这是一张测试幻灯片。", lang="zh-CN", ea="SimSun")) + para(run("こんにちは世界。これはテスト用のスライドです。", lang="ja-JP", ea="MS Gothic"))
           + para(run("안녕하세요 세계. 이것은 테스트 슬라이드입니다.", lang="ko-KR", ea="Malgun Gothic")) + para(run("你好，世界", lang="zh-TW", ea="PMingLiU")))
    s.add(sp(3, "CJK text", 700000, 1500000, 8000000, 2200000, paras=cjk, txbox=True, fill="<a:noFill/>"))
    s.add(sp(4, "Vertical Japanese", 9000000, 1500000, 1000000, 3200000, paras=para(run("縦書きのテキスト", lang="ja-JP", ea="MS Gothic")), bpr='<a:bodyPr vert="eaVert" wrap="square" rtlCol="0"/>', txbox=True, fill="<a:noFill/>"))
    d.write(path)
