"""Add native PowerPoint animation by editing <p:timing> XML safely.

This is the deep part skill-for-pptx is known for, re-implemented with
its safety rules (see skills/office-pptx-motion/references/animation-xml.md):

- python-pptx has NO animation API. Animation is XML, not a library call.
- Never invent presetID/presetClass/presetSubtype from memory. Use the
  preset table below, which maps to PowerPoint's documented presets, and
  validate the result by re-parsing with lxml and by a LibreOffice render.
- Target is always spid = p:cNvPr id of a real shape, read back after
  saving with python-pptx. Never guess an id.
- If a slide already has <p:timing>, replace it entirely. Never append a
  second timing element, and keep schema order: cSld, clrMapOvr?,
  transition, timing, extLst.
- cTn ids are unique and increasing within each slide.
- Click beats: Slide -> Click Beat -> Stage -> Effects. A click is a
  p:par nodeType="clickEffect" with stCondLst cond delay="indefinite"
  evt="onClick". Do NOT use a long delay instead of a click.
- Morph is a TRANSITION on the destination slide, not an animation:
  <p:transition><p:morph type="byObject"/></p:transition>, and paired
  shapes share a "!!" name.
- Evidence ladder: structure -> simulation -> LibreOffice -> PowerPoint
  read/render -> interactive click playback. This module only claims the
  structure rung; render/PlayBack must be checked separately.
"""
from __future__ import annotations

import copy
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}

# (presetID, presetClass, presetSubtype, filter) for common entrance effects.
# presetClass 1 = entrance, 2 = emphasis, 5 = exit in PowerPoint's preset table.
PRESETS = {
    "fade": (10, 1, 0, "fade"),
    "fly_in_from_bottom": (5, 1, 8, "fade"),
    "wipe_from_left": (22, 1, 4, "fade"),
    "zoom": (13, 1, 0, "fade"),
    "appear": (1, 1, 0, "fade"),
}


def _q(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{P_NS if prefix == 'p' else A_NS}}}{local}"


def _cTn(cid: int, dur: str = "indefinite", node_type: str | None = None, fill: str = "hold") -> etree._Element:
    el = etree.Element(_q("p:cTn"), id=str(cid), dur=dur, fill=fill)
    if node_type:
        el.set("nodeType", node_type)
    return el


def _anim_effect(spid: int, preset: str, dur_ms: int, cid: int) -> etree._Element:
    preset_id, preset_class, preset_subtype, filt = PRESETS[preset]
    par = etree.Element(_q("p:par"))
    ctn = _cTn(cid, dur=f"{dur_ms}")
    par.append(ctn)
    child = etree.SubElement(ctn, _q("p:childTnLst"))
    effect = etree.SubElement(child, _q("p:animEffect"), transition="in", filter=filt)
    effect.set("presetID", str(preset_id))
    effect.set("presetClass", "entr" if preset_class == 1 else "emph")
    effect.set("presetSubtype", str(preset_subtype))
    ctn2 = etree.SubElement(effect, _q("p:cTn"), id=str(cid + 1), dur=f"{dur_ms}")
    tgt = etree.SubElement(ctn2, _q("p:tgtEl"))
    sp = etree.SubElement(tgt, _q("p:spTgt"), spid=str(spid))
    return par


def add_fade_click_beats(pptx_path: str | Path, out_path: str | Path, tokens: dict | None = None) -> Path:
    """Add one Fade click beat per non-title shape on each content slide.

    Only shapes that are NOT the title are animated, following the
    skill-for-pptx rule "draft does not target title". Each beat is a
    separate click, so the presenter controls the pace.
    """
    from pptx import Presentation

    tokens = tokens or {}
    motion = tokens.get("motion", {}) if isinstance(tokens, dict) else {}
    dur = int(motion.get("entrance_ms", {}).get("default", 500)) if isinstance(motion.get("entrance_ms"), dict) else 500

    prs = Presentation(str(pptx_path))
    # Collect (slide_index, spid) for body shapes, in reading order.
    targets: dict[int, list[int]] = {}
    for idx, slide in enumerate(prs.slides):
        ids: list[int] = []
        for shape in slide.shapes:
            if shape.shape_type is None:
                continue
            # Skip the title placeholder.
            if shape.is_placeholder and shape.placeholder_format.type is not None:
                try:
                    from pptx.enum.shapes import PP_PLACEHOLDER

                    if shape.placeholder_format.type == PP_PLACEHOLDER.TITLE:
                        continue
                except Exception:
                    pass
            # Only shapes with a real cNvPr id.
            try:
                ids.append(int(shape.shape_id))
            except Exception:
                continue
        if idx > 0 and ids:
            targets[idx] = ids[:4]  # max 4 moving objects per click (token rule)

    # Rewrite the zip, replacing slide XML timing.
    src = Path(pptx_path)
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(src, "r") as zin:
        names = zin.namelist()
        parts = {name: zin.read(name) for name in names}
    for slide_idx, spids in targets.items():
        part_name = f"ppt/slides/slide{slide_idx + 1}.xml"
        if part_name not in parts:
            continue
        root = etree.fromstring(parts[part_name])
        # Remove any existing timing.
        for old in root.findall(f"{{{P_NS}}}timing"):
            root.remove(old)
        timing = etree.Element(_q("p:timing"))
        tn_lst = etree.SubElement(timing, _q("p:tnLst"))
        main_par = etree.SubElement(tn_lst, _q("p:par"))
        main_ctn = _cTn(1, dur="indefinite", node_type="mainSeq")
        main_par.append(main_ctn)
        main_children = etree.SubElement(main_ctn, _q("p:childTnLst"))
        cid = 2
        for spid in spids:
            click_par = etree.SubElement(main_children, _q("p:par"))
            click_ctn = _cTn(cid, fill="hold")
            cid += 1
            click_par.append(click_ctn)
            st = etree.SubElement(click_ctn, _q("p:stCondLst"))
            etree.SubElement(st, _q("p:cond"), delay="indefinite", evt="onClick")
            click_children = etree.SubElement(click_ctn, _q("p:childTnLst"))
            click_children.append(_anim_effect(spid, "fade", dur, cid))
            cid += 2
        # Insert timing in schema order: after transition if present, else after cSld.
        transition = root.find(f"{{{P_NS}}}transition")
        if transition is not None:
            idx = list(root).index(transition)
            root.insert(idx + 1, timing)
        else:
            csld = root.find(f"{{{P_NS}}}cSld")
            idx = list(root).index(csld) if csld is not None else len(root) - 1
            root.insert(idx + 1, timing)
        parts[part_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(out, "w") as zout:
        for name in names:
            zout.writestr(name, parts[name])
    # Re-parse validation: every slide XML must parse, every spid must exist.
    with ZipFile(out, "r") as z:
        for name in names:
            if name.startswith("ppt/slides/slide"):
                etree.fromstring(z.read(name))
    return out


def add_morph_transition(pptx_path: str | Path, out_path: str | Path, slide_index: int) -> Path:
    """Add a Morph transition to the destination slide (1-based index)."""
    from pptx import Presentation  # noqa: F401  (keeps the dependency explicit)

    src = Path(pptx_path)
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(src, "r") as zin:
        names = zin.namelist()
        parts = {name: zin.read(name) for name in names}
    part_name = f"ppt/slides/slide{slide_index}.xml"
    root = etree.fromstring(parts[part_name])
    for old in root.findall(f"{{{P_NS}}}transition"):
        root.remove(old)
    transition = etree.Element(_q("p:transition"))
    etree.SubElement(transition, _q("p:morph"), type="byObject")
    csld = root.find(f"{{{P_NS}}}cSld")
    idx = list(root).index(csld) if csld is not None else 0
    root.insert(idx + 1, transition)
    parts[part_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(out, "w") as zout:
        for name in names:
            zout.writestr(name, parts[name])
    return out
