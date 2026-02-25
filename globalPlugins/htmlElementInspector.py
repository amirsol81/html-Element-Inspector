# -*- coding: utf-8 -*-
"""NVDA HTML Element Inspector.

A JAWS-inspired, single-shortcut inspector (NVDA+Shift+F1) that reports IA2
attributes and a DOM-like ancestry chain, optimized for readability.

Key goals:
- JAWS-like tag blocks: "Tag X has N parameters" with key=value lines
- Always show effective href where possible (links, images inside links, and #document URL)
- Promote nested nodes to canonical elements (e.g., img->a, placeholder->contenteditable host)
- Readable, stable output with blank line before each Tag block

Changelog:
- 0.1.43: Report aria-describedby as description= and include describedby id.
"""

import addonHandler
import api
import ui
import browseMode
import controlTypes
import globalPluginHandler
import html
from scriptHandler import script

# ===== DEBUG MODE (temporary, always active) =====
import logHandler
DEBUG_MODE = False
def _dbg(msg):
    if not DEBUG_MODE:
        return
    try:
        logHandler.log.warning("[WebElementInspector DEBUG] " + str(msg))
    except Exception:
        pass
    try:
        logHandler.log.debug("[WebElementInspector DEBUG] " + str(msg))
    except Exception:
        pass
# ===== END DEBUG MODE =====




def _safe_int_hex(v):
    try:
        return format(int(v), "X")
    except Exception:
        return None

def _dbg_obj(o, label):
    """Log high-signal properties for an NVDAObject without huge dumps."""
    if not DEBUG_MODE:
        return
    if o is None:
        _dbg(f"{label}: <None>")
        return
    try:
        _dbg(f"{label}: {o!r}")
    except Exception:
        _dbg(f"{label}: <unrepr>")
    # Tag and IA2 attrs (keys only + a few important values)
    try:
        ia2 = getattr(o, "IA2Attributes", None) or {}
    except Exception:
        ia2 = {}
    try:
        tag = ia2.get("tag") or _tag(o)
    except Exception:
        tag = None
    try:
        _dbg(f"{label} tag: {tag}")
    except Exception:
        pass
    try:
        keys = sorted(list(ia2.keys()))
        # Keep it small: show keys + selected important ones.
        _dbg(f"{label} IA2 keys({len(keys)}): {keys[:20]}" + (" ..." if len(keys) > 20 else ""))
    except Exception:
        pass

    for k in ("id", "class", "name", "name-from", "description-from", "xml-roles", "role", "aria-describedby", "describedby", "describedBy", "describedby-text", "html-input-name", "text-input-type", "text-model"):
        try:
            if k in ia2 and ia2.get(k) not in (None, ""):
                _dbg(f"{label} IA2 {k}={ia2.get(k)}")
        except Exception:
            pass

    # NVDA computed properties
    for attr in ("name", "value", "description", "role", "states", "isEnabled", "location"):
        try:
            v = getattr(o, attr, None)
        except Exception:
            v = None
        if v is None:
            continue
        try:
            if attr == "states":
                # Don't spam: show up to 12 state names
                sv = list(v) if isinstance(v, (set, list, tuple)) else [v]
                _dbg(f"{label} states({len(sv)}): {sv[:12]}" + (" ..." if len(sv) > 12 else ""))
            else:
                _dbg(f"{label} {attr}={v}")
        except Exception:
            pass

    # Raw MSAA/IA2 roles (COM) — can be unstable; log failures explicitly.
    ia = getattr(o, "IAccessibleObject", None)
    if ia:
        try:
            r = ia.accRole(0)
            hx = _safe_int_hex(r)
            if hx is not None:
                _dbg(f"{label} MSAA accRole hex={hx}")
        except Exception as e:
            _dbg(f"{label} MSAA accRole FAILED: {e}")
        try:
            import comInterfaces
            IA2 = getattr(comInterfaces, "IAccessible2", None)
            if IA2:
                ia2if = ia.QueryInterface(IA2)
                rr = getattr(ia2if, "role", None)
                hx = _safe_int_hex(rr)
                if hx is not None:
                    _dbg(f"{label} IA2 role hex={hx}")
        except Exception as e:
            _dbg(f"{label} IA2 role FAILED: {e}")

addonHandler.initTranslation()


def _safe(v):
    try:
        return "" if v is None else str(v)
    except Exception:
        return ""


def _ia2_attrs(obj):
    out = {}
    try:
        ia2 = getattr(obj, "IA2Attributes", None)
        if isinstance(ia2, dict):
            for k, v in ia2.items():
                ks = _safe(k).strip()
                if ks:
                    ksl = ks.lower()
                if ksl in ("aria-describedby", "describedby"):
                    ks = "describedby"
                elif ksl in ("aria-checked", "checked", "ariachecked"):
                    ks = "checked"
                elif ksl in ("aria-checked", "checked", "ariachecked"):
                    ks = "checked"
                out[ks] = _safe(v).strip()
        elif isinstance(ia2, str) and ia2.strip():
            parts = [p.strip() for p in ia2.split(";") if p.strip()]
            for p in parts:
                if ":" in p:
                    k, v = p.split(":", 1)
                    ks = k.strip()
                    ksl = ks.lower()
                    if ksl in ("aria-describedby", "describedby"):
                        ks = "describedby"
                    out[ks] = v.strip()
    except Exception:
        pass

    # Normalize common ARIA keys for richer, JAWS-like reports (minimal + safe).
    try:
        if "orientation" not in out and "aria-orientation" in out:
            out["orientation"] = out.get("aria-orientation", "").strip()
        if "valuetext" not in out and "aria-valuetext" in out:
            out["valuetext"] = out.get("aria-valuetext", "").strip()
        # Angular: normalize formControlName -> formcontrolname (keep minimal).
        for k in list(out.keys()):
            if k.lower() == "formcontrolname" and k != "formcontrolname":
                out["formcontrolname"] = out.pop(k)
        if "value" not in out:
            if out.get("aria-valuenow", "").strip():
                out["value"] = out.get("aria-valuenow", "").strip()
            elif out.get("aria-valuetext", "").strip():
                out["value"] = out.get("aria-valuetext", "").strip()
            else:
                v = getattr(obj, "value", None)
                vs = _safe(v).strip()
                if vs:
                    out["value"] = vs


        # Link URL refinement (very conservative): if an <a> exposes its URL via NVDA's value,
        # prefer reporting it as href to match HTML semantics. Only applies when href is absent.
        try:
            if out.get("tag", "").strip().lower() == "a":
                href = _safe(out.get("href", "")).strip()
                v = _safe(out.get("value", "")).strip()
                if (not href) and (v.startswith("http://") or v.startswith("https://")):
                    out["href"] = v
                    out.pop("value", None)
        except Exception:
            pass
    except Exception:
        pass


    return out


def _tag(obj):
    try:
        return _ia2_attrs(obj).get("tag", "").strip().lower()
    except Exception:
        return ""



def _prefer_interactive_container(obj):
    """Prefer a nearby interactive container (e.g. button-like role) over decorative children."""
    try:
        cur = obj
        steps = 0
        while cur and steps < 7:
            a = _ia2_attrs(cur)
            t = _safe(a.get("tag", "")).strip().lower()
            roleLower = _safe(a.get("role", "")).strip().lower()
            xmlRoleLower = _safe(a.get("xml-roles", "")).strip().lower()
            haspopup = _safe(a.get("haspopup", "")).strip().lower()
            tabindex = _safe(a.get("tabindex", "")).strip()
            fsff = _safe(a.get("fsFormField", "")).strip().lower()

            if t == "button":
                return cur

            # Common button-like patterns on modern web apps (e.g., Google):
            # role/button + tabindex=0, or haspopup signals menu button.
            if (roleLower == "button" or xmlRoleLower == "button") and tabindex == "0":
                return cur
            if haspopup in ("menu", "listbox", "dialog", "tree", "grid"):
                return cur
            if fsff == "true" and (roleLower or xmlRoleLower):
                return cur

            # Some controls may expose role=button via NVDA role enum.
            try:
                Role = getattr(controlTypes, "Role", None)
                if Role and getattr(cur, "role", None) == getattr(Role, "BUTTON", None):
                    return cur
            except Exception:
                pass

            cur = getattr(cur, "parent", None)
            steps += 1
    except Exception:
        pass
    return obj


def _get_candidate_object():
    """Return the most relevant object under the user.

    Prefer the object at the *browse-mode caret* when available, because
    api.getNavigatorObject() can become stale after opening/closing dialogs
    (e.g. browseableMessage), which may cause us to only see #document.
    """
    # 1) Browse mode caret object (most reliable in virtual buffers)
    try:
        focus = api.getFocusObject()
        ti = getattr(focus, "treeInterceptor", None)
        if ti and isinstance(ti, browseMode.BrowseModeTreeInterceptor):
            try:
                import textInfos
                tiInfo = ti.makeTextInfo(textInfos.POSITION_CARET)
                o = getattr(tiInfo, "NVDAObjectAtStart", None)
                if o:
                    return o
            except Exception:
                pass
    except Exception:
        pass

    # 2) Navigator object (review/browse cursor)
    try:
        nav = api.getNavigatorObject()
        if nav:
            return nav
    except Exception:
        pass

    # 3) Focus object
    try:
        return api.getFocusObject()
    except Exception:
        return None


def _states_set(obj):
    try:
        st = getattr(obj, "states", None)
        return set(st) if st else set()
    except Exception:
        return set()


def _state_in(obj, st):
    try:
        s = getattr(obj, "states", None)
        return (st in s) if s else False
    except Exception:
        return False


def _has_state_name(obj, needle):
    """Best-effort state membership test by string name."""
    try:
        n = needle.lower()
        for s in _states_set(obj):
            if n in _safe(s).lower():
                return True
    except Exception:
        pass
    return False


def _iter_children(root, max_nodes=120):
    """Best-effort BFS over NVDAObject children, capped for safety."""
    if not root:
        return
    seen = set()
    queue = [root]
    yielded = 0
    while queue and yielded < max_nodes:
        cur = queue.pop(0)
        oid = id(cur)
        if oid in seen:
            continue
        seen.add(oid)
        yielded += 1
        yield cur
        try:
            child = getattr(cur, "firstChild", None)
        except Exception:
            child = None
        steps = 0
        while child and steps < 50:
            queue.append(child)
            steps += 1
            try:
                child = getattr(child, "next", None)
            except Exception:
                child = None


def _dom_chain_with_tags(obj, max_depth=40):
    """Return [obj, parent, ..., #document] but only nodes that expose IA2 tag."""
    chain = []
    cur = obj
    for _ in range(max_depth):
        if not cur:
            break
        attrs = _ia2_attrs(cur)
        if attrs.get("tag"):
            chain.append(cur)
            if attrs.get("tag", "").strip().lower() == "#document":
                break
        try:
            cur = getattr(cur, "parent", None)
        except Exception:
            cur = None
    return chain


def _find_nearest_tag(chain, tag_name):
    tag_name = (tag_name or "").lower()
    for o in chain:
        if _tag(o) == tag_name:
            return o
    return None


def _is_contenteditable_host(o):
    a = _ia2_attrs(o)
    if _safe(a.get("contenteditable")).lower() in ("true", "1"):
        return True
    if "multiline" in a or "tabindex" in a or a.get("id", "") == "prompt-textarea":
        return _tag(o) in ("div", "textarea", "section")
    return False


# -----------------------------
# URL / href extraction (best-effort)
# -----------------------------

def _try_acc_value_url(obj):
    """Try MSAA accValue, which sometimes contains a URL for links/docs."""
    try:
        ia = getattr(obj, "IAccessibleObject", None)
        if ia:
            v = ia.accValue(0)
            v = _safe(v).strip()
            if v.startswith("http://") or v.startswith("https://"):
                return v
    except Exception:
        pass
    return ""


def _try_tree_interceptor_url(obj):
    """Try to get document URL from the treeInterceptor when available."""
    try:
        ti = getattr(obj, "treeInterceptor", None)
        if ti:
            for attr in ("documentURL", "URL", "url", "documentConstantIdentifier"):
                v = getattr(ti, attr, None)
                v = _safe(v).strip()
                if v.startswith("http://") or v.startswith("https://"):
                    return v
    except Exception:
        pass
    return ""


def _try_appmodule_url(obj):
    """Try appModule helpers (names vary)."""
    try:
        am = getattr(obj, "appModule", None)
        if not am:
            return ""
        for fn_name in ("getBrowserURL", "getCurrentURL", "getCurrentDocumentURL", "getDocumentURL"):
            fn = getattr(am, fn_name, None)
            if callable(fn):
                v = _safe(fn()).strip()
                if v.startswith("http://") or v.startswith("https://"):
                    return v
    except Exception:
        pass
    return ""


def _document_url(base_obj, chain):
    doc = _find_nearest_tag(chain, "#document")
    if doc:
        href = _ia2_attrs(doc).get("href", "").strip()
        if href:
            return href
        v = _try_acc_value_url(doc)
        if v:
            return v

    for getter in (_try_tree_interceptor_url, _try_appmodule_url, _try_acc_value_url):
        v = getter(base_obj)
        if v:
            return v

    try:
        focus = api.getFocusObject()
    except Exception:
        focus = None
    if focus:
        for getter in (_try_tree_interceptor_url, _try_appmodule_url, _try_acc_value_url):
            v = getter(focus)
            if v:
                return v

    return ""


def _effective_href(obj, chain, base_obj):
    """Return the most meaningful href for this element (JAWS-like).

    Rules:
    - <a>: always show its own href if discoverable
    - <img>/<svg> inside <a>: inherit href from nearest <a>
    - #document: show document URL as href
    - Other tags: do NOT inherit href from surrounding links (avoid noisy reports)
    """
    t = _tag(obj)

    # #document URL
    if t == "#document":
        a = _ia2_attrs(obj)
        href = a.get("href", "").strip()
        if href:
            return href
        du = _document_url(base_obj, chain)
        return du or ""

    # Direct href (prefer IA2)
    a = _ia2_attrs(obj)
    href = a.get("href", "").strip()
    if href:
        return href

    # MSAA value sometimes holds URL for links
    if t == "a":
        v = _try_acc_value_url(obj)
        if v:
            return v
        # Some backends expose URL only on document/focus
        return ""

    # Only inherit href for nested media inside a link
    if t in ("img", "svg"):
        link = _find_nearest_tag(chain, "a")
        if link:
            a2 = _ia2_attrs(link)
            href2 = a2.get("href", "").strip()
            if href2:
                return href2
            v2 = _try_acc_value_url(link)
            if v2:
                return v2

    return ""


# -----------------------------
# Canonical element selection

# -----------------------------

def _promote_canonical(obj, chain):
    """Promote nested nodes to a more meaningful, JAWS-like 'canonical' element."""
    if not obj or not chain:
        return _prefer_interactive_container(obj), None

    cur_tag = _tag(obj)

    if cur_tag in ("img", "svg"):
        link = _find_nearest_tag(chain, "a")
        if link:
            return link, obj

    if cur_tag in ("p", "span", "br"):
        for o in chain:
            if _is_contenteditable_host(o):
                return _prefer_interactive_container(o), obj

    return _prefer_interactive_container(obj), None


# -----------------------------
# Formatting (JAWS-like, readable)
# -----------------------------

def _ordered_params(tag_name, attrs):
    attrs = attrs or {}
    preferred = [
        "tag", "id", "class",
        "role", "xml-roles",
        "orientation",
        "MSAA Role",
        "IA2 Role",
        "href", "src",
        "type", "text-input-type",
        "name", "html-input-name",
        "accessible-name", "accessible-name-from",
        "formcontrolname",
        "value", "valuetext",
        "required", "multiline",
        "contenteditable", "tabindex",
        "maxlength", "autocomplete",
        "haspopup", "expanded",
        "selected",
        "checkable", "checked",
        "pressed",
        "label", "title",
        "describedby", "description",
        "description-from", "labelledby",
        "name-from", "explicit-name", "explicit-name-from",
        "level", "posinset", "setsize",
        "colspan", "rowspan", "table-cell-index",
        "readonly", "fsFormField",
        "display", "layout-guess", "text-align", "text-model",
    ]
    seen = set()
    keys = []
    for k in preferred:
        if k in attrs and k not in seen:
            keys.append(k)
            seen.add(k)
    for k in sorted(attrs.keys(), key=lambda s: s.lower()):
        if k not in seen:
            keys.append(k)
            seen.add(k)
    return keys


def _infer_form_attrs(obj, attrs):
    out = dict(attrs or {})

    # Normalize checked from IA2 attrs when present (JAWS-like true/false).
    if "checked" in out:
        v = _safe(out.get("checked", "")).strip().lower()
        if v in ("1", "true", "yes", "on", "mixed"):
            out["checked"] = "true"
        elif v in ("0", "false", "no", "off"):
            out["checked"] = "false"


    # Tooltip/title should not be reported as label. Keep JAWS-like: title=... + description-from=tooltip.
    try:
        src = _safe(out.get("description-from", "")).strip().lower()
        if src == "tooltip" and "label" in out and "title" not in out:
            out["title"] = out.pop("label")
    except Exception:
        pass
    t = (out.get("tag") or _tag(obj)).lower()

    # Never mark non-controls as form fields.
    if t in ("#document", "body", "html"):
        out.pop("fsFormField", None)

    # Infer basic form-field marker similar to JAWS (best-effort).
    # JAWS often reports fsFormField=true for interactive controls.
    if "fsFormField" not in out:
        try:
            xmlr = _safe(out.get("xml-roles", "")).lower()
            r = _safe(out.get("role", "")).lower()
        except Exception:
            xmlr = ""
            r = ""
        interactive_tags = {"input", "textarea", "select", "button"}
        interactive_roles = {
            "button", "checkbox", "radio", "combobox", "listbox", "textbox", "searchbox",
            "slider", "spinbutton", "menuitem", "option", "switch", "tab", "treeitem",
        }
        if (t in interactive_tags) or (r in interactive_roles) or (xmlr in interactive_roles):
            out["fsFormField"] = "true"

    if t in ("input", "textarea") and "type" not in out:
        if out.get("text-input-type"):
            out["type"] = out.get("text-input-type")

    if t in ("input", "textarea", "select") and "name" not in out:
        hin = out.get("html-input-name", "").strip()
        if hin:
            out["name"] = hin

    if "multiline" not in out:
        if t == "textarea":
            out["multiline"] = "true"
        elif t == "input":
            out["multiline"] = "false"

    if "required" not in out:
        if _state_in(obj, controlTypes.State.REQUIRED) or _has_state_name(obj, "required"):
            out["required"] = "true"

    if "label" not in out:
        # Prefer an explicit label/name when the backend reports that the
        # accessible name comes from an attribute (common on modern web apps).
        try:
            nm = _safe(getattr(obj, "name", "")).strip()
        except Exception:
            nm = ""
        if nm and _safe(out.get("name-from", "")).lower() == "attribute":
            out["label"] = nm
        else:
            try:
                desc = _safe(getattr(obj, "description", "")).strip()
            except Exception:
                desc = ""
            if desc:
                df = _safe(out.get("description-from", "")).lower()
                if df == "aria-describedby":
                    out["description"] = desc
                elif df == "tooltip":
                    out["title"] = desc
                else:
                    out["label"] = desc

    return out


def _infer_expanded_for_combobox(obj, chain, attrs):
    out = dict(attrs or {})
    xml_roles = _safe(out.get("xml-roles", "")).lower()
    role = _safe(out.get("role", "")).lower()
    haspopup = _safe(out.get("haspopup", "")).lower()

    is_combo = ("combobox" in xml_roles) or (role == "combobox") or (haspopup == "listbox")
    if not is_combo:
        return out

    # If we already have expanded from states mapping, keep it
    if "expanded" in out:
        return out

    # Google-style hint: container class includes 'emcav' when suggestions are open.
    for o in chain:
        a = _ia2_attrs(o)
        cls = _safe(a.get("class", "")).lower()
        if "a8sbwf" in cls and "emcav" in cls:
            out["expanded"] = "true"
            return out

    # Choose a broader container to search for a listbox.
    # Prefer the known Google container (A8SBwf) when present; otherwise first ancestor div.
    start = None
    for o in chain:
        a = _ia2_attrs(o)
        if _tag(o) == "div" and "a8sbwf" in _safe(a.get("class", "")).lower():
            start = o
            break
    if start is None:
        for o in chain:
            if _tag(o) == "div":
                start = o
                break
    if start is None:
        start = obj

    found_listbox = False
    for node in _iter_children(start, max_nodes=260):
        a = _ia2_attrs(node)
        xr = _safe(a.get("xml-roles", "")).lower()
        r = _safe(a.get("role", "")).lower()
        if xr == "listbox" or r == "listbox":
            found_listbox = True
            break

    out["expanded"] = "true" if found_listbox else "false"
    return out


def _augment_attrs_for_readability(obj, chain, attrs, base_obj):
    out = dict(attrs or {})

    t = _tag(obj)
    if t and "tag" not in out:
        out["tag"] = t

    # Accessible name (the final name NVDA speaks) + its likely source.
    # This is useful for evaluators who see NVDA announce something that is hard
    # to locate in the DOM.
    if "accessible-name" not in out:
        try:
            out["accessible-name"] = _safe(getattr(obj, "name", "")).strip()
        except Exception:
            out["accessible-name"] = ""

    # Where the accessible name likely comes from (best-effort refinement).
    try:
        ia2n = getattr(obj, "IA2Attributes", None) or {}
    except Exception:
        ia2n = {}
    nf = ia2n.get("name-from") or out.get("name-from") or ""
    lb = ia2n.get("label") or out.get("label") or ""

    computedFrom = _safe(nf).strip()

    # Conservative refinement: only specialize when we have direct evidence.
    # NVDA/IA2 often only reports name-from=attribute without exposing which attribute.
    if computedFrom == "related-element":
        # If labelledby is exposed, call out aria-labelledby explicitly.
        if (ia2n.get("labelledby") or out.get("labelledby") or ia2n.get("aria-labelledby") or out.get("aria-labelledby")):
            computedFrom = "aria-labelledby"
    elif computedFrom == "attribute":
        # Only upgrade to aria-label if aria-label is directly exposed.
        if (ia2n.get("aria-label") or out.get("aria-label") or ia2n.get("aria_label") or out.get("aria_label")):
            computedFrom = "aria-label"
        else:
            # Keep generic attribute (avoid guessing title/content/etc.).
            computedFrom = "attribute"
    elif computedFrom in ("contents", "content"):
        computedFrom = "content"
    elif computedFrom in ("label", "label-for", "labelfor"):
        computedFrom = "label"
    elif computedFrom in ("alt", "title", "aria-label", "aria-labelledby"):
        # keep as-is
        pass

    existingFrom = out.get("accessible-name-from")
    # Only overwrite generic/empty cases; keep any more specific source already set.
    if computedFrom:
        if (not existingFrom) or (existingFrom == "attribute" and computedFrom.startswith("aria-")):
            out["accessible-name-from"] = computedFrom

    # If NVDA flags an explicit-name, expose its source (without duplicating the name value).
    try:
        exp = _safe(out.get("explicit-name", "")).strip().lower()
    except Exception:
        exp = ""
    if exp in ("true", "1", "yes", "on"):
        if "explicit-name-from" not in out or not _safe(out.get("explicit-name-from", "")).strip():
            out["explicit-name-from"] = computedFrom

    # Expanded / collapsed (prefer real NVDA states when available)
    if _state_in(obj, controlTypes.State.COLLAPSED) or _has_state_name(obj, "collapsed"):
        out["expanded"] = "false"
    elif _state_in(obj, controlTypes.State.EXPANDED) or _has_state_name(obj, "expanded"):
        out["expanded"] = "true"

    # Selected state (high-signal for tabs, options, treeitems, etc.).
    if "selected" not in out:
        try:
            if _state_in(obj, controlTypes.State.SELECTED) or _has_state_name(obj, "selected"):
                out["selected"] = "true"
            else:
                # Only report selected=false when the role meaningfully supports selection.
                try:
                    roleLower = _safe(out.get("role", "")).lower()
                    xmlRoleLower = _safe(out.get("xml-roles", "")).lower()
                except Exception:
                    roleLower = ""
                    xmlRoleLower = ""
                if roleLower in ("tab", "option", "treeitem") or xmlRoleLower in ("tab", "option", "treeitem"):
                    out["selected"] = "false"
        except Exception:
            pass

    
    # Checkable/checked (switch/checkbox-like): If IA2 reports checkable, also report checked state.
    # Keep minimal and prefer NVDA states.
    if "checkable" in out and "checked" not in out:
        try:
            roleLower = _safe(out.get("role", "")).lower()
            xmlRoleLower = _safe(out.get("xml-roles", "")).lower()
        except Exception:
            roleLower = ""
            xmlRoleLower = ""
        stChecked = False

        # Switch: NVDA often exposes on/off via state name, not CHECKED.
        if roleLower == "switch" or xmlRoleLower == "switch":
            try:
                if _has_state_name(obj, "on"):
                    stChecked = True
                elif _has_state_name(obj, "off"):
                    stChecked = False
            except Exception:
                pass
            if not stChecked:
                try:
                    stOn = getattr(getattr(controlTypes, "State", None), "ON", None)
                    if stOn is not None and _state_in(obj, stOn):
                        stChecked = True
                except Exception:
                    pass

        if not stChecked:
            try:
                stChecked = _state_in(obj, controlTypes.State.CHECKED) or _has_state_name(obj, "checked")
            except Exception:
                stChecked = False

        # Fallback: some toggles surface as PRESSED.
        if not stChecked and (roleLower == "switch" or xmlRoleLower == "switch"):
            try:
                stChecked = _state_in(obj, controlTypes.State.PRESSED) or _has_state_name(obj, "pressed")
            except Exception:
                stChecked = False

        out["checked"] = "true" if stChecked else "false"

# Toggle state (pressed): Prefer explicit IA2 attributes.
    # Many modern web apps expose aria-pressed rather than pressed.
    # Do NOT invent pressed for normal buttons; only infer when we have a
    # strong toggle signal.
    # Note: On some dynamic web apps, IA2Attributes may lag behind the current
    # visual/speech state. Prefer NVDA states when they indicate a pressed/
    # checked condition, but avoid inventing a pressed value for normal buttons.
    pressedFromAria = None
    ap = _safe(out.get("aria-pressed", "")).strip().lower()
    if ap:
        if ap in ("0", "false", "no"):
            pressedFromAria = False
        elif ap in ("1", "true", "yes"):
            pressedFromAria = True
        # Keep reports JAWS-like: don't show aria-pressed separately.
        out.pop("aria-pressed", None)
    if "pressed" not in out:
        # Fallback: NVDA states, but only when the role/tag strongly suggests a toggle.
        try:
            tagLower = _safe(out.get("tag", "")).lower()
        except Exception:
            tagLower = ""
        try:
            roleLower = _safe(out.get("role", "")).lower()
            xmlRoleLower = _safe(out.get("xml-roles", "")).lower()
        except Exception:
            roleLower = ""
            xmlRoleLower = ""
        is_toggle = (roleLower == "togglebutton") or (xmlRoleLower == "togglebutton")
        try:
            Role = getattr(controlTypes, "Role", None)
            if Role and not is_toggle:
                tb = getattr(Role, "TOGGLEBUTTON", None)
                if tb is not None and getattr(obj, "role", None) == tb:
                    is_toggle = True
        except Exception:
            pass
        if is_toggle and (tagLower in ("button", "a", "div", "span") or tagLower == ""):
            # Prefer explicit NVDA state flags when present.
            try:
                stPressed = _state_in(obj, controlTypes.State.PRESSED)
            except Exception:
                stPressed = False
            try:
                stChecked = _state_in(obj, controlTypes.State.CHECKED)
            except Exception:
                stChecked = False
            if stPressed or stChecked or _has_state_name(obj, "pressed") or _has_state_name(obj, "checked"):
                out["pressed"] = "true"
            elif pressedFromAria is not None:
                out["pressed"] = "true" if pressedFromAria else "false"
            else:
                out["pressed"] = "false"

    # Tabs: JAWS often reports pressed=true for the active tab.
    if "pressed" not in out:
        try:
            roleLower = _safe(out.get("role", "")).lower()
            xmlRoleLower = _safe(out.get("xml-roles", "")).lower()
        except Exception:
            roleLower = ""
            xmlRoleLower = ""
        if roleLower == "tab" or xmlRoleLower == "tab":
            if out.get("selected") == "true":
                out["pressed"] = "true"

    # If aria-pressed was present and we did not set pressed via toggle inference
    # above (e.g. because we couldn't confidently identify a toggle), keep the
    # aria value.
    if "pressed" not in out and pressedFromAria is not None:
        out["pressed"] = "true" if pressedFromAria else "false"

    # Normalize pressed when present.
    if "pressed" in out:
        try:
            pv = _safe(out.get("pressed", "")).strip().lower()
            if pv in ("0", "false", "no"):
                out["pressed"] = "false"
            elif pv in ("1", "true", "yes"):
                out["pressed"] = "true"
        except Exception:
            pass
    if t in ("td","th"):
        if "colspan" not in out:
            try:
                v = getattr(obj, "colSpan", None) or getattr(obj, "colspan", None)
                out["colspan"] = _safe(v) if v else "1"
            except Exception:
                out["colspan"] = "1"
        if "rowspan" not in out:
            try:
                v = getattr(obj, "rowSpan", None) or getattr(obj, "rowspan", None)
                out["rowspan"] = _safe(v) if v else "1"
            except Exception:
                out["rowspan"] = "1"
    if t == "table" and "layout-guess" not in out:
        out["layout-guess"] = "false"

    # tabindex: report only when explicitly exposed.
    # Do NOT infer tabindex for general elements, because many modern web apps (e.g. Google Docs)
    # use roving tabindex and programmatic focus that is not reliably exposed via IA2 attributes.
    # Exception: role=tab (roving tabindex patterns). For tabs, expose implicit tabindex=0 when focusable/focused.
    if "tabindex" not in out:
        try:
            tagLower = (out.get("tag") or _tag(obj) or "").lower()
        except Exception:
            tagLower = ""
        if tagLower not in ("#document", "body", "html"):
            try:
                roleLower = _safe(out.get("role", "")).lower()
                xmlRoleLower = _safe(out.get("xml-roles", "")).lower()
            except Exception:
                roleLower = ""
                xmlRoleLower = ""
            if roleLower == "tab" or xmlRoleLower == "tab":
                try:
                    is_focusable = _state_in(obj, controlTypes.State.FOCUSABLE) or _has_state_name(obj, "focusable")
                    is_focused = _state_in(obj, controlTypes.State.FOCUSED) or _has_state_name(obj, "focused")
                except Exception:
                    is_focusable = False
                    is_focused = False
                if is_focusable or is_focused:
                    out["tabindex"] = "0"

# _infer_form_attrs returns an updated dict; do not treat it like a callable.
    out = _infer_form_attrs(obj, out)
    out = _infer_expanded_for_combobox(obj, chain, out)

    
    # Safety: never show MSAA Role for non-control wrapper tags in the user report.
    try:
        _tL = _safe(out.get("tag", "")).strip().lower()
        if "MSAA Role" in out and not _is_control_tag(_tL):
            out.pop("MSAA Role", None)
    except Exception:
        pass

    # Read-only on #document is noise; don't report it.
    try:
        if _safe(out.get("tag", "")).strip().lower() == "#document":
            out.pop("readonly", None)
    except Exception:
        pass

# Clarify NVDA-computed description vs. attribute provenance.
    # If description came from aria-describedby, present it as describedby-text (JAWS-like) while keeping 'description-from'.
    try:
        df = _safe(out.get("description-from", "")).strip().lower()
        desc = _safe(out.get("description", "")).strip()
        if df == "aria-describedby" and desc:
            if "describedby-text" not in out:
                out["describedby-text"] = desc
    except Exception:
        pass


    # ---- DEBUG: Tabindex provenance ----
    if DEBUG_MODE:
        try:
            raw_tab = None
            try:
                ia2 = getattr(obj, "IA2Attributes", None) or {}
                raw_tab = ia2.get("tabindex")
            except Exception:
                raw_tab = None

            final_tab = out.get("tabindex")
            states = []
            try:
                st = getattr(obj, "states", [])
                states = list(st) if isinstance(st, (set, list, tuple)) else [st]
            except Exception:
                pass

            _dbg(f"[TABINDEX DEBUG] tag={_safe(out.get('tag',''))} raw_ia2={raw_tab} final_reported={final_tab}")
            _dbg(f"[TABINDEX DEBUG] states={states}")
        except Exception:
            pass

    return out



def _is_control_tag(tagLower):
    """Return True for HTML tags where MSAA accRole is usually meaningful."""
    if not tagLower:
        return False
    return tagLower in ("input", "textarea", "select", "button", "option", "meter", "progress")

def _is_web_context(base_obj):
    """Return True when inspection is meaningful.

    Priority: NVDA Browse Mode (virtual buffer) — works in browsers AND apps like Thunderbird.
    Secondary: IA2 #document tag or retrievable document URL (some backends expose these).
    """
    if not base_obj:
        return False

    def _is_browse_mode_obj(o):
        try:
            ti = getattr(o, "treeInterceptor", None)
        except Exception:
            ti = None
        if not ti:
            return False
        # Most reliable: is a BrowseModeTreeInterceptor
        try:
            if isinstance(ti, browseMode.BrowseModeTreeInterceptor):
                return True
        except Exception:
            pass
        # Best-effort fallback: common browse mode properties
        try:
            if hasattr(ti, "passThrough") and hasattr(ti, "script_quickNav_nextHeading"):
                return True
        except Exception:
            pass
        return False

    # 1) Browse mode check (navigator OR focus)
    try:
        if _is_browse_mode_obj(base_obj):
            return True
    except Exception:
        pass

    try:
        focus = api.getFocusObject()
    except Exception:
        focus = None
    if focus and focus is not base_obj:
        try:
            if _is_browse_mode_obj(focus):
                return True
        except Exception:
            pass

    # 2) IA2 tag chain ending in #document
    try:
        base_chain = _dom_chain_with_tags(base_obj)
    except Exception:
        base_chain = []
    try:
        for o in base_chain:
            if _tag(o) == "#document":
                return True
    except Exception:
        pass

    # 3) Fallback: retrievable document URL
    try:
        du = _document_url(base_obj, base_chain)
        if du:
            return True
    except Exception:
        pass

    if focus and focus is not base_obj:
        try:
            focus_chain = _dom_chain_with_tags(focus)
        except Exception:
            focus_chain = []
        try:
            for o in focus_chain:
                if _tag(o) == "#document":
                    return True
        except Exception:
            pass
        try:
            du = _document_url(focus, focus_chain)
            if du:
                return True
        except Exception:
            pass

    return False


def _format_tag_block(tag_name, attrs):
    keys = _ordered_params(tag_name, attrs)
    out_lines = []
    out_lines.append("")  # blank line BEFORE the header
    # Skip empty values to reduce noise (e.g., accessible-name= on containers).
    pairs = []
    for k in keys:
        v = _safe((attrs or {}).get(k, "")).strip()
        if v == "":
            continue
        pairs.append((k, v))
    out_lines.append(f"Tag {tag_name.upper()} has {len(pairs)} parameters:")
    for k, v in pairs:
        out_lines.append(f"{k}={v}")
    out_lines.append("")  # blank line after each block
    return "\n".join(out_lines)



def _build_report(advanced=False):
    base = _get_candidate_object()
    if not base:
        return _("No element found."), False

    # Gate: only run in HTML/virtual-buffer contexts.
    if not _is_web_context(base):
        return _("HTML Element Inspector works only when Browse Mode (virtual buffer) is available."), False

    base_chain = _dom_chain_with_tags(base)
    # Fallback: if we only got #document, try other common candidates.
    # This fixes cases where the navigator object becomes a document-level node
    # after an initial inspection dialog.
    if len(base_chain) == 1 and _tag(base_chain[0]) == "#document":
        alt = None
        try:
            alt = api.getNavigatorObject()
        except Exception:
            alt = None
        if alt and alt is not base:
            alt_chain = _dom_chain_with_tags(alt)
            if len(alt_chain) > len(base_chain):
                base = alt
                base_chain = alt_chain
        if len(base_chain) == 1 and _tag(base_chain[0]) == "#document":
            try:
                alt = api.getFocusObject()
            except Exception:
                alt = None
            if alt and alt is not base:
                alt_chain = _dom_chain_with_tags(alt)
                if len(alt_chain) > len(base_chain):
                    base = alt
                    base_chain = alt_chain
    canonical, promoted_from = _promote_canonical(base, base_chain)
    _dbg("---- INSPECTION START ----")
    _dbg_obj(base, "BASE")
    _dbg(f"Base object: {base}")
    _dbg(f"Base chain tags: {[ _tag(o) for o in base_chain ]}")
    _dbg(f"Canonical object: {canonical}")
    _dbg_obj(canonical, "CANONICAL")
    _dbg(f"Promoted from: {promoted_from}")
    if promoted_from:
        _dbg_obj(promoted_from, "PROMOTED_FROM")
    canonical_chain = _dom_chain_with_tags(canonical)

    lines = []
    lines.append("Advanced Element Information:" if advanced else "Element Information:")

    c_raw = _ia2_attrs(canonical)
    _dbg(f"Canonical IA2 raw attrs: {c_raw}")
    try:
        _dbg(f"Computed description (canonical.description): {getattr(canonical, 'description', None)}")
    except Exception:
        pass
    try:
        ia2 = getattr(canonical, 'IA2Attributes', None) or {}
        for k in ('aria-describedby', 'describedby', 'describedBy', 'describedby-text', 'description-from'):
            if k in ia2:
                _dbg(f"Canonical IA2 {k}: {ia2.get(k)}")
    except Exception:
        pass
    try:
        _dbg(f"Canonical NVDA role: {getattr(canonical, 'role', None)}")
    except Exception:
        pass
    c_attrs = _augment_attrs_for_readability(canonical, canonical_chain, c_raw, base)
    c_tag = c_attrs.get("tag", _tag(canonical)) or "unknown"
    lines.append(_format_tag_block(c_tag, c_attrs).rstrip())

    if promoted_from and promoted_from is not canonical:
        p_chain = _dom_chain_with_tags(promoted_from)
        p_raw = _ia2_attrs(promoted_from)
        p_attrs = _augment_attrs_for_readability(promoted_from, p_chain, p_raw, base)
        p_tag = p_attrs.get("tag", _tag(promoted_from)) or "unknown"
        lines.append("Nested element:")
        lines.append(_format_tag_block(p_tag, p_attrs).rstrip())

    ancCount = 0
    for ancestor in canonical_chain[1:]:
        ancCount += 1
        if ancCount <= 4:
            _dbg_obj(ancestor, f"ANCESTOR#{ancCount}")
        a_chain = _dom_chain_with_tags(ancestor)
        a_raw = _ia2_attrs(ancestor)
        _dbg(f"Ancestor tag: {_tag(ancestor)} | IA2 raw: {a_raw}")
        try:
            _dbg(f"Ancestor NVDA role: {getattr(ancestor, 'role', None)}")
        except Exception:
            pass
        a_attrs = _augment_attrs_for_readability(ancestor, a_chain, a_raw, base)
        a_tag = a_attrs.get("tag", _tag(ancestor)) or "unknown"
        lines.append(_format_tag_block(a_tag, a_attrs).rstrip())

    report = "\n".join([ln for ln in lines if ln is not None]).strip() + "\n"

    if not advanced:
        return report, True

    # Advanced mode: provide real headings for browse-mode navigation.
    def _h_escape(s):
        try:
            return html.escape(_safe(s))
        except Exception:
            return html.escape(str(s))

    def _pre_block(text):
        return "<pre>" + _h_escape(text) + "</pre>"

    # Build subtree (children) under the canonical element.
    subtree_lines = []
    max_depth = 3
    max_nodes = 30

    def _iter_subtree(root):
        if not root:
            return
        try:
            first = getattr(root, "firstChild", None)
        except Exception:
            first = None
        queue = []
        if first:
            queue.append((first, 1))
        seen = set()
        yielded = 0
        while queue and yielded < max_nodes:
            node, depth = queue.pop(0)
            if not node:
                continue
            oid = id(node)
            if oid in seen:
                continue
            seen.add(oid)
            # Only include nodes that expose a tag (DOM-like)
            try:
                a = _ia2_attrs(node)
                if not a.get("tag"):
                    # still traverse children
                    pass
                else:
                    yielded += 1
                    yield node, depth
            except Exception:
                pass

            if depth >= max_depth:
                # do not enqueue deeper
                try:
                    nxt = getattr(node, "next", None)
                except Exception:
                    nxt = None
                if nxt:
                    queue.append((nxt, depth))
                continue

            # enqueue children then siblings
            try:
                child = getattr(node, "firstChild", None)
            except Exception:
                child = None
            if child:
                queue.append((child, depth + 1))
            try:
                nxt = getattr(node, "next", None)
            except Exception:
                nxt = None
            if nxt:
                queue.append((nxt, depth))

    child_blocks = []
    truncated = False
    try:
        for node, depth in _iter_subtree(canonical):
            a_chain = _dom_chain_with_tags(node)
            a_raw = _ia2_attrs(node)
            a_attrs = _augment_attrs_for_readability(node, a_chain, a_raw, base)
            a_tag = a_attrs.get("tag", _tag(node)) or "unknown"
            # Keep children minimal: tag/role/name/states + small essentials.
            keep = {}
            for k in ("tag", "id", "role", "xml-roles", "href", "src", "type", "accessible-name", "accessible-name-from",
                      "explicit-name", "explicit-name-from",
                      "pressed", "expanded", "selected", "checked", "value", "valuetext"):
                if k in a_attrs:
                    keep[k] = a_attrs.get(k)
            # Let form inference add fsFormField when meaningful but don't force it.
            keep = _infer_form_attrs(node, keep)
            block = _format_tag_block(a_tag, keep).strip()
            child_blocks.append((depth, a_tag, block))
        # If queue had more nodes, we won't know; mark truncation only when we hit max_nodes exactly
        if len(child_blocks) >= max_nodes:
            truncated = True
    except Exception:
        child_blocks = []

    # Render advanced HTML with headings.
    parts = []
    parts.append("<h1>Advanced Element Information</h1>")
    parts.append("<h2>Focused element</h2>")
    # Render each tag block as its own heading for quick navigation.
    def _split_tag_blocks(txt):
        txt = (txt or "").replace("\r\n", "\n").replace("\r", "\n")
        blocks = []
        cur = []
        for line in txt.split("\n"):
            if line.strip() in ("Element Information:", "Advanced Element Information:"):
                continue
            if line.startswith("Tag ") and " has " in line and line.endswith("parameters:"):
                if cur:
                    blocks.append("\n".join(cur).strip("\n"))
                    cur = []
            cur.append(line)
        if cur:
            blocks.append("\n".join(cur).strip("\n"))
        return [b for b in blocks if b.strip()]

    for b in _split_tag_blocks(report):
        tag_name = None
        for ln in b.split("\n"):
            if ln.startswith("Tag ") and " has " in ln and " parameters:" in ln:
                try:
                    tag_name = ln.split("Tag ", 1)[1].split(" has ", 1)[0].strip()
                except Exception:
                    tag_name = None
                break
        if tag_name:
            parts.append(f"<h3>Tag {_h_escape(tag_name)}</h3>")
        parts.append(_pre_block(b))

    if child_blocks:
        parts.append("<h2>Children</h2>")
        for idx, (depth, tag, block) in enumerate(child_blocks, 1):
            level = 3 if depth <= 1 else 4 if depth == 2 else 5
            parts.append(f"<h{level}>Child {idx}: {tag.upper()}</h{level}>")
            parts.append(_pre_block(block))
        if truncated:
            parts.append("<p>... truncated (limits reached)</p>")
    else:
        parts.append("<h2>Children</h2>")
        parts.append("<p>No children exposed.</p>")
        parts.append("<p>Note: subtree not exposed by the accessibility API.</p>")

    html_report = "\n".join(parts)
    return html_report, True




def _report_text_to_html(report_text, title_h1):
    """Convert plain text report into simple HTML with headings for quick navigation."""
    def _esc(s):
        try:
            return html.escape(_safe(s))
        except Exception:
            return html.escape(str(s))

    txt = report_text or ""
    # Normalize line endings
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    # Split into tag blocks. Our formatter inserts a blank line before each 'Tag ...' header.
    blocks = []
    cur = []
    for line in txt.split("\n"):
        if line.startswith("Tag ") and " has " in line and line.endswith("parameters:"):
            if cur:
                blocks.append("\n".join(cur).strip("\n"))
                cur = []
        cur.append(line)
    if cur:
        blocks.append("\n".join(cur).strip("\n"))

    parts = []
    parts.append(f"<h1>{_esc(title_h1)}</h1>")
    # Filter out the initial "Element Information:" line if present.
    for b in blocks:
        b_stripped = b.strip()
        if not b_stripped:
            continue
        if b_stripped == "Element Information:" or b_stripped == "Advanced Element Information:":
            continue
        # Try to extract the tag name from the header line.
        tag_name = None
        for ln in b_stripped.split("\n"):
            if ln.startswith("Tag ") and " has " in ln and " parameters:" in ln:
                try:
                    tag_name = ln.split("Tag ", 1)[1].split(" has ", 1)[0].strip()
                except Exception:
                    tag_name = None
                break
        if tag_name:
            parts.append(f"<h2>Tag { _esc(tag_name) }</h2>")
        parts.append("<pre>" + _esc(b_stripped) + "</pre>")
    return "\n".join(parts)


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("Web Element Inspector")
    __gestures = {
        "kb:NVDA+shift+f1": "inspectWebElement",
        # Some NVDA builds normalize modifier order differently for Input Help.
        # Bind both canonical and legacy orders to ensure the command is announced.
        "kb:control+shift+NVDA+f1": "inspectWebElementAdvanced",
        "kb:control+NVDA+shift+f1": "inspectWebElementAdvanced",
    }

    @script(description=_("Inspect the current HTML element under the browse cursor (Basic report)."))
    def script_inspectWebElement(self, gesture):
        report, ok = _build_report(advanced=False)
        if not ok:
            ui.message(report)
            return
        htmlReport = _report_text_to_html(report, "Element Information")
        try:
            ui.browseableMessage(htmlReport, title=_("HTML Element Inspector"), isHtml=True)
        except TypeError:
            ui.browseableMessage(htmlReport, title=_("HTML Element Inspector"))
        ui.message(_("Inspector report shown."))

    @script(description=_("Inspect the focused HTML element and explore its children (advanced report)."))
    def script_inspectWebElementAdvanced(self, gesture):
        report, ok = _build_report(advanced=True)
        if not ok:
            ui.message(report)
            return
        # Advanced report is HTML (for heading navigation).
        try:
            ui.browseableMessage(report, title=_("HTML Element Inspector (Advanced)"), isHtml=True)
        except TypeError:
            # Older signatures: fall back to plain browseable message.
            ui.browseableMessage(report, title=_("HTML Element Inspector (Advanced)"))
        ui.message(_("Advanced inspector report shown."))

