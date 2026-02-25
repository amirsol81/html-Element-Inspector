# HTML Element Inspector

**Author:** Amir Soleimani  
**NVDA compatibility:** 2026.1 and later  

---

## Overview

HTML Element Inspector is a lightweight inspection tool inspired by the “Element Information” feature in JAWS.  
It generates a compact, semantic report for the currently focused web element.

---

## Shortcut

- **NVDA + Shift + F1** — Inspect the focused element.

---

## Requirements

- Use on web content while NVDA is in **Browse Mode** (Virtual Buffer).
- This add-on **does not copy anything to the clipboard**; it only displays a report.

---

## Browser Testing

This add-on has been tested primarily with **NVDA + Google Chrome**.  
Firefox compatibility has not yet been fully verified.

---

## What the Report Is Built From

- NVDA object model (role, states, name/value/description when available)
- IA2 / Web attributes (tag, role/xml-roles, aria relationships, setsize/posinset, etc.)
- A short ancestor chain to provide context

---

## Philosophy

- **Semantic first:** Prefer meaningful attributes over raw noise.
- **Conservative:** Avoid unsafe guessing (e.g., `tabindex` is only reported when explicitly exposed, except for `role="tab"` where it helps debugging roving tabindex patterns).
- **Readable:** Stable formatting designed for easy comparison with other tools.

---

## Feedback and Contributions

Feedback and contributions to improve this add-on are welcome.
