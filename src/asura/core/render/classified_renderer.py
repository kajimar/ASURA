from __future__ import annotations

import base64
import copy
import json
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from asura.core.render.pptx_renderer import render_pptx


def _load_json(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise TypeError(f"expected JSON object at {path}")
    return data


def _split_template_id(template_id: str) -> tuple[str, int | None]:
    m = re.fullmatch(r"(.+)-([0-9]+)", template_id.strip())
    if not m:
        return (template_id, None)
    return (m.group(1), int(m.group(2)))


def _iter_spec_pages(spec: dict[str, Any]) -> list[dict[str, Any]]:
    pages = spec.get("pages")
    if isinstance(pages, list):
        return [p for p in pages if isinstance(p, dict)]
    page_specs = spec.get("page_specs")
    if isinstance(page_specs, list):
        return [p for p in page_specs if isinstance(p, dict)]
    return []


def _slot_fields(spec: dict[str, Any]) -> list[str]:
    fields = spec.get("slot_fields")
    if isinstance(fields, list) and fields and all(isinstance(x, str) for x in fields):
        return [str(x) for x in fields]
    return ["slot_id", "role", "kind", "bbox", "style_fingerprint"]


def _normalize_slot(slot: Any, *, fields: list[str]) -> dict[str, Any] | None:
    if isinstance(slot, dict):
        return dict(slot)
    if isinstance(slot, list):
        out: dict[str, Any] = {}
        for idx, key in enumerate(fields):
            if idx < len(slot):
                out[key] = slot[idx]
        return out
    return None


def _bbox_key(bbox: Any) -> tuple[int, int, int, int] | None:
    if not isinstance(bbox, dict):
        return None
    try:
        x = int(round(float(bbox.get("x", 0))))
        y = int(round(float(bbox.get("y", 0))))
        w = int(round(float(bbox.get("w", 0))))
        h = int(round(float(bbox.get("h", 0))))
    except Exception:
        return None
    return (x, y, w, h)


def _build_pptx_media_sha_map(pptx_path: Path) -> dict[str, tuple[bytes, str]]:
    out: dict[str, tuple[bytes, str]] = {}
    with zipfile.ZipFile(pptx_path, "r") as zf:
        for name in zf.namelist():
            if not name.startswith("ppt/media/"):
                continue
            blob = zf.read(name)
            ext = Path(name).suffix.lower().lstrip(".")
            import hashlib

            sha = hashlib.sha256(blob).hexdigest()
            out[sha] = (blob, ext)
    return out


def _preferred_template_page(
    template_id: str,
    page_data: dict[str, Any],
    spec_pages: list[dict[str, Any]],
) -> dict[str, Any]:
    base_template_id, explicit_page = _split_template_id(template_id)
    if not spec_pages:
        raise ValueError(f"no template spec pages for {template_id}")

    if explicit_page is not None:
        for page in spec_pages:
            try:
                if int(page.get("page", -1) or -1) == explicit_page:
                    return page
            except Exception:
                continue
        return spec_pages[0]

    if base_template_id == "A1":
        for page in spec_pages:
            if str(page.get("sub_class", "")) == "cards_3up_plus_2down":
                return page
        return spec_pages[0]

    if base_template_id == "D1":
        steps = page_data.get("steps")
        step_count = len(steps) if isinstance(steps, list) else 0
        target = f"steps{step_count}_cards3"
        for page in spec_pages:
            if str(page.get("sub_class", "")) == target:
                return page
        for page in spec_pages:
            if str(page.get("sub_class", "")).startswith("steps"):
                return page
        return spec_pages[0]

    return spec_pages[0]


def _sorted_slots(spec_root: dict[str, Any], spec_page: dict[str, Any]) -> list[dict[str, Any]]:
    slots: list[dict[str, Any]] = []
    fields = _slot_fields(spec_root)
    for raw_slot in spec_page.get("slots", []):
        slot = _normalize_slot(raw_slot, fields=fields)
        if slot is not None:
            slots.append(slot)
    return sorted(
        slots,
        key=lambda slot: (
            int((slot.get("bbox") or {}).get("y", 0) or 0),
            int((slot.get("bbox") or {}).get("x", 0) or 0),
        ),
    )


def _text_slot_ids(
    spec_root: dict[str, Any],
    spec_page: dict[str, Any],
    *,
    prefix: str | None = None,
    suffix: str | None = None,
    exclude_substrings: tuple[str, ...] = (),
) -> list[str]:
    out: list[str] = []
    for slot in _sorted_slots(spec_root, spec_page):
        if str(slot.get("kind", "") or "").lower() != "text":
            continue
        slot_id = str(slot.get("slot_id", "") or "")
        if prefix and not slot_id.startswith(prefix):
            continue
        if suffix and not slot_id.endswith(suffix):
            continue
        if any(part in slot_id for part in exclude_substrings):
            continue
        out.append(slot_id)
    return out


def _assign_ordered(out: dict[str, str], slot_ids: list[str], values: list[str]) -> None:
    for idx, slot_id in enumerate(slot_ids):
        out[slot_id] = str(values[idx] if idx < len(values) else "" or "")


def _normalize_deck_input(raw: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    deck_meta = raw.get("deck_meta")
    if isinstance(deck_meta, dict):
        deck = dict(deck_meta)
    else:
        deck = {}

    if not deck:
        deck_src = raw.get("deck")
        if isinstance(deck_src, dict):
            deck = dict(deck_src)

    pages_src = raw.get("pages")
    if not isinstance(pages_src, list) or not pages_src:
        raise ValueError("input JSON has no pages")

    normalized_pages: list[dict[str, Any]] = []
    for page_entry in pages_src:
        if not isinstance(page_entry, dict):
            continue
        content = page_entry.get("content")
        normalized = dict(content) if isinstance(content, dict) else {}
        for key, value in page_entry.items():
            if key == "content":
                continue
            normalized[key] = value
        normalized_pages.append(normalized)

    if not normalized_pages:
        raise ValueError("input JSON has no usable page objects")

    return deck, normalized_pages


def _is_dynamic_text_slot(slot_id: str) -> bool:
    return (
        slot_id.startswith("meta.")
        or slot_id.startswith("content.")
        or slot_id.startswith("flow.step_")
        or slot_id.startswith("reference.item_")
        or slot_id.startswith("decor.footnote.")
    )


def _is_reportable_text_slot(slot_id: str) -> bool:
    return not (
        slot_id.endswith(".panel")
        or slot_id.endswith(".band")
        or slot_id.endswith(".icon_bg")
    )


def _footer_text(deck_meta: dict[str, Any], page_no: int) -> str:
    title = str(deck_meta.get("title", "") or "").strip()
    if title:
        return f"Page {page_no} | {title}"
    return f"Page {page_no}"


def _step_badge_value(raw_label: str, step_idx: int) -> tuple[str, str]:
    label = str(raw_label or "").strip()
    match = re.search(r"(\d+)", label)
    badge = match.group(1) if match else str(step_idx)
    if label and not re.fullmatch(r"(?:STEP\s*)?\d+", label, flags=re.IGNORECASE):
        return (badge, label)
    return (badge, f"STEP {badge}")


def _slot_value_map(
    page_data: dict[str, Any],
    deck_meta: dict[str, Any],
    *,
    spec_root: dict[str, Any],
    spec_page: dict[str, Any],
    out_page_no: int,
) -> dict[str, str]:
    template_id = str(page_data.get("template_id", "")).strip()
    base_template_id, _ = _split_template_id(template_id)
    out: dict[str, str] = {}
    out["meta.footer.page_info"] = _footer_text(deck_meta, out_page_no)
    out["meta.footer.text"] = _footer_text(deck_meta, out_page_no)

    if base_template_id == "E1":
        out["meta.eyebrow"] = str(page_data.get("eyebrow", "") or "")
        out["meta.title"] = str(page_data.get("title", "") or "")
        out["meta.subtitle"] = str(page_data.get("subtitle", "") or "")
        out["content.takeaway.title"] = "Takeaway"
        out["content.takeaway.body"] = str(page_data.get("takeaway", "") or "")
        out["meta.footer.deck_title"] = str(deck_meta.get("title", "") or "")
        tags = page_data.get("tags")
        if isinstance(tags, list):
            for idx in range(1, 8):
                out[f"content.tags.tag_{idx}.label"] = str(tags[idx - 1] if idx - 1 < len(tags) else "")
        return out

    if template_id == "A1-1" or (base_template_id == "A1" and page_data.get("cards")):
        out["meta.subtitle"] = str(page_data.get("subtitle", "") or "")
        out["meta.title"] = str(page_data.get("title", "") or "")
        out["content.takeaway.body"] = str(page_data.get("takeaway", "") or "")
        out["content.note.body"] = ""
        cards = page_data.get("cards")
        if isinstance(cards, list):
            if len(cards) >= 1 and isinstance(cards[0], dict):
                out["content.card_1.title"] = str(cards[0].get("title", "") or "")
                body = cards[0].get("body")
                if isinstance(body, list):
                    for idx in range(1, 4):
                        out[f"content.card_1.body_{idx}"] = str(body[idx - 1] if idx - 1 < len(body) else "")
            if len(cards) >= 2 and isinstance(cards[1], dict):
                out["content.card_2.title"] = str(cards[1].get("title", "") or "")
                body = cards[1].get("body")
                if isinstance(body, list):
                    for idx in range(1, 4):
                        out[f"content.card_2.item_{idx}.label"] = ""
                        out[f"content.card_2.item_{idx}.body"] = str(body[idx - 1] if idx - 1 < len(body) else "")
            if len(cards) >= 3 and isinstance(cards[2], dict):
                out["content.card_3.title"] = str(cards[2].get("title", "") or "")
                body = cards[2].get("body")
                if isinstance(body, list):
                    for idx in range(1, 4):
                        out[f"content.card_3.body_{idx}"] = str(body[idx - 1] if idx - 1 < len(body) else "")
        return out

    if template_id == "A1-2":
        out["meta.subtitle"] = str(page_data.get("subtitle", "") or "")
        out["meta.title"] = str(page_data.get("title", "") or "")
        out["content.takeaway.body"] = str(page_data.get("takeaway", "") or "")
        columns = page_data.get("columns")
        if isinstance(columns, list):
            left_col = columns[0] if len(columns) >= 1 and isinstance(columns[0], dict) else {}
            right_col = columns[1] if len(columns) >= 2 and isinstance(columns[1], dict) else {}
            out["content.card_1.title"] = str(left_col.get("title", "") or "")
            out["content.card_2.title"] = str(right_col.get("title", "") or "")
            left_bodies = [str(x or "") for x in left_col.get("body", [])] if isinstance(left_col.get("body"), list) else []
            right_bodies = [str(x or "") for x in right_col.get("body", [])] if isinstance(right_col.get("body"), list) else []
            _assign_ordered(
                out,
                _text_slot_ids(spec_root, spec_page, prefix="content.card_1.item_", suffix=".body"),
                left_bodies,
            )
            _assign_ordered(
                out,
                _text_slot_ids(spec_root, spec_page, prefix="content.card_2.item_", suffix=".body"),
                right_bodies,
            )
        return out

    if template_id == "A1-4":
        out["meta.subtitle"] = str(page_data.get("subtitle", "") or "")
        out["meta.title"] = str(page_data.get("title", "") or "")
        out["content.takeaway.body"] = str(page_data.get("takeaway", "") or "")
        tiers = page_data.get("tiers")
        if isinstance(tiers, list):
            for card_idx in range(1, 4):
                tier = tiers[card_idx - 1] if card_idx - 1 < len(tiers) and isinstance(tiers[card_idx - 1], dict) else {}
                prefix = f"content.card_{card_idx}."
                out[f"{prefix}header.title"] = str(tier.get("header", "") or "")
                label_slots = _text_slot_ids(
                    spec_root,
                    spec_page,
                    prefix=prefix,
                    suffix=".label",
                    exclude_substrings=("caption",),
                )
                body_slots = _text_slot_ids(spec_root, spec_page, prefix=prefix, suffix=".body")
                body_values = [str(x or "") for x in tier.get("body", [])] if isinstance(tier.get("body"), list) else []
                label_values: list[str] = []
                title = str(tier.get("title", "") or "")
                if title:
                    label_values.append(title)
                for idx in range(1, min(len(body_values), max(0, len(label_slots) - len(label_values))) + 1):
                    label_values.append(f"要点{idx}")
                _assign_ordered(out, label_slots, label_values)
                _assign_ordered(out, body_slots, body_values)
                out[f"{prefix}caption.label"] = str(tier.get("caption", "") or "")
        return out

    if base_template_id == "A3":
        out["meta.eyebrow"] = str(page_data.get("eyebrow", "") or "")
        out["meta.title"] = str(page_data.get("title", "") or "")
        out["content.takeaway.body"] = str(page_data.get("takeaway", "") or "")

        left_col = page_data.get("left_column")
        if isinstance(left_col, dict):
            out["content.card_1.title"] = str(left_col.get("title", "") or "")
            items = left_col.get("items")
            if isinstance(items, list):
                for idx in range(1, 4):
                    out[f"content.card_1.item_{idx}.badge.text"] = ""
                    out[f"content.card_1.item_{idx}.title"] = str(items[idx - 1] if idx - 1 < len(items) else "")
                    out[f"content.card_1.item_{idx}.body"] = ""

        right_col = page_data.get("right_column")
        if isinstance(right_col, dict):
            out["content.card_2.title"] = str(right_col.get("title", "") or "")
            items = right_col.get("items")
            if isinstance(items, list):
                for idx in range(1, 4):
                    out[f"content.card_2.item_{idx}.badge.text"] = ""
                    out[f"content.card_2.item_{idx}.title"] = str(items[idx - 1] if idx - 1 < len(items) else "")
                    out[f"content.card_2.item_{idx}.body"] = ""
        return out

    if base_template_id == "D1":
        out["meta.subtitle"] = str(page_data.get("subtitle", "") or "")
        out["meta.title"] = str(page_data.get("title", "") or "")
        out["content.takeaway"] = str(page_data.get("takeaway", "") or "")
        steps = page_data.get("steps")
        if isinstance(steps, list):
            for idx in range(1, 6):
                step = steps[idx - 1] if idx - 1 < len(steps) and isinstance(steps[idx - 1], dict) else {}
                badge, label = _step_badge_value(str(step.get("label", "") or ""), idx)
                out[f"flow.step_{idx}.badge_number"] = badge
                out[f"flow.step_{idx}.label"] = label
                out[f"flow.step_{idx}.body"] = str(step.get("body", "") or "")
        cards = page_data.get("cards")
        if isinstance(cards, list):
            for card_idx in range(1, 4):
                card = cards[card_idx - 1] if card_idx - 1 < len(cards) and isinstance(cards[card_idx - 1], dict) else {}
                out[f"content.card_{card_idx}.title"] = str(card.get("title", "") or "")
                body = card.get("body")
                if isinstance(body, list):
                    out[f"content.card_{card_idx}.body_1"] = str(body[0] if len(body) >= 1 else "")
                    out[f"content.card_{card_idx}.body_2"] = str(body[1] if len(body) >= 2 else "")
        return out

    raise ValueError(f"unsupported template_id: {template_id}")


def _template_text_style(chunk: dict[str, Any]) -> dict[str, Any]:
    ts = chunk.get("text_struct")
    if isinstance(ts, dict):
        paragraphs = ts.get("paragraphs")
        if isinstance(paragraphs, list):
            for para in paragraphs:
                if not isinstance(para, dict):
                    continue
                runs = para.get("runs")
                if isinstance(runs, list):
                    for run in runs:
                        if isinstance(run, dict):
                            return copy.deepcopy(run)
    return {}


def _set_chunk_text(chunk: dict[str, Any], value: str) -> None:
    text = str(value or "")
    chunk["text"] = text
    chunk["normalized_text"] = text

    style_run = _template_text_style(chunk)
    ts = chunk.get("text_struct")
    if not isinstance(ts, dict):
        if not text:
            return
        chunk["text_struct"] = {
            "paragraphs": [
                {
                    "index": 0,
                    "alignment": "LEFT (1)",
                    "level": 0,
                    "runs": [{"index": 0, "text": text}],
                }
            ]
        }
        return

    old_paragraphs = ts.get("paragraphs")
    base_paragraphs = old_paragraphs if isinstance(old_paragraphs, list) and old_paragraphs else [{}]

    if not text:
        new_paragraphs: list[dict[str, Any]] = []
        for idx, para in enumerate(base_paragraphs[:1]):
            if not isinstance(para, dict):
                para = {}
            new_paragraphs.append(
                {
                    "index": idx,
                    "alignment": para.get("alignment"),
                    "level": para.get("level", 0),
                    "runs": [],
                }
            )
        ts["paragraphs"] = new_paragraphs
        return

    lines = [line for line in text.splitlines()]
    if not lines:
        lines = [text]

    new_paragraphs = []
    for idx, line in enumerate(lines):
        src_para = base_paragraphs[idx] if idx < len(base_paragraphs) and isinstance(base_paragraphs[idx], dict) else {}
        run = copy.deepcopy(style_run)
        run["index"] = 0
        run["text"] = line
        new_paragraphs.append(
            {
                "index": idx,
                "alignment": src_para.get("alignment"),
                "level": src_para.get("level", 0),
                "runs": [run],
            }
        )
    ts["paragraphs"] = new_paragraphs


def _embed_image_bytes(chunks: list[dict[str, Any]], media_map: dict[str, tuple[bytes, str]]) -> None:
    for chunk in chunks:
        img = chunk.get("image")
        if not isinstance(img, dict):
            continue
        sha = img.get("sha256")
        if not isinstance(sha, str) or sha not in media_map:
            continue
        blob, ext = media_map[sha]
        img["bytes_b64"] = base64.b64encode(blob).decode("ascii")
        img["ext"] = img.get("ext") or ext


def _match_slots_to_chunks(
    *,
    spec_root: dict[str, Any],
    spec_page: dict[str, Any],
    extraction_page_chunks: list[dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    bbox_index: dict[tuple[int, int, int, int], list[dict[str, Any]]] = {}
    for chunk in extraction_page_chunks:
        key = _bbox_key(chunk.get("bbox"))
        if key is None:
            continue
        bbox_index.setdefault(key, []).append(chunk)

    out: dict[str, dict[str, Any]] = {}
    for raw_slot in spec_page.get("slots", []):
        slot = _normalize_slot(raw_slot, fields=_slot_fields(spec_root))
        if slot is None:
            continue
        slot_id = slot.get("slot_id")
        key = _bbox_key(slot.get("bbox"))
        if not isinstance(slot_id, str) or key is None:
            continue

        candidates = bbox_index.get(key, [])
        slot_kind = str(slot.get("kind", "") or "").lower().strip()

        chosen = None
        if slot_kind:
            for chunk in candidates:
                if str(chunk.get("kind", "") or "").lower().strip() == slot_kind:
                    chosen = chunk
                    break
        if chosen is None and candidates:
            chosen = candidates[0]
        if chosen is not None:
            out[slot_id] = chosen
    return out


def prepare_classified_render_input(
    *,
    classified_path: Path,
    templates_spec_dir: Path,
    template_runs_dir: Path,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    classified = _load_json(classified_path)
    deck_meta, pages = _normalize_deck_input(classified)

    output_chunks: list[dict[str, Any]] = []
    render_report: list[dict[str, Any]] = []
    page_w = 12192000
    page_h = 6858000

    for out_page_no, page_data in enumerate(pages, start=1):
        template_id = str(page_data.get("template_id", "") or "").strip()
        if not template_id:
            raise ValueError(f"missing template_id on page {out_page_no}")
        base_template_id, _ = _split_template_id(template_id)

        spec_path = templates_spec_dir / f"{template_id}.json"
        if not spec_path.exists():
            spec_path = templates_spec_dir / f"{base_template_id}.json"
        run_dir = template_runs_dir / base_template_id
        extraction_path = run_dir / "extraction.json"
        source_pptx = run_dir / "source.pptx"

        spec_root = _load_json(spec_path)
        spec_pages = _iter_spec_pages(spec_root)
        spec_page = _preferred_template_page(template_id, page_data, spec_pages)
        spec_page_no = int(spec_page.get("page", 1) or 1)

        extraction = _load_json(extraction_path)
        doc = extraction.get("document")
        if isinstance(doc, dict):
            page_meta = doc.get("page")
            if isinstance(page_meta, dict):
                page_w = int(page_meta.get("w_emu", page_w) or page_w)
                page_h = int(page_meta.get("h_emu", page_h) or page_h)

        extraction_chunks = extraction.get("chunks")
        if not isinstance(extraction_chunks, list):
            raise ValueError(f"invalid extraction.chunks in {extraction_path}")

        page_chunks = [
            copy.deepcopy(ch)
            for ch in extraction_chunks
            if isinstance(ch, dict) and int(ch.get("page", -1) or -1) == spec_page_no
        ]

        media_map = _build_pptx_media_sha_map(source_pptx) if source_pptx.exists() else {}
        _embed_image_bytes(page_chunks, media_map)

        slot_to_chunk = _match_slots_to_chunks(
            spec_root=spec_root,
            spec_page=spec_page,
            extraction_page_chunks=page_chunks,
        )
        values = _slot_value_map(
            page_data,
            deck_meta,
            spec_root=spec_root,
            spec_page=spec_page,
            out_page_no=out_page_no,
        )

        page_report = {
            "page": out_page_no,
            "template_id": template_id,
            "cleared_or_unmapped_text_slots": [],
            "unmatched_slots": [],
        }

        for slot in _sorted_slots(spec_root, spec_page):
            slot_id = str(slot.get("slot_id", "") or "")
            if str(slot.get("kind", "") or "").lower() != "text":
                continue
            if not _is_dynamic_text_slot(slot_id):
                continue
            if slot_id not in slot_to_chunk:
                if _is_reportable_text_slot(slot_id):
                    page_report["unmatched_slots"].append(slot_id)
                continue
            chunk = slot_to_chunk[slot_id]
            _set_chunk_text(chunk, values.get(slot_id, ""))
            if _is_reportable_text_slot(slot_id) and not str(values.get(slot_id, "") or "").strip():
                page_report["cleared_or_unmapped_text_slots"].append(slot_id)

        for chunk_idx, chunk in enumerate(page_chunks, start=1):
            chunk["page"] = out_page_no
            chunk["chunk_id"] = f"s{out_page_no:03d}_c{chunk_idx:05d}"
            output_chunks.append(chunk)

        render_report.append(page_report)

    document_id = re.sub(r"[^A-Za-z0-9_-]+", "_", classified_path.stem).strip("_") or "classified_deck"
    extraction = {
        "schema_version": "0.1",
        "document": {
            "document_id": document_id[:64],
            "source_type": "pptx",
            "source_path": str(classified_path),
            "page_count": len(pages),
            "page": {
                "w_emu": page_w,
                "h_emu": page_h,
            },
        },
        "chunks": output_chunks,
    }
    return extraction, render_report


def build_classified_extraction(
    *,
    classified_path: Path,
    templates_spec_dir: Path,
    template_runs_dir: Path,
) -> dict[str, Any]:
    extraction, _ = prepare_classified_render_input(
        classified_path=classified_path,
        templates_spec_dir=templates_spec_dir,
        template_runs_dir=template_runs_dir,
    )
    return extraction


def render_classified_pptx(
    *,
    classified_path: str | Path,
    out_pptx: str | Path,
    templates_spec_dir: str | Path,
    template_runs_dir: str | Path,
    dump_extraction_path: str | Path | None = None,
) -> Path:
    classified_path = Path(classified_path).resolve()
    out_pptx = Path(out_pptx).resolve()
    templates_spec_dir = Path(templates_spec_dir).resolve()
    template_runs_dir = Path(template_runs_dir).resolve()
    dump_path = Path(dump_extraction_path).resolve() if dump_extraction_path else None

    extraction = build_classified_extraction(
        classified_path=classified_path,
        templates_spec_dir=templates_spec_dir,
        template_runs_dir=template_runs_dir,
    )

    if dump_path is not None:
        dump_path.parent.mkdir(parents=True, exist_ok=True)
        dump_path.write_text(json.dumps(extraction, ensure_ascii=False, indent=2), encoding="utf-8")

    with tempfile.TemporaryDirectory(prefix="asura-classified-render-") as tmp:
        run_dir = Path(tmp)
        (run_dir / "extraction.json").write_text(
            json.dumps(extraction, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        out_pptx.parent.mkdir(parents=True, exist_ok=True)
        render_pptx(run_dir=run_dir, out_pptx=out_pptx, mode="dom")

    return out_pptx
