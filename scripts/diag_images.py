import json
import zipfile
import hashlib
import collections
from pathlib import Path
import argparse


def sha256(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def build_media_map(pptx: Path):
    m = {}
    by_size_ext = collections.defaultdict(list)
    with zipfile.ZipFile(pptx, "r") as z:
        for n in z.namelist():
            if not n.startswith("ppt/media/"):
                continue
            b = z.read(n)
            ext = Path(n).suffix.lower().lstrip(".")
            h = sha256(b)
            m[h] = (b, ext, n, len(b))
            by_size_ext[(len(b), ext)].append(h)
    return m, by_size_ext


def iter_image_chunks(extraction):
    stack = [extraction]
    while stack:
        x = stack.pop()
        if isinstance(x, dict):
            if x.get("kind") == "image":
                yield x
            for v in x.values():
                stack.append(v)
        elif isinstance(x, list):
            stack.extend(x)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run_dir", required=True)
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--max_examples", type=int, default=20)
    args = ap.parse_args()

    run_dir = Path(args.run_dir)
    extraction_path = run_dir / "extraction.json"
    pptx_path = Path(args.pptx)

    ex = json.loads(extraction_path.read_text(encoding="utf-8"))
    media_map, by_size_ext = build_media_map(pptx_path)

    img_chunks = list(iter_image_chunks(ex))
    total = len(img_chunks)

    sample_n = min(args.max_examples, total)
    if sample_n:
        print("=== IMAGE CHUNK SAMPLE KEYS ===")
        for i in range(sample_n):
            ch = img_chunks[i]
            keys = sorted(list(ch.keys()))
            print(i, ch.get("chunk_id"), keys)
            for k in keys:
                v = ch.get(k)
                t = type(v).__name__
                if isinstance(v, (str, int, float, bool)) or v is None:
                    s = v
                elif isinstance(v, (list, dict)):
                    s = t
                else:
                    s = t
                print(" ", k, t, s)
            img = ch.get("image")
            if isinstance(img, dict):
                img_keys = sorted(list(img.keys()))
                print("  image_keys", img_keys)
                for kk in img_keys:
                    vv = img.get(kk)
                    tt = type(vv).__name__
                    if isinstance(vv, (str, int, float, bool)) or vv is None:
                        ss = vv
                    elif isinstance(vv, (list, dict)):
                        ss = tt
                    else:
                        ss = tt
                    print("   ", kk, tt, ss)

    ext_count = collections.Counter()
    sha_miss = 0
    size_ext_amb = 0
    vector_like = 0

    sha_present = 0
    sha_missing = 0
    size_present = 0
    size_missing = 0
    ext_present = 0
    ext_missing = 0

    amb_examples = []
    miss_examples = []
    vector_examples = []

    by_size = collections.defaultdict(list)
    for h, (_b, ext2, name2, sz2) in media_map.items():
        by_size[int(sz2)].append((h, ext2, name2))

    for ch in img_chunks:
        img = ch.get("image")
        if isinstance(img, dict):
            ext = (img.get("ext") or img.get("format") or img.get("mime") or "").lower().lstrip(".")
            h = (img.get("sha256") or img.get("sha") or img.get("hash") or "").strip()
            bs = img.get("byte_size") or img.get("bytes") or img.get("size")
        else:
            ext = (ch.get("ext") or "").lower().lstrip(".")
            h = (ch.get("sha256") or ch.get("sha") or "").strip()
            bs = ch.get("byte_size") or ch.get("bytes") or ch.get("size")

        if ext:
            ext_present += 1
        else:
            ext_missing += 1

        if h:
            sha_present += 1
        else:
            sha_missing += 1

        if bs is not None:
            size_present += 1
            try:
                bs_int = int(bs)
            except Exception:
                bs_int = None
        else:
            size_missing += 1
            bs_int = None

        resolved_ext = None
        if h:
            if h in media_map:
                resolved_ext = media_map[h][1]
            else:
                sha_miss += 1
                if len(miss_examples) < args.max_examples:
                    miss_examples.append((ch.get("chunk_id"), ext, bs, h))

        ext_for_stats = ext or resolved_ext
        if ext_for_stats:
            ext_count[ext_for_stats] += 1

        if ext_for_stats in {"svg", "emf", "wmf"}:
            vector_like += 1
            if len(vector_examples) < args.max_examples:
                vector_examples.append((ch.get("chunk_id"), ext_for_stats, bs, h))

        if bs_int is not None:
            cands_all = by_size.get(bs_int, [])
            if len(cands_all) != 1:
                size_ext_amb += 1
                if len(amb_examples) < args.max_examples:
                    amb_examples.append((ch.get("chunk_id"), ext_for_stats or "", bs_int, len(cands_all)))

    print("=== IMAGE CHUNK FIELD PRESENCE ===")
    print("sha_present:", sha_present)
    print("sha_missing:", sha_missing)
    print("size_present:", size_present)
    print("size_missing:", size_missing)
    print("ext_present:", ext_present)
    print("ext_missing:", ext_missing)

    print("=== IMAGE CHUNK DIAGNOSTICS ===")
    print("pptx_media_items:", len(media_map))
    print("image_chunks_total:", total)
    print("ext_breakdown:", dict(ext_count.most_common()))
    if total:
        print("sha_miss:", sha_miss, f"({sha_miss/total*100:.1f}%)")
        print("size_ext_ambiguous:", size_ext_amb, f"({size_ext_amb/total*100:.1f}%)")
        print("vector_like(svg/emf/wmf):", vector_like, f"({vector_like/total*100:.1f}%)")
    else:
        print("sha_miss: 0")
        print("size_ext_ambiguous: 0")
        print("vector_like(svg/emf/wmf): 0")

    if miss_examples:
        print("-- sha_miss examples --")
        for e in miss_examples:
            print(e)

    if amb_examples:
        print("-- size/ext ambiguous examples --")
        for e in amb_examples:
            print(e)

    if vector_examples:
        print("-- vector-like examples --")
        for e in vector_examples:
            print(e)


if __name__ == "__main__":
    main()