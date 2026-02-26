# ai/src/apps/asura/align/build_pdf_index.py
import argparse, pathlib
import numpy as np
import orjson
import hnswlib
from tqdm import tqdm
from sentence_transformers import SentenceTransformer

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf_json", required=True)
    ap.add_argument("--out_dir", required=True)
    ap.add_argument("--embed_model", default="intfloat/multilingual-e5-base")
    ap.add_argument("--batch", type=int, default=64)
    ap.add_argument("--ef_construction", type=int, default=200)
    ap.add_argument("--M", type=int, default=16)
    ap.add_argument("--ef_search", type=int, default=64)
    args = ap.parse_args()

    out = pathlib.Path(args.out_dir)
    out.mkdir(parents=True, exist_ok=True)

    data = orjson.loads(pathlib.Path(args.pdf_json).read_bytes())
    chunks = data.get("chunks", [])

    # PDF schema: block_type/page_no/text/chunk_id...
    recs = []
    for c in chunks:
        if c.get("block_type") != "text":
            continue
        t = (c.get("text") or "").strip()
        if not t:
            continue
        recs.append({
            "id": len(recs),
            "chunk_id": c.get("chunk_id"),
            "page_no": c.get("page_no"),
            "text": t,
            "normalized_text": (c.get("normalized_text") or ""),
        })

    if not recs:
        raise SystemExit("no text chunks found in pdf_json")

    model = SentenceTransformer(args.embed_model)
    texts = [f"passage: {r['text']}" for r in recs]

    embs = []
    for i in tqdm(range(0, len(texts), args.batch), desc="embed pdf"):
        batch = texts[i:i+args.batch]
        vec = model.encode(batch, normalize_embeddings=True)
        embs.append(vec)

    X = np.vstack(embs).astype(np.float32)
    dim = X.shape[1]

    index = hnswlib.Index(space="cosine", dim=dim)
    index.init_index(max_elements=len(recs), ef_construction=args.ef_construction, M=args.M)
    index.set_ef(args.ef_search)
    index.add_items(X, np.arange(len(recs)))

    (out / "dim.txt").write_text(str(dim))
    index.save_index(str(out / "index.hnsw"))

    with (out / "meta.jsonl").open("wb") as f:
        for r in recs:
            f.write(orjson.dumps(r))
            f.write(b"\n")

    print(f"OK: pdf_chunks={len(recs)} dim={dim}")
    print(f"index: {out/'index.hnsw'}")
    print(f"meta : {out/'meta.jsonl'}")

if __name__ == "__main__":
    main()