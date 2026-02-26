import argparse
import time
import json
from pathlib import Path

import orjson
import requests

def build_schema(topn: int) -> dict:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "decision": {"type": "string", "enum": ["MATCH", "PARTIAL", "NO_MATCH"]},
            "chosen": {
                "type": "array",
                "minItems": 0,
                "maxItems": 2,
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "idx": {"type": "integer", "minimum": 0, "maximum": max(0, topn - 1)},
                        "pdf_chunk_id": {"type": "string"},
                    },
                    "required": ["idx", "pdf_chunk_id"],
                },
            },
            "transform": {
                "type": "string",
                "enum": ["VERBATIM", "ABBREVIATED", "PARAPHRASE", "SUMMARY", "NONE"],
            },
        },
        "required": ["decision", "chosen", "transform"],
    }

def ollama_chat_stream(
    session: requests.Session,
    host: str,
    payload: dict,
    read_timeout_s: int,
) -> str:
    url = f"{host}/api/chat"
    # stream=True にして「何かが返る」状態を維持する
    with session.post(url, json=payload, stream=True, timeout=(10, read_timeout_s)) as r:
        r.raise_for_status()
        parts: list[str] = []
        for line in r.iter_lines(decode_unicode=True):
            if not line:
                continue
            obj = json.loads(line)
            msg = obj.get("message") or {}
            if "content" in msg and msg["content"]:
                parts.append(msg["content"])
            if obj.get("done"):
                break
        return "".join(parts).strip()

def shorten(s: str, n: int) -> str:
    s = " ".join((s or "").split())
    if len(s) <= n:
        return s
    return s[:n] se.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--ollama_model", required=True)
    ap.add_argument("--ollama_host", default="http://localhost:11434")

    ap.add_argument("--topn", type=int, default=3, help="候補数（まずは3推奨）")
    ap.add_argument("--pdf_truncate", type=int, default=360, help="候補PDF本文の最大文字数")
    ap.add_argument("--pptx_truncate", type=int, default=200, help="PPTXテキストの最大文字数")

    ap.add_argument("--num_predict", type=int, default=160, help="出力上限（暴走防止）")
    ap.add_argument("--temperature", type=float, default=0.0)

    ap.add_argument("--max", type=int, default=50)
    ap.add_argument("--read_timeout", type=int, default=1800, help="1リクエストの最大待ち秒")
    ap.add_argument("--retries", type=int, default=2)
    ap.add_argument("--retry_sleep", type=float, default=1.5)
    args = ap.parse_args()

    inp = Path(ad_schema(args.topn)

    sess = requests.Session()

    n = 0
    with inp.open("rb") as fin, out.open("wb") as fout:
        for line in fin:
            if not line.strip():
                continue
            rec = orjson.loads(line)

            pptx_text = shorten(rec.get("pptx_text", ""), args.pptx_truncate)
            cands = (rec.get("candidates") or [])[: args.topn]

            # 候補本文を短くする（ここがタイムアウトの主因）
            cand_lines = []
            for i, c in enumerate(cands):
                pdf_text = shorten(c.get("pdf_text", ""), args.pdf_truncate)
                sim = c.get("sim")
                rr  = c.get("rerank_score")
                hdr = f"[{i}] {c.get('pdf_chunk_id')} p={c.get('pdf_page_no')} sim={sim:.3f}" if isinstance(sim, (int, float)) else f"[{i}] {c.get('pdf_chunk_id')} p={c.get('pdf_page_no')}"
                if isinstance(rr, (int, float)):
                    hdr += f" rerank={rr:.4f}"
                cand_lines.append(hdr + "\n"       "あなたは照合器です。PPTX_TEXT が PDF_CANDIDATES のどれに由来するか判定してください。\n"
                "ルール:\n"
                "- 同義・言い換えでも由来が明確なら MATCH\n"
                "- 一部だけ一致/要約なら PARTIAL\n"
                "- 候補に無いなら NO_MATCH\n"
                "- 出力は JSON のみ（説明禁止）\n\n"
                f"PPTX_TEXT:\n{pptx_text}\n\n"
                "PDF_CANDIDATES:\n" + "\n\n".join(cand_lines)
            )

            payload = {
                "model": args.ollama_model,
                "messages": [{"role": "user", "content": prompt}],
                "stream": True,
                "format": schema,
                "options": {
                    "temperature": args.temperature,
                    "num_predict": args.num_predict,
                },
                "keep_alive": "30m",
            }

            raw = ""
            err = None
            for attempt in range(arg  break
                except Exception as e:
                    err = f"{type(e).__name__}: {e}"
                    time.sleep(args.retry_sleep)

            outrec = {
                "pptx_unit_id": rec.get("pptx_unit_id"),
                "pptx_page": rec.get("pptx_page"),
                "pptx_kind": rec.get("pptx_kind"),
                "pptx_text": rec.get("pptx_text"),
                "topn": args.topn,
                "llm_raw": raw,
                "error": err,
            }

            # 可能ならJSONとしてもパースして付加
            if raw and not err:
                try:
                    outrec["llm_json"] = orjson.loads(raw.encode("utf-8"))
                except Exception as e:
                    outrec["error"] = f"parse_error: {type(e).__name__}: {e}"

            fout.write(orjson.dumps(outrec))
            fout.write(b"\n")

            n += 1
            if n % 10 == 0:
                print(f"[judge] done {n}")
            if n >= args.max:
                break> {out}")

if __name__ == "__main__":
    main()
