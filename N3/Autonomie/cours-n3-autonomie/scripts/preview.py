#!/usr/bin/env python3
"""Small dependency-free preview server for project Markdown and SVG files."""

from __future__ import annotations

import argparse
import html
import re
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


class PreviewHandler(SimpleHTTPRequestHandler):
    def do_GET(self) -> None:
        clean_path = self.path.split("?", 1)[0].split("#", 1)[0]
        if clean_path.lower().endswith(".md"):
            local = Path(self.translate_path(clean_path))
            if local.is_file():
                self.render_markdown(local)
                return
        super().do_GET()

    def render_markdown(self, path: Path) -> None:
        source = path.read_text(encoding="utf-8")
        source = re.sub(r"\A---\n.*?\n---\n", "", source, flags=re.DOTALL)
        output: list[str] = []
        in_list = False
        for raw in source.splitlines():
            line = raw.rstrip()
            if line.startswith("# "):
                if in_list:
                    output.append("</ul>")
                    in_list = False
                output.append(f"<h1>{html.escape(line[2:])}</h1>")
            elif line.startswith("## "):
                if in_list:
                    output.append("</ul>")
                    in_list = False
                output.append(f"<h2>{html.escape(line[3:])}</h2>")
            elif line.startswith("- "):
                if not in_list:
                    output.append("<ul>")
                    in_list = True
                output.append(f"<li>{linkify(line[2:])}</li>")
            elif line:
                if in_list:
                    output.append("</ul>")
                    in_list = False
                output.append(f"<p>{linkify(line)}</p>")
        if in_list:
            output.append("</ul>")
        body = "\n".join(output)
        page = f"""<!doctype html><meta charset=\"utf-8\"><title>{html.escape(path.name)}</title>
<style>body{{font:18px/1.5 system-ui;max-width:1000px;margin:2rem auto;padding:0 1rem;color:#162638}}h1{{color:#075985}}h2{{margin-top:2rem;border-bottom:1px solid #ccd6df}}img{{max-width:100%}}code{{background:#eef2f5;padding:.1rem .3rem}}</style>
<body><p><a href=\"./\">← Retour</a></p>{body}</body>"""
        encoded = page.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)


def linkify(text: str) -> str:
    escaped = html.escape(text)
    return re.sub(r"\[([^]]+)]\(([^)]+)\)", r'<a href="\2">\1</a>', escaped)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--project", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    handler = lambda *handler_args, **kwargs: PreviewHandler(*handler_args, directory=str(args.project), **kwargs)
    server = ThreadingHTTPServer(("127.0.0.1", args.port), handler)
    print(f"Prévisualisation : http://127.0.0.1:{args.port}/slides/")
    server.serve_forever()


if __name__ == "__main__":
    main()
