"""Generate the complete Python symbol and dashboard source index from real files."""
from __future__ import annotations

import argparse
import ast
import json
from pathlib import Path
import re

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def iter_symbols(tree: ast.AST, prefix: str = ""):
    """Yield classes, methods, and nested functions with stable qualified names."""
    for node in ast.iter_child_nodes(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            name = prefix + node.name
            signature = name + ("(" + ast.unparse(node.args) + ")" if hasattr(node, "args") else "")
            doc = (ast.get_docstring(node) or "").split("\n")[0]
            yield {"name": name, "signature": signature, "line": node.lineno,
                   "kind": "class" if isinstance(node, ast.ClassDef) else "function", "summary": doc}
            yield from iter_symbols(node, name + ".")
        else:
            yield from iter_symbols(node, prefix)


def render_code_index(project_root: Path = PROJECT_ROOT) -> str:
    catalog = json.loads((project_root / "config/architecture/modules.json").read_text(encoding="utf-8"))["modules"]
    lines = ["# 코드 색인", "", "`python tools/update_code_index.py`로 재생성합니다. 파일 역할은 `config/architecture/modules.json`에서 관리합니다.",
             "클래스·메서드·내부 함수까지 실제 AST에서 추출합니다. 함수 이름 뒤 괄호는 입력 인자이며, 설명이 있는 항목은 함수 docstring을 함께 표시합니다.", ""]
    for relative, entry in catalog.items():
        path = project_root / "src/solar_forecast" / relative
        lines += [f"## [{relative}](../src/solar_forecast/{relative})", "", entry["purpose"], ""]
        symbols = list(iter_symbols(ast.parse(path.read_text(encoding="utf-8"))))
        for symbol in symbols:
            summary = (" — " + symbol["summary"]) if symbol["summary"] else ""
            signature = symbol["signature"].replace("`", "'").replace("\n", " ")
            lines.append(f"- `{signature}`{summary}")
        if not symbols:
            lines.append("공개 export 또는 설정 상수만 정의합니다.")
        lines.append("")
    lines += ["## 대시보드 원본", "", "`dashboard/src/`를 수정하고 빌드 도구로 `dashboard/assets/dashboard.js`를 생성합니다.", ""]
    for path in sorted((project_root / "dashboard/src").glob("*.js")):
        relative = path.relative_to(project_root).as_posix()
        names = re.findall(r"^\s*(?:async\s+)?function\s+(\w+)\(", path.read_text(encoding="utf-8"), re.M)
        lines += [f"### [{relative}](../{relative})", "", ", ".join(f"`{name}`" for name in names) or "화면 상태 초기화 또는 데이터 로딩 진입부.", ""]
    lines += ["## 개발 도구와 검증 파일", ""]
    for folder in ("tools", "tests"):
        for path in sorted((project_root / folder).glob("*.py")):
            relative = path.relative_to(project_root).as_posix()
            tree = ast.parse(path.read_text(encoding="utf-8"))
            names = [symbol["name"] for symbol in iter_symbols(tree)]
            lines += [f"### [{relative}](../{relative})", "", (ast.get_docstring(tree) or "자동 검증 파일.").split("\n")[0], "", ", ".join(f"`{name}`" for name in names), ""]
    return "\n".join(lines).rstrip() + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    content = render_code_index()
    output = PROJECT_ROOT / "docs/CODE_INDEX.md"
    if args.check:
        if not output.exists() or output.read_text(encoding="utf-8") != content:
            raise SystemExit("Code index is stale: run python tools/update_code_index.py")
        print("Code index matches all source files and symbols")
    else:
        output.write_text(content, encoding="utf-8")
        print(output.relative_to(PROJECT_ROOT))


if __name__ == "__main__":
    main()
