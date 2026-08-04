"""命令行入口：支持文件、URL 与 stdin 管道。"""

from __future__ import annotations

import argparse
import codecs
import sys
from textwrap import dedent

from converter import (
    APP_VERSION,
    ConversionError,
    convert_local_path,
    convert_path,
    convert_stream,
    default_output_path,
)
from plugin_loader import format_plugin_list


def _parse_extension(raw: str | None) -> str | None:
    if raw is None:
        return None
    ext = raw.strip().lower()
    if not ext:
        return None
    return ext if ext.startswith(".") else f".{ext}"


def _parse_charset(raw: str | None) -> str | None:
    if raw is None:
        return None
    cs = raw.strip()
    if not cs:
        return None
    try:
        return codecs.lookup(cs).name
    except LookupError as exc:
        raise ConversionError(f"无效 charset：{cs}") from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="doc2md",
        description="将 PDF、Word、Excel、PPT、网页等转换为 Markdown。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=dedent(
            """
            示例：
              doc2md report.pdf
              doc2md report.pdf -o report.md
              type report.pdf | doc2md -x pdf
              doc2md --local-only memo.docx -o memo.md
              doc2md --keep-data-uris slides.pptx -o slides.md
              doc2md --list-plugins
              doc2md https://example.com/page.html -o page.md
            """
        ).strip(),
    )
    parser.add_argument("-v", "--version", action="version", version=f"%(prog)s {APP_VERSION}")
    parser.add_argument(
        "input",
        nargs="?",
        help="本地文件路径或 URL；省略则从 stdin 读取",
    )
    parser.add_argument("-o", "--output", help="输出 .md 文件；省略则写入 stdout")
    parser.add_argument(
        "-x",
        "--extension",
        help="stdin 模式下提供扩展名提示，如 pdf、docx",
    )
    parser.add_argument("-m", "--mime-type", help="stdin 模式下提供 MIME 类型提示")
    parser.add_argument("-c", "--charset", help="stdin 模式下提供字符集提示，如 UTF-8")
    parser.add_argument(
        "--local-only",
        action="store_true",
        help="窄接口模式：跳过 Excel / Notebook / ZIP 本地增强",
    )
    parser.add_argument(
        "--keep-data-uris",
        action="store_true",
        help="保留输出中的 base64 内嵌图片（默认截断以减小体积）",
    )
    parser.add_argument(
        "--list-plugins",
        action="store_true",
        help="列出已安装的第三方转换插件后退出",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="不输出进度信息（stderr）",
    )
    return parser


def _write_output(text: str, output: str | None) -> None:
    if output:
        with open(output, "w", encoding="utf-8") as fh:
            fh.write(text)
    else:
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        sys.stdout.write(text.encode(enc, errors="replace").decode(enc))


def run_cli(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.list_plugins:
        print(format_plugin_list())
        return 0

    def on_progress(msg: str) -> None:
        if not args.quiet:
            print(msg, file=sys.stderr)

    keep_uris = bool(args.keep_data_uris) or None

    try:
        ext = _parse_extension(args.extension)
        charset = _parse_charset(args.charset)
        mime = (args.mime_type or "").strip() or None

        if args.input is None:
            if sys.stdin.isatty():
                parser.print_help()
                return 2
            text = convert_stream(
                sys.stdin.buffer,
                extension=ext,
                mime_type=mime,
                charset=charset,
                progress=on_progress,
                keep_data_uris=keep_uris,
            )
            _write_output(text, args.output)
            return 0

        source = args.input.strip()
        if args.local_only and not source.startswith(("http://", "https://")):
            text = convert_local_path(source, progress=on_progress, keep_data_uris=keep_uris)
        else:
            text = convert_path(
                source,
                progress=on_progress,
                local_only=args.local_only,
                keep_data_uris=keep_uris,
            )

        out = args.output
        if not out and not source.startswith(("http://", "https://")):
            out = str(default_output_path(source))
        _write_output(text, out)
        return 0
    except ConversionError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("已取消。", file=sys.stderr)
        return 130


def main() -> None:
    raise SystemExit(run_cli())


if __name__ == "__main__":
    main()
