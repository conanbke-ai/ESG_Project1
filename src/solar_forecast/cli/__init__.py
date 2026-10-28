"""  init  : CLI input translation and command dispatch."""
from __future__ import annotations

from typing import Sequence
from solar_forecast.cli.parser import build_parser


def main(argv: Sequence[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    args.func(args)

__all__ = ["build_parser", "main"]
