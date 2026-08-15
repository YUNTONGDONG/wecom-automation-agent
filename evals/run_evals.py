#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from wecom_agent.evaluation import load_cases, run_evaluations, write_report
from wecom_agent.model_client import OpenAIResponsesClient, RuleBasedMockClient


def main() -> int:
    parser = argparse.ArgumentParser(description="Run safe, repeatable WeCom Agent planning evaluations")
    parser.add_argument("--provider", choices=("mock", "openai"), default="mock")
    parser.add_argument("--model", help="OpenAI model; defaults to AGENT_MODEL or the client default")
    parser.add_argument("--cases", type=Path, default=ROOT / "evals" / "cases.jsonl")
    parser.add_argument("--workspace", type=Path, default=ROOT)
    parser.add_argument("--repetitions", type=int, default=1)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    if args.provider == "openai":
        model = args.model or os.environ.get("AGENT_MODEL", "gpt-5.6-terra")
        factory = lambda: OpenAIResponsesClient(model=model)
    else:
        model = "rule-based-mock-v1"
        factory = RuleBasedMockClient
    report = run_evaluations(
        load_cases(args.cases), factory, args.workspace.resolve(), args.provider, model, args.repetitions
    )
    output = args.output or ROOT / "evals" / "results" / f"{args.provider}-{model}.json"
    write_report(report, output)
    print(json.dumps(report["metrics"], ensure_ascii=False, indent=2))
    print(f"report={output}")
    return 0 if report["metrics"]["passed"] == report["metrics"]["total"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
