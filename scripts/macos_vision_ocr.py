#!/usr/bin/env python3
"""OCR an image with macOS Vision.

This helper is intentionally tiny so the sender can prefer Apple's native OCR
for Chinese UI text on macOS, where Tesseract struggles with small WeCom titles.
"""

from __future__ import annotations

import sys
import subprocess
import tempfile
import os
from pathlib import Path


def ocr_with_swift(image_path: Path) -> int:
    swift_source = r'''
import Foundation
import Vision
import AppKit

if CommandLine.arguments.count < 2 {
    fputs("usage: vision_ocr.swift /path/to/image\n", stderr)
    exit(2)
}

let path = CommandLine.arguments[1]
guard let image = NSImage(contentsOfFile: path) else {
    fputs("failed to load image: \(path)\n", stderr)
    exit(3)
}

var proposed = CGRect(origin: .zero, size: image.size)
guard let cgImage = image.cgImage(forProposedRect: &proposed, context: nil, hints: nil) else {
    fputs("failed to create CGImage: \(path)\n", stderr)
    exit(3)
}

let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.recognitionLanguages = ["zh-Hans", "en-US"]
request.usesLanguageCorrection = false

let handler = VNImageRequestHandler(cgImage: cgImage, options: [:])
do {
    try handler.perform([request])
} catch {
    fputs("vision ocr failed: \(error)\n", stderr)
    exit(4)
}

let lines = (request.results ?? []).compactMap { $0.topCandidates(1).first?.string }.filter { !$0.isEmpty }
print(lines.joined(separator: "\n"))
'''
    with tempfile.TemporaryDirectory(prefix="vision_ocr_") as tmp:
        script_path = Path(tmp) / "vision_ocr.swift"
        script_path.write_text(swift_source, encoding="utf-8")
        env = os.environ.copy()
        env.setdefault("CLANG_MODULE_CACHE_PATH", "/private/tmp/codex_clang_module_cache")
        env.setdefault("TMPDIR", "/private/tmp")
        completed = subprocess.run(
            ["/usr/bin/xcrun", "swift", str(script_path), str(image_path)],
            capture_output=True,
            text=True,
            timeout=60,
            env=env,
        )
    if completed.stdout:
        print(completed.stdout.strip())
    if completed.returncode != 0 and completed.stderr:
        print(completed.stderr.strip(), file=sys.stderr)
    return completed.returncode


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: macos_vision_ocr.py /path/to/image", file=sys.stderr)
        return 2

    image_path = Path(sys.argv[1]).resolve()
    try:
        import Quartz
        import Vision
        from Cocoa import NSURL
    except Exception:
        return ocr_with_swift(image_path)

    image_url = NSURL.fileURLWithPath_(str(image_path))
    source = Quartz.CGImageSourceCreateWithURL(image_url, None)
    if source is None:
        print(f"failed to load image: {image_path}", file=sys.stderr)
        return 3
    cg_image = Quartz.CGImageSourceCreateImageAtIndex(source, 0, None)
    if cg_image is None:
        print(f"failed to create CGImage: {image_path}", file=sys.stderr)
        return 3

    request = Vision.VNRecognizeTextRequest.alloc().init()
    request.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    request.setRecognitionLanguages_(["zh-Hans", "en-US"])
    request.setUsesLanguageCorrection_(False)

    handler = Vision.VNImageRequestHandler.alloc().initWithCGImage_options_(cg_image, {})
    ok, error = handler.performRequests_error_([request], None)
    if not ok:
        print(f"vision ocr failed: {error}", file=sys.stderr)
        return 4

    lines: list[str] = []
    for observation in request.results() or []:
        candidates = observation.topCandidates_(1)
        if candidates:
            text = candidates[0].string()
            if text:
                lines.append(str(text))
    print("\n".join(lines))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
