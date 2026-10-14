"""Run model requests through the installed Claude Code CLI (`claude -p`) instead of an API endpoint.

Models appear in the GUI as "Claude Code: <model>". Each request is one headless `claude -p` call with
streaming JSON in and out: the prompt and page images go in on stdin as content blocks, so Claude needs no
file or shell tools, and the answer streams back token by token. The CLI uses the user's existing Claude
login, so no API key is involved. Each request is its own session (nothing persisted, nothing carried over
between pages).
"""
import base64
import io
import json
import shutil
import subprocess
import sys
import threading
from pathlib import Path
from typing import Callable, List, Optional

PREFIX = "Claude Code: "
# Aliases the CLI resolves itself, then explicit model ids. "default" means: don't pass --model.
MODEL_CHOICES = [
    "default", "sonnet", "opus", "haiku",
    "claude-fable-5-1", "claude-opus-5-5", "claude-sonnet-5-5", "claude-haiku-5-5",
]
SYSTEM_PROMPT = "You are an engineering document analyst. Extract the PLC-connected instruments and return valid JSON."
REQUEST_TIMEOUT_SECONDS = 1200
MAX_IMAGE_BYTES = 3_500_000  # larger pages are re-encoded as JPEG; the API caps images at 5 MB (base64 adds a third)


_active: set = set()  # running `claude` processes, so closing the app can stop them
_active_lock = threading.Lock()


def kill_all() -> None:
    """Stop any in-flight Claude Code requests (used when the window is closed)."""
    with _active_lock:
        for process in list(_active):
            try:
                process.kill()
            except OSError:
                pass


def is_claude_code_model(model_name: Optional[str]) -> bool:
    return bool(model_name) and model_name.startswith(PREFIX)


def cli_path() -> Optional[str]:
    return shutil.which("claude")


def available_models() -> List[str]:
    """The GUI model names offered when the Claude Code CLI is installed (empty list otherwise)."""
    return [PREFIX + name for name in MODEL_CHOICES] if cli_path() else []


def _image_block(image_path: Path) -> dict:
    data = Path(image_path).read_bytes()
    media_type = "image/png"
    if len(data) > MAX_IMAGE_BYTES:
        from PIL import Image
        with Image.open(io.BytesIO(data)) as im:
            im = im.convert("RGB")
            im.thumbnail((4000, 4000))
            buffer = io.BytesIO()
            im.save(buffer, format="JPEG", quality=88)
            data, media_type = buffer.getvalue(), "image/jpeg"
    return {"type": "image", "source": {"type": "base64", "media_type": media_type,
                                        "data": base64.b64encode(data).decode("ascii")}}


def send_request(
    model: str,
    prompt: str,
    image_paths: List[Path],
    docling_markdown: Optional[str] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    stream_callback: Optional[Callable[[str], None]] = None,
) -> str:
    """Send one prompt (plus images) to Claude Code and return the final answer text."""
    log = log_callback or (lambda _m: None)
    stream = stream_callback or (lambda _t: None)

    exe = cli_path()
    if not exe:
        raise RuntimeError("The Claude Code CLI ('claude') was not found on PATH. Install it or pick another model.")

    cli_model = model[len(PREFIX):].strip() if is_claude_code_model(model) else model
    command = [
        exe, "-p",
        "--input-format", "stream-json", "--output-format", "stream-json",
        "--verbose", "--include-partial-messages",
        "--no-session-persistence", "--disable-slash-commands",
        "--system-prompt", SYSTEM_PROMPT,
        "--tools", "",  # the page is in the message itself; Claude needs no tools
    ]
    if cli_model and cli_model != "default":
        command += ["--model", cli_model]

    text = prompt
    if docling_markdown:
        text = f"{prompt}\n\n---\nStructured document context (Docling extraction):\n{docling_markdown}"
    content = [{"type": "text", "text": text}] + [_image_block(p) for p in image_paths]
    message = json.dumps({"type": "user", "message": {"role": "user", "content": content}}) + "\n"

    log(f"[CLAUDE CODE] {cli_model or 'default'}: sending {len(image_paths)} image(s), {len(text)} characters of text")
    process = subprocess.Popen(
        command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0,
    )
    with _active_lock:
        _active.add(process)
    timer = threading.Timer(REQUEST_TIMEOUT_SECONDS, process.kill)
    timer.start()
    stderr_lines: List[str] = []
    stderr_thread = threading.Thread(target=lambda: stderr_lines.extend(process.stderr), daemon=True)
    stderr_thread.start()

    streamed = False
    final_text: Optional[str] = None
    error: Optional[str] = None
    try:
        try:
            process.stdin.write(message)
            process.stdin.close()
        except OSError:
            pass  # the process already exited; its stderr/stdout explain why

        for line in process.stdout:
            line = line.strip()
            if not line.startswith("{"):
                continue
            try:
                event = json.loads(line)
            except ValueError:
                continue
            kind = event.get("type")
            if kind == "stream_event":
                delta = (event.get("event") or {}).get("delta") or {}
                if delta.get("type") == "text_delta" and delta.get("text"):
                    streamed = True
                    stream(delta["text"])
            elif kind == "result":
                if event.get("is_error"):
                    error = str(event.get("result") or event.get("subtype") or "unknown error")
                else:
                    final_text = str(event.get("result") or "")
        process.wait()
    finally:
        timer.cancel()
        with _active_lock:
            _active.discard(process)
        if process.poll() is None:
            process.kill()
        stderr_thread.join(timeout=2)

    if error:
        raise RuntimeError(f"Claude Code reported an error: {error}")
    if final_text is None:
        detail = "".join(stderr_lines).strip()[-500:]
        raise RuntimeError(f"Claude Code returned no result (exit code {process.returncode}). {detail}")
    if not streamed:
        stream(final_text)
    return final_text.strip()
