from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.request
from typing import Any


def enabled(config: dict[str, str] | None) -> bool:
    return bool(config and str(config.get("enabled", "")).lower() == "true")


def sanitize_status(status: dict[str, Any]) -> dict[str, Any]:
    return {
        "enabled": bool(status.get("enabled")),
        "status": status.get("status", ""),
        "message": status.get("message", ""),
        "model": status.get("model", ""),
        "baseUrl": status.get("baseUrl", ""),
        "elapsedMs": status.get("elapsedMs", 0),
    }


def validate_config(config: dict[str, str] | None) -> tuple[bool, str]:
    if not enabled(config):
        return False, "未启用大模型增强，使用本地规则分析"
    missing = []
    for key, label in [("baseUrl", "API Base URL"), ("model", "Model Name"), ("apiKey", "API Key")]:
        if not str(config.get(key, "")).strip():
            missing.append(label)
    if missing:
        return False, "大模型配置不完整：" + "、".join(missing)
    if str(config.get("protocol", "openai-compatible")) != "openai-compatible":
        return False, "当前版本仅支持 OpenAI Compatible 接口协议"
    return True, ""


def chat_completions(config: dict[str, str], messages: list[dict[str, str]], timeout: int = 60, retries: int = 1) -> dict[str, Any]:
    ok, message = validate_config(config)
    if not ok:
        return {
            "enabled": enabled(config),
            "status": "disabled" if not enabled(config) else "failed",
            "message": message,
            "model": str(config.get("model", "")),
            "baseUrl": str(config.get("baseUrl", "")),
            "elapsedMs": 0,
        }

    base_url = str(config["baseUrl"]).rstrip("/")
    url = base_url if base_url.endswith("/chat/completions") else f"{base_url}/v1/chat/completions"
    payload = {
        "model": str(config["model"]).strip(),
        "messages": messages,
        "temperature": 0.2,
        "response_format": {"type": "json_object"},
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {
        "Authorization": f"Bearer {config['apiKey']}",
        "Content-Type": "application/json",
    }

    started = time.time()
    last_error = ""
    for attempt in range(retries + 1):
        try:
            request = urllib.request.Request(url, data=body, headers=headers, method="POST")
            with urllib.request.urlopen(request, timeout=timeout) as response:
                raw = response.read().decode("utf-8", errors="replace")
            data = json.loads(raw)
            content = data["choices"][0]["message"]["content"]
            parsed = parse_json_object(content)
            return {
                "enabled": True,
                "status": "success",
                "message": "模型增强成功",
                "model": payload["model"],
                "baseUrl": base_url,
                "elapsedMs": int((time.time() - started) * 1000),
                "data": parsed,
            }
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
            last_error = f"HTTP {exc.code}: {detail}"
        except Exception as exc:
            last_error = str(exc)
        if attempt < retries:
            time.sleep(1)

    return {
        "enabled": True,
        "status": "failed",
        "message": f"模型增强失败，已回退本地规则：{last_error}",
        "model": payload["model"],
        "baseUrl": base_url,
        "elapsedMs": int((time.time() - started) * 1000),
    }


def parse_json_object(text: str) -> dict[str, Any]:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text, flags=re.I).strip()
        text = re.sub(r"```$", "", text).strip()
    try:
        data = json.loads(text)
        return data if isinstance(data, dict) else {"raw": data}
    except Exception:
        match = re.search(r"\{.*\}", text, flags=re.S)
        if not match:
            return {"raw": text}
        data = json.loads(match.group(0))
        return data if isinstance(data, dict) else {"raw": data}
