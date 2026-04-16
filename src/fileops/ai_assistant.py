from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from urllib import error, request

from .models import OperationResult


AI_PROVIDER_CATALOG: dict[str, dict[str, Any]] = {
    "chatgpt": {
        "label": "ChatGPT (OpenAI)",
        "api_style": "openai",
        "base_url": "https://api.openai.com/v1",
        "models": ["gpt-4.1-mini", "gpt-4.1", "gpt-4o-mini"],
    },
    "deepseek": {
        "label": "DeepSeek",
        "api_style": "openai",
        "base_url": "https://api.deepseek.com/v1",
        "models": ["deepseek-chat", "deepseek-reasoner"],
    },
    "glm": {
        "label": "智谱 GLM",
        "api_style": "openai",
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "models": ["glm-4-plus", "glm-4-air", "glm-4-flash"],
    },
    "claude": {
        "label": "Claude (Anthropic)",
        "api_style": "anthropic",
        "base_url": "https://api.anthropic.com/v1",
        "models": ["claude-3-7-sonnet-latest", "claude-3-5-haiku-latest"],
    },
    "kimi": {
        "label": "Kimi (Moonshot)",
        "api_style": "openai",
        "base_url": "https://api.moonshot.cn/v1",
        "models": ["moonshot-v1-8k", "moonshot-v1-32k", "moonshot-v1-128k"],
    },
}


def list_ai_providers() -> list[tuple[str, str]]:
    return [(provider, str(spec["label"])) for provider, spec in AI_PROVIDER_CATALOG.items()]


def list_models_for_provider(provider: str) -> list[str]:
    spec = AI_PROVIDER_CATALOG.get(provider, {})
    models = spec.get("models", [])
    return [str(model) for model in models]


def normalize_ai_config(config: dict[str, Any]) -> dict[str, Any]:
    api_key = str(config.get("api_key") or "").strip()
    if not api_key:
        raise ValueError("AI API key is required.")

    provider = str(config.get("provider") or "").strip().lower()
    model = str(config.get("model") or "").strip()

    if not provider:
        provider = _infer_provider_from_model(model)
    if provider not in AI_PROVIDER_CATALOG:
        supported = ", ".join(AI_PROVIDER_CATALOG.keys())
        raise ValueError(f"Unsupported AI provider: {provider}. Supported: {supported}")

    spec = AI_PROVIDER_CATALOG[provider]
    if not model:
        model = str(spec["models"][0])

    base_url = str(config.get("base_url") or spec["base_url"]).strip().rstrip("/")
    timeout = float(config.get("timeout", 60))
    max_items = int(config.get("max_items", 30))

    return {
        "provider": provider,
        "provider_label": str(spec["label"]),
        "api_style": str(spec["api_style"]),
        "base_url": base_url,
        "model": model,
        "api_key": api_key,
        "timeout": timeout,
        "max_items": max_items,
    }


def generate_compare_ai_report(analysis: dict[str, Any], output_path: Path, config: dict[str, Any]) -> Path:
    resolved = normalize_ai_config(config)
    max_items = max(1, int(resolved["max_items"]))
    prompt = _build_compare_prompt(analysis, max_items=max_items)
    content = _request_ai_text(prompt=prompt, resolved_config=resolved)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8-sig")
    return output_path


def generate_operation_ai_report(
    operation: str,
    source: Path,
    results: list[OperationResult],
    output_path: Path,
    config: dict[str, Any],
) -> Path:
    resolved = normalize_ai_config(config)
    prompt = _build_operation_prompt(operation=operation, source=source, results=results)
    content = _request_ai_text(prompt=prompt, resolved_config=resolved)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8-sig")
    return output_path


def _build_compare_prompt(analysis: dict[str, Any], max_items: int) -> dict[str, str]:
    overview = analysis.get("overview", {})
    summary = analysis.get("summary", {})
    issues = list(analysis.get("issues", []))[:max_items]
    issue_brief = [
        {
            "id": item.get("id"),
            "severity": item.get("severity_label"),
            "category": item.get("category_label"),
            "location": item.get("location"),
            "detail": item.get("detail"),
            "expected": item.get("expected"),
            "actual": item.get("actual"),
            "adjustment": item.get("adjustment"),
        }
        for item in issues
    ]

    system_prompt = (
        "你是论文格式修订专家。请根据检测结果生成可执行、可核验的修订建议。"
        "输出必须是中文 Markdown，结构包括："
        "1) 总体评估 2) 高优先级问题 3) 分类别修复步骤 4) 最终复检清单。"
    )
    user_prompt = (
        f"模板：{overview.get('template_name', '')}\n"
        f"文档：{overview.get('source_name', '')}\n"
        f"状态：{overview.get('status_text', '')}\n"
        f"统计：{json.dumps(summary, ensure_ascii=False)}\n\n"
        f"问题明细（节选）：\n{json.dumps(issue_brief, ensure_ascii=False, indent=2)}\n"
    )
    return {"system": system_prompt, "user": user_prompt}


def _build_operation_prompt(operation: str, source: Path, results: list[OperationResult]) -> dict[str, str]:
    status_counts: dict[str, int] = {}
    result_items: list[dict[str, str]] = []

    for item in results:
        key = item.status.value
        status_counts[key] = status_counts.get(key, 0) + 1
        result_items.append(
            {
                "status": key,
                "source": item.source,
                "destination": item.destination or "",
                "message": item.message,
            }
        )

    system_prompt = (
        "你是文档处理流程优化助手。请根据执行结果输出实操建议。"
        "输出必须是中文 Markdown，结构包括："
        "1) 执行结果摘要 2) 风险与问题定位 3) 下一轮参数建议 4) 操作清单。"
    )
    user_prompt = (
        f"操作类型：{operation}\n"
        f"源文件：{source}\n"
        f"状态统计：{json.dumps(status_counts, ensure_ascii=False)}\n"
        f"结果详情：\n{json.dumps(result_items, ensure_ascii=False, indent=2)}\n"
    )
    return {"system": system_prompt, "user": user_prompt}


def _request_ai_text(prompt: dict[str, str], resolved_config: dict[str, Any]) -> str:
    api_style = str(resolved_config["api_style"])
    if api_style == "anthropic":
        payload, endpoint, headers = _build_anthropic_request(prompt=prompt, config=resolved_config)
    else:
        payload, endpoint, headers = _build_openai_request(prompt=prompt, config=resolved_config)

    req = request.Request(endpoint, data=json.dumps(payload, ensure_ascii=False).encode("utf-8"), method="POST")
    for header, value in headers.items():
        req.add_header(header, value)

    try:
        with request.urlopen(req, timeout=float(resolved_config["timeout"])) as response:  # noqa: S310
            body = response.read().decode("utf-8")
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"AI request failed: HTTP {exc.code} {detail}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"AI request failed: {exc.reason}") from exc

    response_payload = json.loads(body)
    text = _extract_response_text(response_payload, api_style=api_style).strip()
    if not text:
        raise RuntimeError("AI response did not contain usable content.")
    return text


def _build_openai_request(prompt: dict[str, str], config: dict[str, Any]) -> tuple[dict[str, Any], str, dict[str, str]]:
    endpoint = f"{str(config['base_url']).rstrip('/')}/chat/completions"
    payload = {
        "model": config["model"],
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": prompt["system"]},
            {"role": "user", "content": prompt["user"]},
        ],
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {config['api_key']}",
    }
    return payload, endpoint, headers


def _build_anthropic_request(prompt: dict[str, str], config: dict[str, Any]) -> tuple[dict[str, Any], str, dict[str, str]]:
    endpoint = f"{str(config['base_url']).rstrip('/')}/messages"
    payload = {
        "model": config["model"],
        "max_tokens": 1400,
        "temperature": 0.2,
        "system": prompt["system"],
        "messages": [
            {
                "role": "user",
                "content": prompt["user"],
            }
        ],
    }
    headers = {
        "Content-Type": "application/json",
        "x-api-key": str(config["api_key"]),
        "anthropic-version": "2023-06-01",
    }
    return payload, endpoint, headers


def _extract_response_text(payload: dict[str, Any], api_style: str) -> str:
    if api_style == "anthropic":
        content = payload.get("content", [])
        if isinstance(content, list):
            parts: list[str] = []
            for item in content:
                if isinstance(item, dict) and item.get("type") == "text":
                    text = item.get("text")
                    if isinstance(text, str):
                        parts.append(text)
            return "\n".join(parts)
        return ""

    choices = payload.get("choices", [])
    if not choices:
        return ""
    message = choices[0].get("message", {})
    content = message.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts: list[str] = []
        for item in content:
            if isinstance(item, dict):
                item_text = item.get("text")
                if isinstance(item_text, str):
                    parts.append(item_text)
        return "\n".join(parts)
    return ""


def _infer_provider_from_model(model: str) -> str:
    lowered = model.lower()
    if lowered.startswith("deepseek"):
        return "deepseek"
    if lowered.startswith("glm"):
        return "glm"
    if lowered.startswith("claude"):
        return "claude"
    if lowered.startswith("moonshot") or lowered.startswith("kimi"):
        return "kimi"
    return "chatgpt"
