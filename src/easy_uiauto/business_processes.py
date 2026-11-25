"""Business goals derived at import, with coverage computed from the workflow catalog."""
from pathlib import Path
import re
import unicodedata

from . import workflows
from .mcp import image_storage, knowledge


def from_topics(topics, reference):
    records = []
    for topic in topics:
        steps = [s for s in topic.get("steps", []) if s.strip() not in {"", "未说明"}]
        if not steps:
            continue
        title = (topic.get("process_title") or topic["title"]).strip()
        if not re.fullmatch(r"【[^【】\r\n]+】[^【】\r\n]+-[^【】\r\n]+", title):
            entry = topic.get("entry", "").strip()
            if entry in {"", "未说明"}:
                entry = "资料入口"
            feature = " ".join(topic["title"].replace("【", "").replace("】", "").split())
            branch = " ".join(entry.replace("【", "").replace("】", "").split())
            title = f"【业务操作】{feature}-{branch}"
        records.append(dict(id=knowledge.stable_id("process", _key(title)), name=title,
            aliases=[topic["title"], *topic.get("aliases", [])], description=topic["description"],
            entry=topic.get("entry", ""), preconditions=topic.get("preconditions", []),
            steps=steps, expected=topic.get("expected", []),
            sources=[dict(id=reference["id"], name=reference["name"],
                          path=f"references/{reference['id']}.md",
                          sections=topic.get("source_ids", []), quote=topic["quote"])]))
    return records


def _key(text):
    return "".join(unicodedata.normalize("NFKC", text).casefold().split())


def records(directory, controls, catalog):
    directory = Path(directory)
    combined = {}
    for path in sorted((directory / "business-processes").glob("*.md")):
        if image_storage._safe_path(directory, path) is None:
            raise ValueError("业务流程路径不能为链接")
        batch, _ = knowledge.read_markdown(path)
        for stored in batch["processes"]:
            record = combined.setdefault(stored["id"], dict(stored, sources=[], aliases=[]))
            record["sources"].extend(s for s in stored["sources"] if s not in record["sources"])
            record["aliases"] = list(dict.fromkeys([*record["aliases"], *stored["aliases"]]))
    if not combined:
        return []
    feature_counts = {}
    for record in combined.values():
        for alias in {_key(a) for a in record["aliases"]}:
            feature_counts[alias] = feature_counts.get(alias, 0) + 1
    for record in combined.values():
        names = {_key(record["name"])}
        names.update(_key(a) for a in record["aliases"] if feature_counts[_key(a)] == 1)
        matches = []
        for flow in catalog.values():
            names_in_catalog = {re.sub(r"\(\d+\)$", "", _key(name))
                                for name in (flow.get("name", ""), flow["keyword"])}
            title_key = _key(record["name"])
            branch_matches = any(name.startswith(title_key + "/") for name in names_in_catalog)
            linked = flow.get("business_process_id") == record["id"]
            if (linked or branch_matches or names & names_in_catalog) and workflows.readiness(flow, controls, catalog)["ready"]:
                steps = workflows.expand_calls(flow, catalog)["steps"]
                if any(s["action"] not in {"inspect", "manual"} for s in steps):
                    matches.append(dict(id=flow["id"], name=flow.get("name") or flow["keyword"]))
        record.update(status="已学习" if matches else "待学习", workflows=matches)
    return sorted(combined.values(), key=lambda record: record["name"])


def display(record):
    lines = [record["name"], record["status"], "", record["description"]]
    for heading, items in (("入口", [record["entry"]] if record["entry"] else []),
                           ("前置条件", record["preconditions"]), ("操作步骤", record["steps"]),
                           ("预期结果", record["expected"]),
                           ("匹配的组合动作", [flow["name"] for flow in record["workflows"]]),
                           ("参考来源", [source["name"] for source in record["sources"]])):
        if items:
            lines.extend(["", heading, *[f"{i}. {item}" for i, item in enumerate(items, 1)]])
    return "\n".join(lines)
