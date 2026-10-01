#!/usr/bin/env python
"""Generate a complete monthly DOCX delivery from one checklist workbook."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

from build_docx_from_spec import (
    build_document,
    import_dependencies as import_docx_dependencies,
    validate_document_spec,
)
from build_standard_template import build_template
from plan_monthly_docs import (
    PLACEHOLDER_RE,
    as_text,
    branch_label,
    extract_activities,
    import_openpyxl,
    normalize_title,
    ordinal_number,
)
from validate_delivery import validate_delivery


MONTH_RE = re.compile(r"(\d{1,2})月")
YEAR_RE = re.compile(r"(20\d{2})年?")
ACTUAL_BRANCH_RE = re.compile(r"[\u4e00-\u9fff]{2,40}党支部")
CONTENT_PREFIX_RE = re.compile(
    r"^(?:"
    r"学习“第一议题”|集中学习内容|集中学习|学习内容|"
    r"主题党日|谈心谈话内容|谈话内容|党课内容"
    r")[:：]\s*"
)
DEFAULT_CONFIG_NAME = "party-building.local.json"
STANDARD_TEMPLATE_NAME = "standard-party-template.docx"


def configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def skill_root() -> Path:
    return Path(__file__).resolve().parents[1]


def default_config_path() -> Path:
    return skill_root() / DEFAULT_CONFIG_NAME


def default_template_path() -> Path:
    return skill_root() / "templates" / STANDARD_TEMPLATE_NAME


def load_local_config(path: Path | None) -> dict:
    config_path = (path or default_config_path()).resolve()
    if not config_path.exists():
        return {}
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise SystemExit(f"Cannot read branch config: {exc}") from exc
    if not isinstance(data, dict):
        raise SystemExit("Branch config must be a JSON object.")
    return data


def save_local_config(path: Path | None, branch: str) -> Path:
    config_path = (path or default_config_path()).resolve()
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(
        json.dumps({"branch_name": branch}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return config_path


def workbook_text_values(workbook) -> list[str]:
    values = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                text = as_text(cell.value)
                if text:
                    values.append(text)
    return values


def infer_month(checklist_path: Path, values: list[str]) -> int:
    path_match = MONTH_RE.search(checklist_path.name)
    if path_match:
        month = int(path_match.group(1))
        if not 1 <= month <= 12:
            raise SystemExit(f"Month is outside 1 to 12: {month}")
        return month
    counts = Counter(
        int(match)
        for value in values
        for match in MONTH_RE.findall(value)
        if 1 <= int(match) <= 12
    )
    if not counts:
        raise SystemExit(
            "Could not infer the target month from the checklist filename or cells."
        )
    return counts.most_common(1)[0][0]


def infer_year(checklist_path: Path, values: list[str]) -> int:
    for text in [checklist_path.name, *values]:
        match = YEAR_RE.search(text)
        if match:
            return int(match.group(1))
    return 2026


def infer_branch(values: list[str]) -> str | None:
    exact_candidates = []
    embedded_candidates = []
    for value in values:
        match = ACTUAL_BRANCH_RE.search(value)
        if not match:
            continue
        candidate = match.group().strip()
        if PLACEHOLDER_RE.search(candidate) or "各支部" in candidate:
            continue
        if value == candidate:
            exact_candidates.append(candidate)
        else:
            embedded_candidates.append(candidate)
    candidates = exact_candidates or embedded_candidates
    if not candidates:
        return None
    return Counter(candidates).most_common(1)[0][0]


def resolve_branch(
    requested: str | None,
    config: dict,
    checklist_values: list[str],
) -> tuple[str, str]:
    if requested:
        return branch_label(requested), "command"
    configured = as_text(config.get("branch_name"))
    if configured:
        return branch_label(configured), "config"
    inferred = infer_branch(checklist_values)
    if inferred:
        return branch_label(inferred), "checklist"
    raise SystemExit(
        "The checklist uses a generic xx党支部 placeholder. Run once with "
        "--branch-name <name> --save-config, then future runs only need "
        "--checklist."
    )


def normalize_participants(value: str, branch: str) -> str:
    text = as_text(value)
    if not text:
        return f"{branch}全体党员"
    text = text.replace("各支部全体党员", f"{branch}全体党员")
    text = text.replace("各支部", branch)
    return text


def clean_content(value: str) -> str:
    text = re.sub(r"[\r\n]+", " ", value)
    text = re.sub(r"\s+", " ", text).strip()
    previous = None
    while text and text != previous:
        previous = text
        text = CONTENT_PREFIX_RE.sub("", text).strip()
    return text


def extract_theme(raw_title: str, content: str, category: str) -> str:
    title_text = as_text(raw_title)
    content_text = clean_content(content)
    quoted = re.findall(r"[“\"]([^”\"]+)[”\"]", title_text)
    if quoted and not PLACEHOLDER_RE.search(quoted[0]):
        return quoted[0].strip()
    if "主题党日" in content_text:
        content_text = re.sub(r"^主题党日[:：]?\s*", "", content_text)
    if title_text and not PLACEHOLDER_RE.search(title_text):
        title_theme = re.sub(r"^x{1,}(?:支行)?党支部", "", title_text, flags=re.I)
        title_theme = re.sub(r"^\d{1,2}月", "", title_theme)
        title_theme = re.sub(r"主题党日.*$", "", title_theme).strip()
        if title_theme:
            return title_theme
    if category == "主题党日":
        return content_text.split("。", 1)[0].strip() or "主题党日活动"
    return content_text.split("。", 1)[0].strip() or "专题学习"


def learning_topic(activity: dict) -> str:
    content = clean_content(activity["content"])
    if content:
        return content
    title = as_text(activity["raw_title"])
    return title or activity["category"]


def date_from_due(due: str, month: int, year: int) -> str:
    match = re.search(r"(\d{1,2})月(\d{1,2})日", as_text(due))
    if match:
        day = int(match.group(2))
    else:
        day_match = re.search(r"(\d{1,2})日", as_text(due))
        day = int(day_match.group(1)) if day_match else 1
    return f"{year}年{month}月{max(1, min(day, 28))}日"


def extract_name_list(activity: dict) -> list[str]:
    for value in activity.get("extra_values", []):
        parts = [
            part.strip()
            for part in re.split(r"[，,、;；\s]+", as_text(value))
            if part.strip()
        ]
        if len(parts) >= 3 and all(
            2 <= len(part) <= 4
            and re.fullmatch(r"[\u4e00-\u9fff]+", part)
            for part in parts
        ):
            return parts
    return []


def discussion_speakers(activity: dict, topic: str) -> list[str]:
    focus = topic[:48]
    names = extract_name_list(activity)
    labels = names[:3] if len(names) >= 3 else [
        "支部书记",
        "组织委员",
        "党员代表",
    ]
    statements = [
        (
            "要把政治建设摆在首位，严格落实“第一议题”和“三会一课”"
            f"等制度，围绕{focus}先学一步、学深一层，推动党建工作与"
            "业务工作同谋划、同部署。"
        ),
        (
            "要增强学习的针对性和实效性，把清单要求转化为具体的学习"
            "安排和实践任务，联系岗位谈认识、找差距、明措施，避免学习"
            "记录停留在简单摘抄和原则表态上。"
        ),
        (
            "要把学习成果落实到日常服务中，主动深入基层、了解群众需求，"
            "认真办好每一笔业务、回应合理诉求，在遵规守纪、履职尽责中"
            "体现党员担当。"
        ),
    ]
    return [
        f"{label}：{statement}"
        for label, statement in zip(labels, statements)
    ]


def short_record_paragraphs(activity: dict, branch: str, topic: str) -> list[dict]:
    category = activity["category"]
    participants = normalize_participants(activity["participants"], branch)
    if "支委会" in category or "党员大会" in category:
        opening = (
            f"会议由支部书记主持，{participants}参加。会议严格落实“第一议题”"
            f"制度，围绕{topic}开展集体学习。"
        )
    else:
        opening = f"{participants}集中学习了{topic}。"

    learning_text = (
        opening
        + "学习坚持读原文、学原著、悟原理，重点把握核心要义、实践要求"
        "和与基层工作的结合点。大家一致认为，要坚持学思用贯通、"
        "知信行统一，把政治要求落实到业务经营、基层服务、风险防控和"
        "队伍建设中，把学习成果转化为履职尽责、服务群众和推动发展的"
        "实际行动。"
    )
    paragraphs = [
        {"style": "党建标题1", "text": "学习内容记录："},
        {"style": "党建正文", "text": learning_text},
        {"style": "党建标题1", "text": "讨论发言摘要："},
    ]
    paragraphs.extend(
        {"style": "党建正文", "text": text}
        for text in discussion_speakers(activity, topic)
    )
    return paragraphs


def theme_day_paragraphs(
    activity: dict,
    branch: str,
    month: int,
    year: int,
    topic: str,
    theme: str,
) -> list[dict]:
    participants = normalize_participants(activity["participants"], branch)
    activity_date = date_from_due(activity["due"], month, year)
    title = f"{branch}{month}月主题党日活动记录"
    metadata = [
        ("活动时间", activity_date),
        ("活动地点", f"{branch}党员活动室及服务一线"),
        ("参与单位", branch),
        ("参与人员", participants),
        ("活动主题", f"“{theme}”主题党日"),
    ]
    paragraphs = [{"style": "党建标题1", "text": title}]
    paragraphs.extend(
        {
            "style": "党建正文",
            "text": f"{label}：{value}",
            "bold": True,
            "first_line_indent_pt": 0,
            "line_spacing_pt": 31,
        }
        for label, value in metadata
    )

    background = (
        f"{month}月，{branch}围绕“{theme}”开展主题党日活动。活动坚持"
        "理论学习、党性锻炼和服务实践相结合，教育引导党员把初心使命"
        "落实到服务基层、服务群众和推动发展的具体行动中。"
    )
    learning = (
        f"活动中，全体党员集中学习了{topic}。学习紧扣清单确定的专题和"
        "实践要求，引导党员联系岗位职责谈认识、谈体会、谈打算，进一步"
        "统一思想、凝聚共识。"
    )

    if any(term in theme for term in ("中秋", "国庆", "节日")):
        practice = (
            "结合节日节点和群众实际需求，党员分组开展走访服务和宣传，"
            "围绕普惠金融、防范诈骗、便民业务等内容讲解政策、收集需求，"
            "对能够现场解决的问题立即办理，对需要后续处理的事项登记台账，"
            "明确责任人和反馈时间。"
        )
    elif any(term in theme for term in ("志愿", "服务", "为民", "雷锋")):
        practice = (
            "党员志愿者结合群众需求设置服务点并开展走访，宣传普惠金融、"
            "消费者权益保护和风险防范知识，认真听取群众对服务流程、"
            "工作作风和产品办理的意见，建立问题清单并逐项跟进。"
        )
    elif any(term in theme for term in ("安全", "反诈", "合规", "风险")):
        practice = (
            "活动围绕安全发展和风险防控开展实践，组织党员排查服务环节"
            "中的风险点，向群众宣讲防范非法集资、电信网络诈骗和账户安全"
            "知识，帮助群众增强风险识别能力，守好资金安全底线。"
        )
    elif any(term in theme for term in ("廉洁", "清廉", "纪律")):
        practice = (
            "活动把廉洁教育、警示提醒和岗位实践结合起来，组织党员对照"
            "纪律规矩和业务制度查找风险点，围绕客户服务、业务办理和"
            "日常管理明确改进措施，推动形成知敬畏、守底线、重实干的"
            "良好作风。"
        )
    else:
        practice = (
            "活动结合近期重点工作和群众需求开展实践服务，组织党员深入"
            "一线了解情况、宣传政策、收集意见。对群众反映的问题建立"
            "清单，能够立即解决的现场办理，需要协调推进的明确措施和"
            "反馈节点，确保活动取得实际效果。"
        )

    reflection = (
        "在交流中，党员们表示，主题党日既是一次理论学习，也是一次"
        "党性锻炼。要把群众满意作为检验工作的根本标准，主动改进服务"
        "方式和工作作风，把学习成果转化为解决问题、推动落实的能力。"
    )
    effect = (
        "活动效果：本次主题党日把理论学习、党性教育和实践服务贯通起来，"
        "进一步强化了党员的宗旨意识和责任意识，也拉近了与群众的距离。"
        "全体党员将持续跟进走访和活动中收集的问题，把“问题清单”转化为"
        "履职清单，以更加务实的作风推动各项工作落地见效。"
    )
    paragraphs.extend(
        [
            {"style": "党建正文", "text": background},
            {"style": "党建正文", "text": learning},
            {"style": "党建正文", "text": practice},
            {
                "style": "党建正文",
                "text": reflection,
                "page_break_before": True,
            },
            {"style": "党建正文", "text": effect},
        ]
    )
    return paragraphs


def talk_paragraphs(
    activity: dict,
    branch: str,
    month: int,
    year: int,
    topic: str,
) -> list[dict]:
    activity_date = date_from_due(activity["due"], month, year)
    learning_text = (
        f"全体党员集中学习了{topic}。学习聚焦理论学习、岗位实践和"
        "服务群众的方法要求，引导党员把学习成果转化为解决实际问题、"
        "推动工作落实的具体行动。"
    )
    paragraphs = [
        {"style": "党建标题1", "text": "学习内容记录："},
        {"style": "党建正文", "text": learning_text},
        {
            "style": "党建正文",
            "text": (
                "集中学习结束后，支部书记围绕学习收获、岗位履职和成长"
                "困惑，与党员代表开展谈心谈话。"
            ),
        },
        {"style": "党建标题1", "text": "谈心谈话记录"},
    ]
    metadata = [
        ("谈话时间", activity_date),
        ("谈话地点", f"{branch}办公室"),
        ("谈话人", "支部书记"),
        ("谈话对象", "党员代表"),
    ]
    paragraphs.extend(
        {
            "style": "党建正文",
            "text": f"{label}：{value}",
            "bold": True,
            "first_line_indent_pt": 0,
            "line_spacing_pt": 31,
        }
        for label, value in metadata
    )
    conversation = [
        "支部书记：近期集中学习了相关专题，结合这段时间的工作，请谈谈自己的收获和困惑，今天我们坦诚交流。",
        "党员代表：学习的最大感受是，基层工作没有捷径，必须经常走近群众、了解真实需求。但近期业务任务比较集中，有时分不清轻重缓急，工作做了不少，效果还不够理想。",
        "支部书记：任务多更要讲究方法。可以先把群众需求和工作事项分类，建立清单，按轻重缓急逐项推进；对暂时解决不了的问题，要说明原因和时限，及时反馈，不能简单搁置。",
        "党员代表：您说得对。我过去有时更关注办了多少笔业务，对后续服务效果总结不够，面对个别客户反复咨询时耐心也不足。",
        "支部书记：能够主动查找不足是好事。我们的工作成效不能只看数字，还要看群众是否满意。面对老年客户、小微经营主体等群体，更要把政策讲清楚、把流程说明白。",
        "党员代表：我会在改进服务态度和提高办事效率上同步用力，多向经验丰富的同志请教沟通方法。不过有时也会觉得成长较慢，担心自己的努力得不到体现。",
        "支部书记：基层一线就是锻炼能力的最好课堂，一次及时的服务、一次准确的风险提示，都能体现岗位价值。要沉下心来积累，遇到问题及时沟通，组织也会在实践中有针对性地培养锻炼。",
        "党员代表：听您这样分析，我心里更踏实了。今后我会坚持深入基层，多听、多记、多思考，把学习到的方法运用到实际工作中。",
        "支部书记：工作中既要敢于担当，也要守住合规底线，注意改进作风、加强学习。遇到困难及时反映，我们一起研究解决办法。",
        "党员代表：明白，我会认真改进不足，把每一项任务做细做实，以更加积极的状态服务群众、推动工作。",
    ]
    paragraphs.extend(
        {"style": "党建正文", "text": text} for text in conversation
    )
    paragraphs.append(
        {
            "style": "党建正文",
            "text": (
                "谈话效果：通过谈心谈话，党员代表进一步认识到基层工作的"
                "价值，查找了工作统筹、服务耐心和成长心态等方面的问题，"
                "明确了改进方向。支部书记及时掌握思想动态，达到了沟通思想、"
                "释疑解惑、激励担当的目的。"
            ),
        }
    )
    return paragraphs


def lecture_paragraphs(
    activity: dict,
    branch: str,
    topic: str,
) -> list[dict]:
    participants = normalize_participants(activity["participants"], branch)
    paragraphs = [
        {"style": "党建标题1", "text": "党课内容记录"},
        {
            "style": "党建正文",
            "text": (
                f"{participants}参加了本次党课。党课围绕{topic}展开，"
                "坚持理论联系实际，引导党员从政治要求、岗位职责和工作"
                "作风等方面深化认识。"
            ),
        },
        {
            "style": "党建正文",
            "text": (
                "党课强调，要持续加强理论武装，站稳人民立场，把党的创新"
                "理论转化为发现问题、分析问题和解决问题的思路方法。党员"
                "干部要主动担当作为，在服务基层、服务群众和推动发展中"
                "发挥先锋模范作用。"
            ),
        },
        {
            "style": "党建正文",
            "text": (
                "课后，党员结合岗位工作交流学习体会，表示要坚持学以致用，"
                "严守纪律规矩和业务规范，主动改进服务作风，把党课学习成果"
                "落实到每一项具体任务中。"
            ),
        },
    ]
    return paragraphs


def generic_paragraphs(activity: dict, branch: str, topic: str) -> list[dict]:
    participants = normalize_participants(activity["participants"], branch)
    return [
        {"style": "党建标题1", "text": activity["category"]},
        {
            "style": "党建正文",
            "text": (
                f"{participants}围绕{topic}开展了本次活动。活动坚持理论"
                "学习与岗位实践相结合，进一步统一思想、明确任务、凝聚力量。"
            ),
        },
        {
            "style": "党建正文",
            "text": (
                "大家表示，要立足岗位履职尽责，加强协作配合，把活动成果"
                "转化为服务群众、防范风险和推动落实的具体行动。"
            ),
        },
    ]


def classify_activity(activity: dict) -> str:
    category = activity["category"]
    if "谈心谈话" in category and "集中学习" in category:
        return "talk"
    if "谈心谈话" in category:
        return "talk"
    if "主题党日" in category:
        return "theme_day"
    if "党课" in category:
        return "lecture"
    if "集中学习" in category or "党员大会" in category or "支委会" in category:
        return "short"
    return "generic"


def merge_talk_and_third_study(activities: list[dict]) -> list[dict]:
    merged = []
    consumed = set()
    for index, activity in enumerate(activities):
        if index in consumed:
            continue
        category = activity["category"]
        if "谈心谈话" not in category or "集中学习" in category:
            merged.append(activity)
            continue

        partner_index = None
        for candidate_index in range(index + 1, len(activities)):
            if candidate_index in consumed:
                continue
            candidate = activities[candidate_index]
            if "集中学习" not in candidate["category"]:
                continue
            if ordinal_number(candidate["raw_title"]) != 3:
                continue
            partner_index = candidate_index
            break

        if partner_index is None:
            merged.append(activity)
            continue

        partner = activities[partner_index]
        consumed.add(partner_index)
        second_title = re.sub(
            r"^x{1,}(?:支行)?党支部",
            "",
            as_text(partner["raw_title"]),
            flags=re.IGNORECASE,
        )
        combined = dict(activity)
        combined["category"] = "谈心谈话和集中学习"
        combined["raw_title"] = (
            f"{as_text(activity['raw_title'])}和{second_title}"
        )
        combined["content"] = "\n".join(
            part
            for part in [
                as_text(activity["content"]),
                as_text(partner["content"]),
            ]
            if part
        )
        combined["participants"] = (
            as_text(activity["participants"])
            or as_text(partner["participants"])
        )
        combined["due"] = as_text(partner["due"]) or as_text(activity["due"])
        combined["requirements"] = (
            as_text(activity["requirements"])
            or as_text(partner["requirements"])
        )
        combined["extra_values"] = list(
            dict.fromkeys(
                [
                    *activity.get("extra_values", []),
                    *partner.get("extra_values", []),
                ]
            )
        )
        combined["merged_source_rows"] = [
            activity["source_row"],
            partner["source_row"],
        ]
        merged.append(combined)
    return merged


def filename_with_resolved_theme(filename: str, category: str, theme: str) -> str:
    if "主题党日" not in category or not PLACEHOLDER_RE.search(filename):
        return filename
    safe_theme = re.sub(r'[\\/:*?"<>|]', "", theme).strip()
    return PLACEHOLDER_RE.sub(safe_theme, filename)


def document_paragraphs(
    activity: dict,
    branch: str,
    month: int,
    year: int,
) -> tuple[str, list[dict]]:
    topic = learning_topic(activity)
    theme = extract_theme(activity["raw_title"], activity["content"], activity["category"])
    kind = classify_activity(activity)
    if kind == "theme_day":
        return kind, theme_day_paragraphs(
            activity, branch, month, year, topic, theme
        )
    if kind == "talk":
        return kind, talk_paragraphs(activity, branch, month, year, topic)
    if kind == "lecture":
        return kind, lecture_paragraphs(activity, branch, topic)
    if kind == "short":
        return kind, short_record_paragraphs(activity, branch, topic)
    return kind, generic_paragraphs(activity, branch, topic)


def ensure_standard_template(template_path: Path, overwrite: bool) -> Path:
    if template_path.exists() and not overwrite:
        return template_path
    return build_template(template_path)


def load_checklist(checklist_path: Path):
    openpyxl = import_openpyxl()
    try:
        workbook = openpyxl.load_workbook(checklist_path, data_only=True)
    except Exception as exc:
        raise SystemExit(f"Cannot read checklist workbook: {exc}") from exc
    if not workbook.worksheets:
        raise SystemExit("Checklist workbook contains no worksheets.")
    return workbook


def build_delivery(
    checklist_path: Path,
    output_dir: Path,
    branch: str,
    template_path: Path,
    overwrite: bool,
    dry_run: bool,
) -> dict:
    workbook = load_checklist(checklist_path)
    values = workbook_text_values(workbook)
    month = infer_month(checklist_path, values)
    year = infer_year(checklist_path, values)
    activities = merge_talk_and_third_study(
        extract_activities(workbook.worksheets[0])
    )
    if not activities:
        raise SystemExit("No activities were found in the checklist.")

    template_path = ensure_standard_template(template_path, overwrite=False)
    Document, WD_ALIGN_PARAGRAPH, Pt = import_docx_dependencies()
    document_specs = []
    output_names = set()
    for activity in activities:
        filename, _warnings = normalize_title(
            activity["raw_title"], activity["category"], branch, month
        )
        theme = extract_theme(
            activity["raw_title"], activity["content"], activity["category"]
        )
        filename = filename_with_resolved_theme(
            filename, activity["category"], theme
        )
        if filename in output_names:
            filename = f"{Path(filename).stem}-{activity['sequence']}.docx"
        output_names.add(filename)
        kind, paragraphs = document_paragraphs(
            activity, branch, month, year
        )
        activity["output_kind"] = kind
        activity["filename"] = filename
        activity["paragraph_count"] = len(paragraphs)
        document_specs.append(
            validate_document_spec(
                {
                    "template": str(template_path),
                    "filename": filename,
                    "title": Path(filename).stem,
                    "paragraphs": paragraphs,
                },
                template_path.parent,
                0,
            )
        )

    outputs = [
        build_document(
            Document,
            WD_ALIGN_PARAGRAPH,
            Pt,
            document_spec,
            output_dir,
            overwrite,
            dry_run,
        )
        for document_spec in document_specs
    ]
    return {
        "checklist": str(checklist_path),
        "month": month,
        "year": year,
        "branch": branch,
        "template": str(template_path),
        "output_dir": str(output_dir),
        "activity_count": len(activities),
        "document_count": len(outputs),
        "documents": [path.name for path in outputs],
        "outputs": outputs,
    }


def print_result(result: dict, dry_run: bool) -> None:
    action = "验证" if dry_run else "生成"
    print(f"{action}月份: {result['year']}年{result['month']}月")
    print(f"党支部: {result['branch']}")
    print(f"清单: {result['checklist']}")
    print(f"统一模板: {result['template']}")
    print(f"活动项: {result['activity_count']}")
    print(f"交付文件: {result['document_count']}")
    for name in result["documents"]:
        print(f"  {name}")
    if not dry_run:
        print(f"输出目录: {result['output_dir']}")


def main() -> None:
    configure_stdout()
    parser = argparse.ArgumentParser(
        description=(
            "Generate all monthly party-building DOCX files from one checklist. "
            "Historical monthly documents are not required."
        )
    )
    parser.add_argument(
        "--checklist",
        required=True,
        type=Path,
        help="Path to the monthly重点工作清单 XLSX file.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Output folder. Defaults to a folder beside the checklist.",
    )
    parser.add_argument(
        "--branch-name",
        help=(
            "Party branch name for this run. Can be saved once for future "
            "checklist-only runs."
        ),
    )
    parser.add_argument(
        "--config",
        type=Path,
        help=f"Local branch config path. Default: {DEFAULT_CONFIG_NAME}.",
    )
    parser.add_argument(
        "--save-config",
        action="store_true",
        help="Save --branch-name into the local ignored config file.",
    )
    parser.add_argument(
        "--template",
        type=Path,
        help="Override the bundled standard DOCX template.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing DOCX files in the output folder.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate parsing and generation without writing DOCX files.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print the result summary as JSON.",
    )
    args = parser.parse_args()

    checklist = args.checklist.resolve()
    if not checklist.is_file() or checklist.suffix.lower() != ".xlsx":
        raise SystemExit(f"Checklist must be an XLSX file: {checklist}")

    workbook = load_checklist(checklist)
    values = workbook_text_values(workbook)
    config = load_local_config(args.config)
    branch, branch_source = resolve_branch(
        args.branch_name, config, values
    )
    if args.save_config:
        if not args.branch_name:
            raise SystemExit("--save-config requires --branch-name.")
        config_path = save_local_config(args.config, branch)
        print(f"Saved local branch config: {config_path}")

    month = infer_month(checklist, values)
    year = infer_year(checklist, values)
    output_dir = (
        args.output_dir.resolve()
        if args.output_dir
        else checklist.parent / f"{year}年{month}月党建生成"
    )
    template_path = (
        args.template.resolve()
        if args.template
        else default_template_path()
    )

    result = build_delivery(
        checklist,
        output_dir,
        branch,
        template_path,
        args.overwrite,
        args.dry_run,
    )
    result["branch_source"] = branch_source
    if args.json:
        printable = dict(result)
        printable.pop("outputs", None)
        print(json.dumps(printable, ensure_ascii=False, indent=2))
    else:
        print_result(result, args.dry_run)

    if args.dry_run:
        return

    report = validate_delivery(
        output_dir,
        month,
        branch,
        len(result["outputs"]),
        None,
        result["documents"],
    )
    if report["errors"]:
        for message in report["errors"]:
            print(f"错误: {message}", file=sys.stderr)
        raise SystemExit(1)
    print("结构验收通过。请继续完成逐页渲染检查。")


if __name__ == "__main__":
    main()
