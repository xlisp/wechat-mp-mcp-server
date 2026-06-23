#!/usr/bin/env python3
"""微信公众号流量分析 MCP Server (wechat_mp_mcp)

解析微信公众号后台导出的「数据趋势」xls/xlsx 文件，并把一套实战经验沉淀成工具：
渠道结构拆解、文章排行、互动率、以及最关心的「某篇文章什么时候 / 能不能过万阅读」预测。

风格对齐本机的 filesystem MCP：FastMCP + stdio，单文件、零外部网络依赖。

依赖：mcp、pandas、xlrd(读 .xls)、openpyxl(读 .xlsx)
运行：python wechat_mcp.py            # stdio，供 Claude Desktop 调用
Claude Desktop 配置示例：
    {
      "mcpServers": {
        "wechat-mp": { "command": "python", "args": ["/绝对路径/wechat_mcp.py"] }
      }
    }
"""

from __future__ import annotations

import json
import math
import os
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from pydantic import BaseModel, ConfigDict, Field

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("wechat_mp_mcp")

# ──────────────────────────────────────────────────────────────────────────
# 经验常量：判断飞轮 / 衰减、爆款类型的阈值，都集中在这里，便于日后校准。
# ──────────────────────────────────────────────────────────────────────────
FLYWHEEL_DAILY_GAIN = 350      # 日增 ≥ 此值视为「飞轮仍在转」，破圈未停
PLATEAU_DAILY_GAIN = 150       # 日增 < 此值视为「基本定型」，进入平台期
HEALTHY_SHARE_RATE = 0.10      # 分享率 ≥ 10% 算健康（每 10 个读者 1 次转发）
HOT_SHARE_RATE = 0.20          # 分享率 ≥ 20% 属罕见高传播，强烈助推算法
ALGO_DOMINANT = 0.60           # 单一「推荐」渠道占比 ≥ 60% → 算法型爆款
SOCIAL_CHANNELS = {"聊天会话", "朋友圈"}  # 社交传播渠道
ALGO_CHANNELS = {"推荐", "搜一搜"}        # 算法/搜索分发渠道


class ResponseFormat(str, Enum):
    MARKDOWN = "markdown"
    JSON = "json"


# ──────────────────────────────────────────────────────────────────────────
# 解析层：把导出文件里塞在一个 sheet 中的三块表拆开。
#   Block A 列[1-3]  : 日期 / 渠道 / 阅读人数      —— 每日分渠道
#   Block B 列[5-9]  : 日期 / 分享 / 跳转原文 / 收藏 / 发表篇数 —— 每日账号级
#   Block C 列[11-15]: 渠道 / 发表日期 / 标题 / 阅读人数 / 阅读占比 —— 每篇分渠道
# ──────────────────────────────────────────────────────────────────────────
_CACHE: Dict[str, Dict[str, Any]] = {}


def _read_raw(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    ext = os.path.splitext(path)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    return pd.read_excel(path, header=None, engine=engine)


def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def parse_report(path: str) -> Dict[str, Any]:
    """解析一份导出文件，结果带 (路径+修改时间) 缓存。返回三块 DataFrame 与日期范围。"""
    key = f"{os.path.abspath(path)}::{os.path.getmtime(path)}"
    if key in _CACHE:
        return _CACHE[key]

    raw = _read_raw(path)

    a = raw.iloc[2:, [1, 2, 3]].copy()
    a.columns = ["date", "channel", "readers"]
    a = a[a["date"].astype(str).str.match(r"\d{4}-\d{2}-\d{2}")].copy()
    a["readers"] = _to_num(a["readers"]).fillna(0).astype(int)

    b = raw.iloc[2:, [5, 6, 7, 8, 9]].copy()
    b.columns = ["date", "shares", "read_orig", "saves", "published"]
    b = b[b["date"].astype(str).str.match(r"\d{4}-\d{2}-\d{2}")].copy()
    for c in ["shares", "read_orig", "saves", "published"]:
        b[c] = _to_num(b[c]).fillna(0).astype(int)

    c = raw.iloc[2:, [11, 12, 13, 14, 15]].copy()
    c.columns = ["channel", "pubdate", "title", "readers", "ratio"]
    c = c.dropna(subset=["title"]).copy()
    c["readers"] = _to_num(c["readers"]).fillna(0).astype(int)
    c["pubdate"] = c["pubdate"].astype(str).str.replace(r"\.0$", "", regex=True)

    dr = (str(a["date"].min()), str(a["date"].max())) if len(a) else ("?", "?")
    result = {"daily_channel": a, "daily_account": b, "article_channel": c, "date_range": dr}
    _CACHE[key] = result
    return result


# ──────────────────────────────────────────────────────────────────────────
# 分析层：可独立测试的纯函数，工具只做参数校验 + 调用 + 格式化。
# ──────────────────────────────────────────────────────────────────────────
def _overview(rep: Dict[str, Any]) -> Dict[str, Any]:
    a, b, c = rep["daily_channel"], rep["daily_account"], rep["article_channel"]
    total_reads = int(a[a["channel"] == "全部"]["readers"].sum())
    shares = int(b["shares"].sum())
    saves = int(b["saves"].sum())
    read_orig = int(b["read_orig"].sum())
    n_articles = c[c["channel"] == "全部"]["title"].nunique()
    safe = total_reads or 1
    return {
        "date_range": rep["date_range"],
        "total_reads": total_reads,
        "total_shares": shares,
        "total_saves": saves,
        "total_read_original": read_orig,
        "articles_published": int(n_articles),
        "share_rate": round(shares / safe, 4),
        "save_rate": round(saves / safe, 4),
    }


def _channel_mix(rep: Dict[str, Any], title: Optional[str]) -> Tuple[int, List[Tuple[str, int, float]]]:
    c = rep["article_channel"]
    if title:
        total = int(c[(c["title"] == title) & (c["channel"] == "全部")]["readers"].sum())
        sub = c[(c["title"] == title) & (c["channel"] != "全部")]
    else:
        a = rep["daily_channel"]
        total = int(a[a["channel"] == "全部"]["readers"].sum())
        sub = a[a["channel"] != "全部"]
    g = sub.groupby("channel")["readers"].sum().sort_values(ascending=False)
    safe = total or 1
    mix = [(ch, int(r), round(r / safe, 4)) for ch, r in g.items() if r > 0]
    return total, mix


def _classify_article(mix: List[Tuple[str, int, float]]) -> Tuple[str, str]:
    """据渠道构成判定爆款类型，返回 (类型, 一句话解读)。"""
    share = {ch: pct for ch, _, pct in mix}
    algo = sum(share.get(ch, 0) for ch in ALGO_CHANNELS)
    social = sum(share.get(ch, 0) for ch in SOCIAL_CHANNELS)
    if algo >= ALGO_DOMINANT:
        return "算法型", "主要靠系统「推荐」分发给陌生人，天花板取决于选题是否踩中算法，无法靠人脉复制"
    if social >= ALGO_DOMINANT:
        return "社交型", "主要靠真人转发（聊天会话/朋友圈），靠的是共鸣与社交货币，标题情绪钩子与观点是关键"
    if algo >= 0.30 and social >= 0.30:
        return "双引擎", "算法推荐与社交转发兼得，这是最理想的破圈形态，天花板最高"
    return "混合型", "流量来源较分散，尚未形成明确的分发主引擎"


def _article_ranking(rep: Dict[str, Any]) -> List[Dict[str, Any]]:
    c = rep["article_channel"]
    g = (
        c[c["channel"] == "全部"]
        .groupby(["pubdate", "title"])["readers"].sum()
        .sort_values(ascending=False)
    )
    return [{"pubdate": d, "title": t, "reads": int(r)} for (d, t), r in g.items()]


def forecast_milestone(
    observations: List[Tuple[str, int]],
    milestone: int = 10000,
    share_rate: Optional[float] = None,
) -> Dict[str, Any]:
    """核心预测：给定一篇文章在若干天的累计阅读快照，推算能否 / 何时达到 milestone。

    模型并行跑两套，因为公众号文章命运取决于「自然衰减」还是「飞轮继续转」：
      · 衰减模型：日增按几何级数 r(<1) 衰减，剩余增量 = last_gain * r/(1-r)，得到自然天花板。
      · 飞轮模型：r≥1（持平或加速）时，按当前日增线性外推达成天数。
    再结合日增分水岭与分享率给出过万概率档位与「盯哪个指标」的建议。
    """
    obs = sorted(observations, key=lambda x: x[0])
    if len(obs) < 2:
        return {"error": "至少需要两个不同日期的累计阅读快照才能估算增速"}

    reads = [int(r) for _, r in obs]
    current = reads[-1]
    gains = [reads[i] - reads[i - 1] for i in range(1, len(reads))]
    last_gain = gains[-1]

    if current >= milestone:
        return {"verdict": "已达成", "current_reads": current, "milestone": milestone,
                "note": f"已超过 {milestone}，无需预测"}

    gap = milestone - current

    # 估衰减比 r：用最近两个日增，缺则用全程几何均值
    if len(gains) >= 2 and gains[-2] > 0:
        r = last_gain / gains[-2]
    elif len(gains) >= 2 and gains[0] > 0:
        r = (gains[-1] / gains[0]) ** (1 / (len(gains) - 1)) if gains[-1] > 0 else 0.0
    else:
        r = 0.85  # 信息不足时的保守经验值

    scenarios: Dict[str, Any] = {}

    # —— 衰减模型 ——
    if 0 < r < 1 and last_gain > 0:
        remaining = last_gain * r / (1 - r)
        ceiling = round(current + remaining)
        # 等比衰减下逼近 milestone 需要的天数（若天花板够高）
        days = None
        if ceiling >= milestone:
            # current + last_gain*(r + r^2 + ... + r^n) >= milestone
            need = gap / last_gain
            ratio_term = 1 - need * (1 - r)
            if ratio_term > 0:
                days = math.ceil(math.log(ratio_term) / math.log(r))
        scenarios["decay"] = {"r": round(r, 3), "natural_ceiling": ceiling,
                              "reaches_milestone": ceiling >= milestone,
                              "est_days_if_reaches": days}
    else:
        scenarios["decay"] = {"r": round(r, 3), "natural_ceiling": current,
                              "reaches_milestone": False, "est_days_if_reaches": None}

    # —— 飞轮模型（持平/加速）——
    flywheel_days = math.ceil(gap / last_gain) if last_gain > 0 else None
    scenarios["flywheel"] = {"assume_daily_gain": last_gain, "est_days": flywheel_days,
                             "premise": "日增维持当前水平不衰减"}

    # —— 概率档位 ——
    # 动量（单日日增）只是必要不充分条件，权重不宜过大；
    # 真正能把概率推过 60% 的是「趋势 r≥1」，但这需要 ≥3 个观测点才可信。
    score = 0.0
    if last_gain >= FLYWHEEL_DAILY_GAIN:
        score += 0.25
    elif last_gain >= PLATEAU_DAILY_GAIN:
        score += 0.12
    enough_data = len(obs) >= 3  # 两点无法区分衰减 vs 飞轮
    if enough_data and r >= 1.0:
        score += 0.30
    elif enough_data and r >= 0.9:
        score += 0.12
    if share_rate is not None:
        if share_rate >= HOT_SHARE_RATE:
            score += 0.15
        elif share_rate >= HEALTHY_SHARE_RATE:
            score += 0.07
    if scenarios["decay"]["reaches_milestone"]:
        score += 0.25
    prob = max(0.05, min(0.95, score))
    # 数据不足且自然天花板达不到目标时，诚实封顶：最多算「有机会但非基准情形」
    if not enough_data and not scenarios["decay"]["reaches_milestone"]:
        prob = min(prob, 0.45)

    if prob >= 0.6:
        band = "较有希望"
    elif prob >= 0.35:
        band = "有机会但非基准情形"
    else:
        band = "可能性较低"

    watch = (
        f"接下来 2-3 天盯「日增」即可：≥{FLYWHEEL_DAILY_GAIN} 说明飞轮在转、过万有戏；"
        f"掉到 <{PLATEAU_DAILY_GAIN} 基本定型于自然天花板附近；"
        f"若某天突然跳一个台阶（如单日 +600 以上）则是推荐池放量，概率大幅上升。"
    )

    return {
        "current_reads": current,
        "milestone": milestone,
        "gap": gap,
        "latest_daily_gain": last_gain,
        "decay_ratio_r": round(r, 3),
        "scenarios": scenarios,
        "probability": round(prob, 2),
        "probability_band": band,
        "watch_metric": watch,
    }


# ──────────────────────────────────────────────────────────────────────────
# 格式化
# ──────────────────────────────────────────────────────────────────────────
def _pct(x: float) -> str:
    return f"{x * 100:.1f}%"


def _fmt_overview(o: Dict[str, Any]) -> str:
    d0, d1 = o["date_range"]
    return (
        f"## 公众号流量概况（{d0} ~ {d1}）\n"
        f"- 总阅读人数：**{o['total_reads']:,}**\n"
        f"- 发文篇数：{o['articles_published']}\n"
        f"- 分享 {o['total_shares']:,}（分享率 **{_pct(o['share_rate'])}**）\n"
        f"- 收藏 {o['total_saves']:,}（收藏率 {_pct(o['save_rate'])}）\n"
        f"- 跳转阅读原文：{o['total_read_original']:,}"
        + ("（为 0，文末缺少外链引导，是流失点）" if o["total_read_original"] == 0 else "")
    )


def _fmt_channels(total: int, mix: List[Tuple[str, int, float]], title: Optional[str]) -> str:
    head = f"## 渠道构成 · {title}（总阅读 {total:,}）\n" if title else f"## 全账号渠道构成（总阅读 {total:,}）\n"
    lines = [f"- {ch}：{r:,}（{_pct(pct)}）" for ch, r, pct in mix]
    kind, note = _classify_article(mix)
    return head + "\n".join(lines) + f"\n\n**类型判定：{kind}** — {note}"


# ──────────────────────────────────────────────────────────────────────────
# Pydantic 输入模型
# ──────────────────────────────────────────────────────────────────────────
class _Base(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")


class ReportInput(_Base):
    path: str = Field(..., description="导出文件路径，如 /Users/xlisp/Downloads/tendency_xxx.xls", min_length=1)
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN, description="markdown 人读 / json 机读")


class ChannelInput(ReportInput):
    title: Optional[str] = Field(default=None, description="文章标题；留空则看全账号渠道构成")


class Observation(_Base):
    date: str = Field(..., description="观测日期 YYYY-MM-DD")
    reads: int = Field(..., description="该日的累计阅读人数", ge=0)


class ForecastInput(_Base):
    observations: List[Observation] = Field(..., description="同一篇文章在不同日期的累计阅读快照，至少两条", min_length=2)
    milestone: int = Field(default=10000, description="目标阅读量阈值，默认过万", ge=1)
    share_rate: Optional[float] = Field(default=None, description="该文分享率(分享数/阅读数)，提供可提升预测准确度", ge=0, le=1)
    response_format: ResponseFormat = Field(default=ResponseFormat.MARKDOWN, description="markdown / json")


# ──────────────────────────────────────────────────────────────────────────
# 工具
# ──────────────────────────────────────────────────────────────────────────
@mcp.tool(name="wechat_overview", annotations={"title": "公众号流量概况", "readOnlyHint": True,
          "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def wechat_overview(params: ReportInput) -> str:
    """读取导出文件，给出整体概况：总阅读、发文数、分享/收藏率、跳转原文等核心指标。

    Args: params.path 导出文件路径；params.response_format 输出格式。
    Returns: markdown 文本或 JSON 字符串。
    """
    try:
        o = _overview(parse_report(params.path))
    except Exception as e:  # noqa: BLE001
        return f"Error: 解析失败 — {type(e).__name__}: {e}。请确认是公众号后台导出的「数据趋势」xls/xlsx。"
    return json.dumps(o, ensure_ascii=False, indent=2) if params.response_format == ResponseFormat.JSON else _fmt_overview(o)


@mcp.tool(name="wechat_channel_breakdown", annotations={"title": "渠道构成与爆款类型", "readOnlyHint": True,
          "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def wechat_channel_breakdown(params: ChannelInput) -> str:
    """拆解流量来源渠道，并自动判定爆款类型（算法型 / 社交型 / 双引擎 / 混合型）。

    Args: params.path 文件路径；params.title 文章标题(可选)；params.response_format。
    Returns: 各渠道阅读与占比 + 类型判定。
    """
    try:
        rep = parse_report(params.path)
        total, mix = _channel_mix(rep, params.title)
    except Exception as e:  # noqa: BLE001
        return f"Error: {type(e).__name__}: {e}"
    if not mix:
        return f"Error: 未找到{'文章「' + params.title + '」的' if params.title else ''}渠道数据，请核对标题是否与导出一致。"
    if params.response_format == ResponseFormat.JSON:
        kind, note = _classify_article(mix)
        return json.dumps({"title": params.title, "total_reads": total,
                           "channels": [{"channel": ch, "reads": r, "share": pct} for ch, r, pct in mix],
                           "type": kind, "note": note}, ensure_ascii=False, indent=2)
    return _fmt_channels(total, mix, params.title)


@mcp.tool(name="wechat_article_ranking", annotations={"title": "文章阅读排行", "readOnlyHint": True,
          "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def wechat_article_ranking(params: ReportInput) -> str:
    """按总阅读量给所有文章排行，识别头部效应（单篇贡献占比）。

    Args: params.path 文件路径；params.response_format。
    Returns: 文章排行列表（发表日期、标题、阅读）。
    """
    try:
        rep = parse_report(params.path)
        rank = _article_ranking(rep)
    except Exception as e:  # noqa: BLE001
        return f"Error: {type(e).__name__}: {e}"
    if params.response_format == ResponseFormat.JSON:
        return json.dumps(rank, ensure_ascii=False, indent=2)
    tot = sum(x["reads"] for x in rank) or 1
    lines = [f"## 文章阅读排行（共 {len(rank)} 篇，合计 {tot:,}）"]
    for i, x in enumerate(rank, 1):
        lines.append(f"{i}. **{x['reads']:,}**（{_pct(x['reads']/tot)}）· {x['pubdate']} · {x['title']}")
    top3 = sum(x["reads"] for x in rank[:3])
    lines.append(f"\n> 头部效应：前 3 篇占全月 **{_pct(top3/tot)}**。")
    return "\n".join(lines)


@mcp.tool(name="wechat_forecast_milestone", annotations={"title": "文章过万预测", "readOnlyHint": True,
          "destructiveHint": False, "idempotentHint": True, "openWorldHint": False})
async def wechat_forecast_milestone(params: ForecastInput) -> str:
    """预测某篇文章能否 / 何时达到目标阅读量（默认过万）。

    输入同一篇文章在不同日期的累计阅读快照（至少两条），并行用衰减模型与飞轮模型推算，
    结合日增分水岭与分享率给出概率档位和「盯哪个指标」的可执行建议。

    Args:
        params.observations: [{date, reads}, ...] 累计阅读快照，至少两条不同日期。
        params.milestone: 目标阈值，默认 10000。
        params.share_rate: 分享率(可选)，提供后预测更准。
    Returns: 含 current_reads / scenarios(decay+flywheel) / probability / watch_metric 的结果。
    """
    obs = [(o.date, o.reads) for o in params.observations]
    res = forecast_milestone(obs, milestone=params.milestone, share_rate=params.share_rate)
    if "error" in res:
        return f"Error: {res['error']}"
    if params.response_format == ResponseFormat.JSON:
        return json.dumps(res, ensure_ascii=False, indent=2)
    if res.get("verdict") == "已达成":
        return f"✅ 当前 {res['current_reads']:,} 已超过 {res['milestone']:,}，无需预测。"
    d, f = res["scenarios"]["decay"], res["scenarios"]["flywheel"]
    out = [
        f"## 过万预测：当前 {res['current_reads']:,} → 目标 {res['milestone']:,}（差 {res['gap']:,}）",
        f"- 最新日增：**{res['latest_daily_gain']:,}** ｜ 衰减比 r≈{res['decay_ratio_r']}",
        "",
        f"**衰减剧本**：自然天花板约 **{d['natural_ceiling']:,}**，"
        + (f"可达标，约需 {d['est_days_if_reaches']} 天。" if d["reaches_milestone"] else "达不到目标。"),
        f"**飞轮剧本**：若日增维持 {f['assume_daily_gain']:,} 不衰减，约 "
        + (f"{f['est_days']} 天达标。" if f["est_days"] else "无法估算。"),
        "",
        f"**过万概率：{_pct(res['probability'])}（{res['probability_band']}）**",
        "",
        res["watch_metric"],
    ]
    return "\n".join(out)


if __name__ == "__main__":
    mcp.run()

