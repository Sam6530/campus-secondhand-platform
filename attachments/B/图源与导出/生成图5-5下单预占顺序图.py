# -*- coding: utf-8 -*-
"""生成图5-5 下单与预占顺序图（PNG + SVG）。

运行：python 生成图5-5下单预占顺序图.py
依据：正文5.5、UC05、BR06、BR10、BR14、BR20；ST-B-08、ST-B-13。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from 绘图工具 import Diagram, NOTE_FILL, wrap  # noqa: E402

BUYER, UI, ACCOUNT, ITEM, ORDER = 170, 470, 750, 1020, 1300


def build():
    g = Diagram(1500, 1500)

    # 生命线
    for x, label in [(BUYER, "买方"), (UI, "下单界面"), (ACCOUNT, "账户"),
                     (ITEM, "物品"), (ORDER, "订单")]:
        g.lifeline(x, 60, 1270, label, width=200, size=26)

    # 激活期
    g.activation(BUYER, 250, 1235)
    g.activation(UI, 250, 1235)
    g.activation(ACCOUNT, 330, 725)
    g.activation(ITEM, 780, 1095)
    g.activation(ORDER, 480, 1165)

    # 消息
    g.message(BUYER, UI, 250, "提交下单意图（物品、价格）")
    g.message(UI, ACCOUNT, 330, "核验身份与归属（BR14）")
    g.message(ACCOUNT, UI, 400, "有效／无权（拒绝）", dash=True)
    g.message(UI, ORDER, 480, "识别同次操作（BR10）")
    g.message(ORDER, UI, 550, "原回执／处理中／无记录", dash=True)
    g.message(UI, ACCOUNT, 630, "检查双方新交易资格（BR06、BR12）")
    g.message(ACCOUNT, UI, 700, "允许／拒绝", dash=True)
    g.message(UI, ITEM, 780, "校验公开可交易（BR03）")
    g.message(ITEM, UI, 850, "可交易／竞争失败", dash=True)
    g.message(UI, ORDER, 930, "建立订单并保存快照（BR20）")
    g.message(ORDER, ITEM, 1000, "预占物品（BR06）")
    g.message(ITEM, ORDER, 1070, "占用成立", dash=True)
    g.message(ORDER, UI, 1140, "订单标识（待卖方确认）", dash=True)
    g.message(UI, BUYER, 1210, "回执／进度／拒绝原因", dash=True)

    # 失败与约束说明
    g.box((90, 1290, 1420, 1470), "", fill=NOTE_FILL, pad=16, color="#B45309")
    g.text((755, 1322), "失败、重试与并发约束", 24)
    note = ("失败分支：M3／M7／M9 返回拒绝时，不建立订单、不产生占用，物品状态保持（ST-B-S1）。\n"
            "同次重试：已有回执或处理中时返回原结果，不新增订单（ST-B-S2、BR10）。\n"
            "同时成立：M10—M12 中订单建立与物品占用为同一业务结果，中途失败不留下半个占用；并发下单至多一笔成立（BR06）。")
    y = 1360
    for para in note.split("\n"):
        for line in wrap(para, 21, 1270):
            g.text((112, y), line, 21, anchor="lm")
            y += 30
        y += 4

    return g.save("图5-5_下单预占顺序图")


if __name__ == "__main__":
    png, svg = build()
    print("已生成：", png.name, "、", svg.name)
