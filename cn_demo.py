#!/usr/bin/env python3
"""
Generate a Chinese Education Demo PPT — showcasing all enhanced features.
Matches the quality level of reference PPTs (小瓶盖儿 / 垃圾分类).
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from engines.pptx_builder import build_ppt, THEMES
import random

seed = random.randint(1, 999999)

config = {
    "title": "我的动物朋友",
    "theme": "cn_education",
    "seed": seed,
    "slides": [
        # 1. Title
        {
            "type": "cn_title",
            "heading": "我的动物朋友",
            "subheading": "小学语文习作指导课",
            "bg_style": "bokeh",
            "palette": "forest",
            "seed": seed,
        },

        # 2. Task separator
        {
            "type": "cn_task_separator",
            "text": "任务一",
            "palette": "forest",
            "seed": seed + 1,
        },

        # 3. Scenario — lost sheep
        {
            "type": "cn_scenario",
            "scenario_label": "情境一",
            "text": "星期天放羊回来，发现我最喜爱的一只小羊不见了，\n我想请小伙伴帮忙找一找。",
            "thought": "我要跟小伙伴强调，小羊的左眼圈是黑色的……",
            "palette": "forest",
            "seed": seed + 2,
        },

        # 4. Scenario — neighbor's dog
        {
            "type": "cn_scenario",
            "scenario_label": "情境二",
            "text": "我们全家要外出旅行一段时间，\n只好请邻居帮忙喂养我的小狗。",
            "thought": "我要给邻居讲清楚，我家小狗特别爱吃肉……",
            "label_color": (220, 80, 40),
            "palette": "warm",
            "seed": seed + 3,
        },

        # 5. Task separator
        {
            "type": "cn_task_separator",
            "text": "任务二",
            "palette": "ocean",
            "seed": seed + 4,
        },

        # 6. Content with side label
        {
            "type": "cn_content",
            "side_label": "习作内容",
            "heading": "选择情境介绍动物朋友",
            "body": [
                "• 有时候，我们需要向别人介绍自己的动物朋友",
                "• 从下面的情境中选择一个，向别人介绍你的动物朋友",
                "• 如果你没有养过这些动物，也可以就自己熟悉的动物创设一个情境来写",
                "• 写之前想一想，你打算从哪些方面介绍它",
            ],
            "palette": "ocean",
            "seed": seed + 5,
        },

        # 7. Mind map
        {
            "type": "cn_mind_map",
            "heading": "介绍动物朋友的方面",
            "root_text": "动物特点",
            "children": [
                {"text": "外形", "sub": ["毛色", "体态", "眼睛"]},
                {"text": "习性", "sub": ["饮食", "睡眠", "活动"]},
                {"text": "趣事", "sub": ["互动", "趣闻"]},
                {"text": "性格", "sub": ["温顺", "调皮", "忠诚"]},
            ],
            "palette": "forest",
            "seed": seed + 6,
        },

        # 8. Task separator
        {
            "type": "cn_task_separator",
            "text": "任务三",
            "palette": "coral",
            "seed": seed + 7,
        },

        # 9. Flowchart — writing process
        {
            "type": "cn_flowchart",
            "heading": "写作步骤",
            "steps": [
                {"title": "选择情境", "description": "想一想需要向谁介绍动物朋友"},
                {"title": "确定特点", "description": "抓住动物最突出的特点来写"},
                {"title": "总写分述", "description": "先总写后分述，条理清晰"},
                {"title": "生动描写", "description": "用具体事例展现动物特点"},
            ],
            "palette": "ocean",
            "seed": seed + 8,
        },

        # 10. Reading passage
        {
            "type": "cn_reading",
            "title": "范例欣赏",
            "passage": "    欢欢特别爱吃肉，我家一做肉，欢欢闻到味儿就蹲在厨房等着了。\n到我吃肉时，欢欢就会走过来，两条前腿跪下来，两只水灵灵的小眼睛一直望着我，像是在讨好地对我说：'给我吃一点儿吧，你可别吃光了'。\n    欢欢很聪明，爱让人逗它玩儿。我们两个常玩的游戏是梳毛和装'死'。用手一摸它的背，它就会乖乖地趴在地上……",
            "sidebar": [
                "先总写：欢欢的特点",
                "分述一：爱吃肉（趣事）",
                "分述二：聪明（性格）",
                "用动作描写展现特点",
                "拟人化手法增加趣味",
            ],
            "palette": "sunset",
            "seed": seed + 9,
        },

        # 11. Evaluation table
        {
            "type": "cn_evaluation",
            "heading": "《我的动物朋友》评价表",
            "headers": ["评价内容", "评价标准", "自评"],
            "rows": [
                ["情境选择", "情境合理，目的明确", "⭐⭐⭐"],
                ["特点突出", "抓住动物最突出的特点", "⭐⭐⭐"],
                ["描写生动", "用具体事例展现特点", "⭐⭐⭐"],
                ["结构清晰", "先总后分，条理清楚", "⭐⭐⭐"],
                ["语言表达", "语句通顺，用词准确", "⭐⭐⭐"],
            ],
            "header_color": (0, 120, 215),
            "palette": "ocean",
            "seed": seed + 10,
        },

        # 12. Brackets analysis
        {
            "type": "cn_brackets",
            "heading": "写作结构分析",
            "groups": [
                {
                    "label": "总写",
                    "items": ["开篇点明动物朋友", "引出主要特点"],
                    "position": {"left": 0.3, "top": 1.8, "width": 3.5, "height": 1.5},
                },
                {
                    "label": "分述",
                    "items": ["外形特点", "生活习性", "有趣故事"],
                    "position": {"left": 0.3, "top": 3.8, "width": 3.5, "height": 2},
                },
                {
                    "label": "总结",
                    "items": ["表达喜爱之情", "首尾呼应"],
                    "position": {"left": 0.3, "top": 5.8, "width": 3.5, "height": 1.2},
                },
            ],
            "palette": "forest",
            "seed": seed + 11,
        },

        # 13. End
        {
            "type": "cn_title",
            "heading": "谢谢大家",
            "subheading": "期待你的精彩作品！",
            "bg_style": "bokeh",
            "palette": "sakura",
            "seed": seed + 12,
        },
    ],
}

output = "cn_education_demo.pptx"
build_ppt(config, output)
print(f"✅ Generated: {output}")
print(f"   Theme: cn_education")
print(f"   Slides: {len(config['slides'])}")
print(f"   Seed: {seed}")
