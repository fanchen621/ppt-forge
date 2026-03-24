#!/usr/bin/env python3
"""Generate a professional 垃圾分类 (Garbage Classification) PPT."""
import sys, os, random
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from engines.pptx_builder import build_ppt

seed = random.randint(1, 999999)

config = {
    "title": "垃圾分类 从我做起",
    "theme": "cn_bright",
    "seed": seed,
    "slides": [
        # 1. Title
        {
            "type": "cn_title",
            "heading": "垃圾分类",
            "subheading": "从我做起 · 垃圾分类一小步 健康文明一大步",
            "bg_style": "gradient_mesh",
            "palette": "emerald",
            "seed": seed,
        },
        # 2. Agenda
        {
            "type": "agenda",
            "heading": "目录",
            "items": [
                "什么是生活垃圾分类",
                "为什么要实施生活垃圾分类",
                "垃圾分类的投放指导",
                "生活垃圾分类的体系建设",
                "我们能做什么",
            ],
            "palette": "emerald",
            "seed": seed + 1,
        },
        # 3. Section header - What
        {
            "type": "cn_task_separator",
            "text": "01\n什么是生活垃圾分类",
            "palette": "emerald",
            "seed": seed + 2,
        },
        # 4. Definition
        {
            "type": "cn_content",
            "side_label": "概念",
            "heading": "什么是生活垃圾分类？",
            "body": [
                "按规定将垃圾分类储存、分类投放和搬运，",
                "从而转变成公共资源的一系列活动的总称。",
                "",
                "垃圾分类的目的是提高垃圾的资源价值和经济价值，",
                "力争物尽其用，减少垃圾处理量和处理设备的使用，",
                "降低处理成本，减少土地资源的消耗。",
            ],
            "palette": "emerald",
            "seed": seed + 3,
        },
        # 5. Classification flow
        {
            "type": "cn_flowchart",
            "heading": "分类流程",
            "steps": [
                {"title": "分类投放", "description": "居民将垃圾按类别投放到对应垃圾桶"},
                {"title": "分类收集", "description": "环卫工人分类收集各类垃圾"},
                {"title": "分类运输", "description": "不同类别垃圾用专车分别运输"},
                {"title": "分类处理", "description": "根据垃圾性质采用不同方式处理"},
            ],
            "palette": "emerald",
            "seed": seed + 4,
        },
        # 6. Section header - Why
        {
            "type": "cn_task_separator",
            "text": "02\n为什么要实施生活垃圾分类",
            "palette": "ocean",
            "seed": seed + 5,
        },
        # 7. Why - Reasons
        {
            "type": "cn_content",
            "side_label": "原因",
            "heading": "为什么要实施生活垃圾分类？",
            "body": [
                "🔹 生活垃圾逐年增多，传统填埋方式难以为继",
                "🔹 垃圾分类可有效减少环境污染",
                "🔹 可回收物再利用，节约自然资源",
                "🔹 有害垃圾安全处理，保护公众健康",
                "🔹 厨余垃圾堆肥处理，变废为宝",
                "🔹 推动绿色生活方式，建设美丽中国",
            ],
            "palette": "ocean",
            "seed": seed + 6,
        },
        # 8. Data/impact
        {
            "type": "metrics",
            "heading": "垃圾分类的意义",
            "metrics": [
                {"value": "30%", "label": "可回收利用", "sublabel": "垃圾中有30%可回收"},
                {"value": "50%", "label": "减少填埋", "sublabel": "分类后填埋量减半"},
                {"value": "0.3吨", "label": "有机肥料", "sublabel": "每吨厨余可产0.3吨肥料"},
            ],
            "palette": "ocean",
            "seed": seed + 7,
        },
        # 9. Section header - How
        {
            "type": "cn_task_separator",
            "text": "03\n垃圾分类的投放指导",
            "palette": "coral",
            "seed": seed + 8,
        },
        # 10. Four categories
        {
            "type": "cn_content",
            "side_label": "分类",
            "heading": "四大分类标准",
            "body": [
                "🔵 可回收物 — 废纸、塑料、玻璃、金属、布料",
                "🔴 有害垃圾 — 废电池、废灯管、废水银温度计、过期药品",
                "🟢 厨余垃圾 — 剩菜剩饭、骨头、菜根菜叶、果皮",
                "⚫ 其他垃圾 — 砖瓦陶瓷、渣土、卫生间废纸、纸巾",
            ],
            "palette": "coral",
            "seed": seed + 9,
        },
        # 11. Recyclable
        {
            "type": "cn_reading",
            "title": "可回收物",
            "passage": "    主要包括废纸、塑料、玻璃、金属和布料五大类。\n\n    废纸主要包括报纸、期刊、图书、各种包装纸、办公用纸、广告纸、纸盒等等。但是要注意，纸巾和厕所纸由于水溶性太强不可回收。\n\n    塑料主要包括各种塑料袋、塑料包装物、一次性塑料餐盒和餐具、牙刷、杯子、矿泉水瓶等。\n\n    玻璃主要包括各种玻璃瓶、碎玻璃片、镜子、暖瓶等。\n\n    金属物主要包括易拉罐、罐头盒、牙膏皮等。",
            "sidebar": [
                "纸类：报纸、纸箱、书本",
                "塑料：饮料瓶、玩具",
                "玻璃：酒瓶、玻璃杯",
                "金属：易拉罐、铁锅",
                "织物：旧衣服、毛巾",
                "注意：受污染的不可回收",
            ],
            "palette": "emerald",
            "seed": seed + 10,
        },
        # 12. Harmful waste
        {
            "type": "cn_reading",
            "title": "有害垃圾",
            "passage": "    主要包括废电池、废日光灯管、废水银温度计、过期药品等，这些垃圾需要进行特殊安全处理。\n\n    一节废旧电池能污染60万升水，相当于一个人一生的饮水量。一节电池烂在地里，能使一平方米的土地失去利用价值。\n\n    过期药品随意丢弃会造成土壤和水源污染，危害人体健康。",
            "sidebar": [
                "废电池（充电电池、纽扣电池）",
                "废荧光灯管",
                "废水银温度计",
                "过期药品及包装",
                "油漆桶、杀虫剂",
                "需专门安全处理",
            ],
            "palette": "crimson",
            "seed": seed + 11,
        },
        # 13. Kitchen waste
        {
            "type": "cn_reading",
            "title": "厨余垃圾",
            "passage": "    主要包括剩菜剩饭、骨头、菜根菜叶、果皮等食品类废物。\n\n    经生物技术就地处理堆肥，每吨可生产约0.3吨有机肥料。\n\n    投放要求：纯流质食物垃圾，如牛奶等，应直接倒进下水口；有包装物的厨余垃圾应将包装物去除后分类投放，包装物投放到对应的可回收物或其他垃圾容器中。",
            "sidebar": [
                "剩菜剩饭、菜叶果皮",
                "蛋壳、鱼骨、虾壳",
                "茶渣、中药渣",
                "瓜皮果核",
                "需沥干水分投放",
                "可堆肥处理变有机肥",
            ],
            "palette": "forest",
            "seed": seed + 12,
        },
        # 14. Other waste
        {
            "type": "cn_reading",
            "title": "其他垃圾",
            "passage": "    主要包括除上述几类垃圾之外的砖瓦陶瓷、渣土、卫生间废纸、纸巾等难以回收的废弃物。\n\n    采取卫生填埋可有效减少对地下水、地表水、土壤及空气的污染。\n\n    事实上，大棒骨因为「难腐蚀」被列入其他垃圾。类似的还有玉米核、坚果壳、果核等。",
            "sidebar": [
                "砖瓦陶瓷、渣土",
                "卫生间废纸、纸巾",
                "大棒骨、坚果壳",
                "一次性餐具",
                "烟蒂、灰尘",
                "采取卫生填埋处理",
            ],
            "palette": "mono",
            "seed": seed + 13,
        },
        # 15. Tips
        {
            "type": "two_column",
            "heading": "投放小贴士",
            "left": {
                "title": "✅ 正确做法",
                "items": [
                    "投放前先分类",
                    "厨余垃圾沥干水分",
                    "可回收物保持清洁干燥",
                    "有害垃圾轻放",
                    "尖锐物品包裹后投放",
                ],
            },
            "right": {
                "title": "❌ 常见错误",
                "items": [
                    "所有垃圾混在一起丢",
                    "纸巾当作可回收物",
                    "大骨头丢进厨余垃圾",
                    "电池丢进普通垃圾桶",
                    "塑料瓶不倒空直接丢",
                ],
            },
            "palette": "golden",
            "seed": seed + 14,
        },
        # 16. Section - System
        {
            "type": "cn_task_separator",
            "text": "04\n生活垃圾分类的体系建设",
            "palette": "lavender",
            "seed": seed + 15,
        },
        # 17. System overview
        {
            "type": "cn_content",
            "side_label": "体系",
            "heading": "分类体系建设",
            "body": [
                "📋 法规制度 — 完善垃圾分类相关法律法规",
                "🏗️ 设施建设 — 合理布局分类投放收集设施",
                "🚛 运输体系 — 建立分类收运体系，杜绝混装混运",
                "🏭 处理能力 — 提升分类处理设施能力和水平",
                "📢 宣传教育 — 加强社会宣传，提升全民意识",
                "📊 监督考核 — 建立健全监督考核机制",
            ],
            "palette": "lavender",
            "seed": seed + 16,
        },
        # 18. What we can do
        {
            "type": "cn_content",
            "side_label": "行动",
            "heading": "我们能做什么？",
            "body": [
                "🌍 从自身做起，养成垃圾分类好习惯",
                "👨‍👩‍👧‍👦 带动家人朋友一起参与",
                "📖 学习垃圾分类知识，正确投放",
                "🛒 减少使用一次性用品",
                "♻️ 旧物再利用，践行低碳生活",
                "📢 积极宣传，争做环保先锋",
            ],
            "palette": "mint",
            "seed": seed + 17,
        },
        # 19. End
        {
            "type": "cn_title",
            "heading": "谢谢大家",
            "subheading": "垃圾分类 从我做起",
            "bg_style": "bokeh",
            "palette": "emerald",
            "seed": seed + 18,
        },
    ],
}

output = "垃圾分类_从我做起.pptx"
build_ppt(config, output)
print(f"✅ Generated: {output}")
print(f"   Slides: {len(config['slides'])}")
print(f"   Seed: {seed}")
