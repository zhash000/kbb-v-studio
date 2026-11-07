from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = r"E:\cursor\kbb-v\同一账号抖音视频内容分析与创作框架.docx"

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(.72)
sec.bottom_margin = Inches(.7)
sec.left_margin = Inches(.78)
sec.right_margin = Inches(.78)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Microsoft YaHei"
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor(42, 49, 56)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.18
for nm, size, color in [("Title", 25, "1E3436"), ("Heading 1", 17, "1E3436"), ("Heading 2", 12.5, "28645F"), ("Heading 3", 10.5, "28645F")]:
    s = styles[nm]
    s.font.name = "Microsoft YaHei"
    s._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    s.font.size = Pt(size)
    s.font.bold = nm != "Title"
    s.font.color.rgb = RGBColor.from_string(color)
    s.paragraph_format.space_before = Pt(13 if nm != "Title" else 0)
    s.paragraph_format.space_after = Pt(6)
    s.paragraph_format.keep_with_next = True
# Keep Word's built-in Title style free of decorative paragraph borders.
title_ppr = styles["Title"]._element.get_or_add_pPr()
title_borders = title_ppr.find(qn("w:pBdr"))
if title_borders is not None:
    title_ppr.remove(title_borders)

footer = sec.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = footer.add_run("抖音同账号样本分析  |  ")
r.font.size = Pt(8); r.font.color.rgb = RGBColor(115, 126, 130)
fld = OxmlElement("w:fldSimple"); fld.set(qn("w:instr"), "PAGE")
footer._p.append(fld)

def para(text="", style=None, boldlead=None):
    p = doc.add_paragraph(style=style)
    if boldlead and text.startswith(boldlead):
        p.add_run(boldlead).bold = True
        p.add_run(text[len(boldlead):])
    else:
        p.add_run(text)
    return p

def bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    p.add_run(text)
    return p

def heading(text, level=1):
    doc.add_heading(text, level=level)

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = h
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for run in c.paragraphs[0].runs:
            run.bold = True; run.font.color.rgb = RGBColor(255,255,255); run.font.size = Pt(9)
        shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), "28645F"); c._tc.get_or_add_tcPr().append(shd)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = val
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs: run.font.size = Pt(9)
    if widths:
        for row in t.rows:
            for i, width in enumerate(widths): row.cells[i].width = Inches(width)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t

title = doc.add_paragraph(style="Title")
title.add_run("同一账号抖音视频内容分析与创作框架")
sub = doc.add_paragraph()
sub.paragraph_format.space_after = Pt(12)
r = sub.add_run("样本：《茶叶风暴：一片不起眼的树叶，如何搅动世界？》与《护理：一把斩向死亡的温柔刀》")
r.font.size = Pt(10); r.font.color.rgb = RGBColor(95, 111, 113)

heading("分析结论", 1)
para("两条样本都把知识题材组织成电影化的叙事短纪录片：先用反差标题制造问题，再沿着人物、事件和时代变化推进，最后把一个具体对象提升为更大的历史或人文意义。可复用的核心不是原片台词或镜头，而是“普通入口—冲突升级—视听证明—意义回收”的讲述方法。")
para("本文依据用户提供的两段视频进行观察。视频时长约 7 分 13 秒与 7 分 01 秒，均为 720×1280 竖屏、30 fps，画面中可见中英双语字幕，并以旁白配合情境化影像。由于本次没有完成逐字语音转写和逐帧剪辑检测，关于具体台词、音乐曲目、切镜频率和音效位置的描述不作精确断言。", boldlead="本文依据")
para("阅读标记：‘观察’指样本画面或文件属性直接支持；‘推断’指由标题和叙事线索判断的受众或效果；‘建议’指后续创作方案。两条视频只能帮助提炼账号风格，不能单独证明抖音推荐机制或普遍爆款规律。")

heading("一 账号内容定位", 1)
para("从题材跨度看，一条讲茶叶与世界史，一条讲护理职业与医疗发展；共同点并非某个垂直知识领域，而是把熟悉的对象放进宏大叙事。可暂将账号定位概括为：用电影化视觉讲述日常事物、职业或制度背后的历史影响与文明价值。")
table(["维度", "茶叶样本", "护理样本", "账号共性"], [
    ["题材入口", "日常饮品与一片叶子", "护理职业与医院场景", "从具体、易懂的对象开始"],
    ["主要张力", "小物件与世界格局的反差", "温柔照护与对抗死亡的反差", "标题提出反常识命题"],
    ["叙事尺度", "贸易、帝国、航海与社会变化", "职业史、医疗技术与生命照护", "由个人或物件扩大到时代"],
    ["情绪方向", "好奇、惊奇、历史纵深", "敬意、共情、职业价值", "知识信息与情绪回报并行"],
])

heading("二 题材与受众", 1)
heading("题材选择逻辑", 2)
para("两条内容都选择了‘熟悉但被低估’的入口。茶是日常消费品，护理是常见职业；标题再把它们连接到更大的命题，形成认知落差。此类题材适合做系列，因为对象可以持续替换，而叙事方法保持稳定。")
heading("可能的受众", 2)
para("以下是内容定位推断，不是平台后台人群数据。茶叶视频可能吸引历史、贸易、文明交流和生活文化兴趣者；护理视频可能吸引医护与医学生、患者及照护者、职业史和女性史兴趣者，也容易触发对医疗现场有亲身经验的观众。两者的交集是喜欢有知识增量、有人物命运和电影感呈现的泛知识用户。")
heading("选题筛选问题", 2)
for x in [
    "对象是否足够熟悉，观众能否在一句话内理解？",
    "它是否能连接一个更大的冲突：权力、贸易、技术、制度、身份或生命？",
    "能否找到人物、事件和可视化证据，避免全片停留在抽象概念？",
    "是否有可靠资料支撑关键历史或职业事实？",
]: bullet(x)

heading("三 文案结构与叙事", 1)
para("从标题和画面序列看，两条内容都采用‘问题先行、背景后置’的叙事倾向：先让观众知道这个对象不普通，再逐步解释它如何变得重要。完整文案可按以下功能拆分；这是对样本结构的归纳，不是逐字稿还原。")
table(["段落功能", "作用", "写作要点"], [
    ["开场钩子", "制造认知缺口", "用反差或结果开场，避免先报年份和背景"],
    ["对象落地", "让观众看懂主角", "一句话交代对象、常识认知与核心问题"],
    ["起源铺垫", "建立理解基础", "只保留后续冲突必需的背景"],
    ["冲突升级", "推动继续观看", "引入人物、利益、障碍、竞争或技术变化"],
    ["影响扩展", "增加叙事尺度", "从个体扩展到城市、行业、国家或制度"],
    ["意义回收", "完成情绪与观点闭环", "回到开头对象，说明今天为何仍值得关注"],
])
para("建议每一段都承担一个明确功能：提供新事实、改变观众判断、提出下一问，或给出情绪回报。删除只重复前文、没有推进作用的句子。长视频中可在章节转换处用小问题、因果悬念或时间跳转续接注意力。")

heading("四 画面语言", 1)
para("观察到的画面不是单一素材类型：包括历史场景、人物行动、物件特写、环境建立镜头、医疗设备与病房场景；两条中都有较强的情境化、电影化生成画面。茶叶样本还有卡通蜜蜂这一重复出现的视觉角色，可帮助串联知识段落并减轻历史叙事的距离感。护理样本则以护士、病床、手术室和医疗设备建立职业与技术的视觉识别。")
table(["画面类型", "适合承担的功能", "对应例子"], [
    ["建立场景", "交代时代、地点和规模", "茶园、港口、医院走廊"],
    ["人物行动", "把历史或职业变成可见事件", "采茶、航海、护理和医疗操作"],
    ["关键物件", "突出知识点并便于记忆", "茶叶、茶具、护理笔记、呼吸机"],
    ["象征镜头", "压缩抽象概念或情绪", "远航船队、空病床、时代转换"],
    ["重复视觉符号", "增强系列识别与段落连接", "茶叶片、蜜蜂角色、护士制服"],
])
para("画面建议：优先让镜头回应旁白中的名词、动作和因果；在抽象观点出现时，用人物选择、关键物件或结果场景承接。避免连续使用相似构图，也避免用漂亮但无关的素材填满时间。历史和医疗题材都应检查生成画面的事实感，不让虚构细节被误认作真实档案。")

heading("五 声音设计", 1)
para("样本具有旁白叙事和双语字幕这一可见特征；本次没有可靠地拆分人声、配乐和音效轨，因此不对具体配乐风格、音色、语速或混音关系下定论。声音维度仍应独立分析，因为它决定信息理解、情绪方向和画面节奏。")
table(["声音层", "分析与制作关注点"], [
    ["旁白", "语速是否给关键信息留出理解时间；重音是否落在冲突、转折和结论上；专有名词是否清楚。"],
    ["配乐", "是否随着段落改变情绪和能量；是否压住人声；是否在转折处提供推进，而非全程同强度铺底。"],
    ["音效与环境声", "船、火、器械、病房等声音是否帮助建立空间和动作；是否克制使用，避免堆叠造成廉价感。"],
    ["静默与停顿", "重要观点前后是否留有短暂呼吸；是否用停顿放大结果，而不是把每秒都填满。"],
    ["字幕可读性", "双语字幕会增加信息量；中文应优先清晰，英文作为辅助，字号、行数和停留时间需适配手机。"],
])

heading("六 文案 画面与声音如何协同", 1)
para("这是最值得单独拆分的一类，因为视频的吸引力通常来自多轨信息相互补充，而不是三类元素分别漂亮。可用‘主关系’判断每个镜头是否有效：")
table(["关系", "做法", "适用场景"], [
    ["对应", "画面直接呈现旁白提到的人、物或地点", "核心知识点首次出现"],
    ["补充", "画面提供旁白未说出的细节", "人物动作、生活环境、时代物件"],
    ["证明", "用过程或结果让观点变得可信", "技术如何改变护理流程、贸易如何扩张"],
    ["对照", "画面与旁白形成过去/现在、强/弱或理想/现实对比", "职业演变、技术进步、社会反差"],
    ["留白", "旁白暂退，让音乐、环境声或关键画面承接情绪", "牺牲、死亡、重大转折后的收束"],
])
para("每段可按以下顺序检查：旁白提出什么信息或问题；画面是在重复、补充还是证明；声音把观众的注意力引向哪里；字幕是否突出需要记住的词；镜头结束时是否留下下一段的动机。若画面只是重复念出的句子，信息效率偏低；若画面与旁白无关，理解链条会断开。")

heading("七 剪辑节奏与观看推进", 1)
para("从 18 秒间隔的抽样画面可以确认两条视频在较短区间内更换了多个场景、人物和物件，但该抽样不足以准确计算平均镜头时长或切镜频率。更可靠的结论是：叙事依赖场景变化和信息节点推进，而不是固定节拍的快速切镜。")
for x in [
    "用章节节奏控制信息：开头快给问题，中段交代必要背景，冲突处加速，关键结论处放慢。",
    "每一段至少发生一种变化：新人物、新证据、新地点、新时间，或观众对事件的新判断。",
    "转场要说明关系：因果、时间跳跃、立场变化或过去与现在的对照。",
    "避免为了‘抖音节奏’而持续快切；长视频仍需要呼吸、解释和情绪停留。",
]: bullet(x)

heading("八 抖音适配与效果判断", 1)
para("抖音适配建议从观众行为设计，而不是假设存在某个固定推荐阈值。两条样本时长都约 7 分钟，意味着内容必须持续兑现标题承诺：观众每看一段，都要获得新信息或新的情绪理由。可优先检查开头是否清楚、章节是否有推进、字幕是否适合手机观看、结尾是否形成完整观点。")
para("发布后应结合账号后台数据验证：前段流失位置、平均播放时长、完播情况、互动和关注转化；再对照对应时间点的文案、画面和声音变化。不要仅凭播放量判断哪种元素有效，也不要把单条数据直接归因于某一个镜头。平台分发规则会变化，创作时应以当期平台公开规则和账号实际数据为准。")
para("合规与原创：复用叙事结构和一般创作方法，不复制原片台词、分镜顺序、独特角色形象、音乐或未经授权的素材。历史事实、医疗信息和职业描述应核对来源；尤其医疗内容应避免把戏剧化表达写成具体诊疗建议或夸大效果。")

heading("九 可复用的原创制作模板", 1)
table(["环节", "交付内容", "检查问题"], [
    ["选题", "对象、核心问题、目标观众", "熟悉对象背后是否有真实冲突？"],
    ["资料", "时间线、关键人物、证据来源", "关键事实能否核实？有没有争议或不确定性？"],
    ["文案", "钩子、背景、冲突、影响、收束", "每段是否推进问题并兑现承诺？"],
    ["分镜", "逐句旁白、画面意图、字幕重点", "画面是在对应、补充、证明还是对照？"],
    ["声音", "旁白节奏、配乐情绪、音效位置", "人声是否清晰，声音是否帮助理解？"],
    ["剪辑", "段落长度、转场理由、情绪停留", "是否有重复信息或无效镜头？"],
    ["复盘", "后台曲线与评论反馈", "观众在哪离开、停留、讨论或关注？"],
])
heading("开场句式示例", 2)
para("“你每天接触的【普通对象】，曾经改变过【更大范围的事情】。”")
para("“今天我们熟悉的【职业或制度】，最初解决的其实是【具体危机】。”")
para("“看起来只是【日常动作】，背后却藏着一场关于【权力、技术或生命】的改变。”")
para("句式只提供结构，不是保证流量的公式。正式写作时应换成有来源支撑的具体事实，避免空泛夸张。建议先确定‘普通物件如何改变世界’或‘一个职业如何改变现代生活’这样的系列主线，再选择新题材试作完整旁白、分镜与声音方案，并依据发布后的留存曲线复盘。")

heading("十 主题正确性与标志性事件套用原则", 1)
para("后续创作应把‘内容正确’放在‘表达好看’之前。选定主题后，先确认它的核心事实、时间范围、因果关系和争议边界，再决定用什么故事讲。电影化画面可以承担解释和情绪功能，但不能把虚构的细节包装成真实历史、科学结论或医疗事实。")
table(["步骤", "要做什么", "交付检查"], [
    ["确定主题", "写清对象、核心问题和希望观众理解的结论", "一句话能说明视频要解释什么"],
    ["核对事实", "查找可靠来源，确认时间、人物、数据和因果关系", "关键事实至少有可追溯来源"],
    ["划分边界", "区分确定事实、学界争议、合理推断和艺术化重演", "文案不把推断写成定论"],
    ["选择锚点", "挑选与主题真正相关的标志性事件、人物、物件或场景", "锚点能推进问题，而非只负责好看"],
    ["套用公式", "用‘熟悉入口—反常识问题—可视化解释—冲突升级—意义回收’组织材料", "每段都有新信息或新的判断变化"],
    ["视听核验", "逐句检查旁白、画面、字幕和声音是否表达同一事实关系", "画面不制造与文案相反的误导"],
])
para("标志性事件不是装饰性案例，而是叙事的支点。它应至少完成一种任务：证明主题确实产生过影响、展示冲突如何发生、让抽象机制变得可见，或把个人经验连接到更大的历史和社会背景。若一个事件只因为知名而与主题关系很弱，应舍弃或改成背景提及。")
para("推荐的写作顺序是：先列出事实时间线，再从时间线中选出一到三个最能承载冲突的锚点，最后把锚点重新编排成适合观看的节奏。这样可以同时保留知识准确性和叙事张力。")

doc.core_properties.title = "同一账号抖音视频内容分析与创作框架"
doc.core_properties.subject = "茶叶与护理两条视频样本拆解"
doc.core_properties.author = "OpenAI"
doc.save(OUT)
print(OUT)
