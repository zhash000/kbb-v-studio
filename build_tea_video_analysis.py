from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = r"E:\cursor\kbb-v\茶叶风暴单视频文案画面配合分析.docx"
doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(.62); sec.bottom_margin = Inches(.62)
sec.left_margin = Inches(.65); sec.right_margin = Inches(.65)

normal = doc.styles["Normal"]
normal.font.name = "Microsoft YaHei"; normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
normal.font.size = Pt(9.5); normal.font.color.rgb = RGBColor(40,48,52)
normal.paragraph_format.space_after = Pt(5); normal.paragraph_format.line_spacing = 1.12
for nm, size, col in [("Title",24,"20383A"),("Heading 1",15,"20383A"),("Heading 2",11.5,"28645F"),("Heading 3",10,"28645F")]:
    s=doc.styles[nm]; s.font.name="Microsoft YaHei"; s._element.rPr.rFonts.set(qn("w:eastAsia"),"Microsoft YaHei")
    s.font.size=Pt(size); s.font.bold=(nm!="Title"); s.font.color.rgb=RGBColor.from_string(col)
    s.paragraph_format.space_before=Pt(10); s.paragraph_format.space_after=Pt(5); s.paragraph_format.keep_with_next=True
tp=doc.styles["Title"]._element.get_or_add_pPr(); bd=tp.find(qn("w:pBdr"));
if bd is not None: tp.remove(bd)
footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
rr=footer.add_run("茶叶风暴单视频分析  |  "); rr.font.size=Pt(8); rr.font.color.rgb=RGBColor(115,126,130)
f=OxmlElement("w:fldSimple"); f.set(qn("w:instr"),"PAGE"); footer._p.append(f)

def p(text="", style=None, lead=None):
    x=doc.add_paragraph(style=style)
    if lead and text.startswith(lead): x.add_run(lead).bold=True; x.add_run(text[len(lead):])
    else: x.add_run(text)
    return x
def h(text, level=1): doc.add_heading(text,level)
def bullet(text):
    x=doc.add_paragraph(style="List Bullet"); x.paragraph_format.space_after=Pt(2); x.add_run(text)
def table(headers, rows, widths=None, fs=8):
    t=doc.add_table(rows=1,cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=False; t.style="Table Grid"
    for i,txt in enumerate(headers):
        c=t.rows[0].cells[i]; c.text=txt; c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        sh=OxmlElement("w:shd"); sh.set(qn("w:fill"),"28645F"); c._tc.get_or_add_tcPr().append(sh)
        for r in c.paragraphs[0].runs: r.bold=True; r.font.size=Pt(fs); r.font.color.rgb=RGBColor(255,255,255)
    for row in rows:
        cells=t.add_row().cells
        for i,txt in enumerate(row):
            cells[i].text=txt; cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for pp in cells[i].paragraphs:
                pp.paragraph_format.space_after=Pt(1); pp.paragraph_format.line_spacing=1.05
                for r in pp.runs: r.font.size=Pt(fs)
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Inches(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(1)
    return t

title=doc.add_paragraph(style="Title"); title.add_run("茶叶风暴单视频文案画面配合分析")
sub=doc.add_paragraph(); sub.paragraph_format.space_after=Pt(9)
r=sub.add_run("样本：茶叶风暴 一片不起眼的树叶 如何搅动世界  |  时长约 7 分 13 秒")
r.font.size=Pt(9); r.font.color.rgb=RGBColor(93,110,112)

h("分析结论",1)
p("这条视频的画面不是给旁白配一组泛化的历史插图，而是随着叙事尺度逐级改变：从饮茶人物与叶片开始，进入加工动作和运输路线，再转向欧洲消费、帝国贸易、植物采集与跨洋运输，最后回到茶叶和远航意象。文案负责建立因果链，画面负责把抽象因果变成可见的人、物、动作、空间和风险。")
p("最值得学习的是局部配画的功能性：提到成分时用微距和微观意象，讲加工时给手部工序，讲贸易时给货箱、港口和船，讲历史行动时给人物任务线。整体上，小物件逐渐被放大成全球系统；结尾再把全球系统收回到一片叶子。")
p("证据边界：视频约 433 秒，720×1280 竖屏、30 fps。以下时间码为分段抽帧定位的近似范围；文案依据可读字幕和画面语义概述，并非逐字转录。‘为什么配这个画面’是叙事功能判断，不代表已知作者真实意图。历史重演或生成画面不能作为史实证据。")

h("一 整体叙事地图",1)
table(["近似位置","叙事段落","文案推进","主要画面系统","段落作用"],[
 ["0:00–约1:00","熟悉入口与问题","从古代饮茶习惯引出茶为何重要","古装饮茶人物、茶叶与茶树微距、蜜蜂角色","让观众先认出日常对象，再产生‘一片叶子为何影响这么大’的疑问"],
 ["约1:00–2:00","自然属性与使用价值","介绍茶叶成分、提神经验及加工基础","叶片、微观可视化、饮茶、炒制与作坊动作","解释茶为什么被消费、加工后如何成为有价值的商品"],
 ["约2:00–3:00","从生产到流通","讲制茶、茶马古道和远距离运输","揉捻炒制、驮队、山路、货物、港口/船只","把商品价值转化为路线、成本和运输能力"],
 ["约3:00–5:00","跨国需求与贸易竞争","进入欧洲饮茶、英国贸易利益及茶种/制茶知识获取","欧洲饮茶、商人、船队、东印度公司、Robert Fortune相关重演","由消费习惯升级为贸易、知识和帝国利益的冲突"],
 ["约5:00–6:30","社会影响与生产版图","扩展到工业城市、公共生活和种植扩张","工业城市与水源、茶园、植物运输箱、种植与加工","让影响超出商品本身，进入社会生活和全球产业链"],
 ["约6:30–7:13","回扣标题","将茶叶贸易重新概括为全球性影响","茶叶、茶箱、航海者、远洋船与风浪","把宏大历史压缩回具体物件，兑现标题承诺"],
],[.72,1.06,1.62,1.73,2.02],7.5)
p("整体节奏靠叙事尺度变大来推进，而不只是靠更快切镜：个人饮茶 → 叶片与加工 → 商路 → 跨国交易 → 帝国和社会 → 全球传播。观众每跨一段，都应更清楚茶如何从普通消费品变成更大系统中的关键商品。")

h("二 局部文案与画面配合",1)
table(["位置与文案信息","画面观察","配合关系","为什么这样配","整体作用"],[
 ["开场：古代饮茶观念/茶与人的生活","古装人物室内饮茶、交谈，茶具入镜","对应、设定","先用人物和动作让观众看见饮茶习惯与文化身份；比直接上地图或年份更容易进入","把主题落在生活经验，避免知识片头过于抽象"],
 ["开场提问：茶为什么不只是简单饮品（字幕语义概述）","人物惊讶反应，随后出现叶片、茶树与蜜蜂","设问、具象化","人物反应替观众显出疑问；叶片特写让问题回到‘一片叶子’；蜜蜂作为轻松视觉引导，降低解释段落的距离感","形成开场认知缺口，并建立可重复的视觉锚点"],
 ["字幕涉及咖啡因、儿茶素等茶叶成分","茶叶微距、茶园采摘、昆虫/蜜蜂角色","解释、补充","肉眼看不见成分，叶片微距把概念附着到具体植物；采摘动作把科学信息接回生产活动。蜜蜂不是科学证据，只是角色化提示","为后续消费需求与茶叶价值提供基础"],
 ["字幕谈饮茶的提神作用","人物饮茶、生活场景与身体状态","生活化补充","把抽象的作用翻译成观众熟悉的使用场景，提供‘人为什么需要它’的直觉","从成分过渡到消费动机；具体生理因果仍需谨慎表述"],
 ["字幕讲制茶加工和品质形成","手部揉捻、炒制、筛选，茶叶从鲜叶变成成品","过程证明、感官化","动作展示工序、劳动和物料变化；比只有茶山空镜更能说明商品价值是如何被加工出来的","连接自然植物与可交易商品，是进入贸易故事的因果桥"],
 ["字幕谈茶马古道、长途储运或路线延伸","驮队穿越山路，大全景交代地势，货物/人物近景补细节","对应、尺度扩展","山路使距离、时间和运输难度可见；远景给规模，近景给人的劳动与货物重量","把茶的故事从产地扩展到交通网络"],
 ["字幕转到欧洲饮茶习惯和船只运茶","货箱标记茶类、船员搬运，切到欧洲贵族饮茶","跨文化对照、转场","同一件商品在产地运输和海外消费场景间切换，观众能直接感到茶跨过文化边界；货箱与船只保留物流因果","完成从地方文化到跨国商品的关键转折"],
 ["字幕谈英国市场需求、贸易公司或获取知识的行动","欧洲饮茶、商人/公司场景、契约或任务、人物观察茶叶","人物锚点、因果可视化","市场需求和制度很抽象；把它落到商人、任务、货物和植物样本上，观众可以跟随一个具体行动理解系统如何运转","把宏观贸易变成可追踪的人物行动线"],
 ["字幕提到 Robert Fortune 与茶树/制茶知识转移","人物考察茶叶、植物样本和玻璃运输箱等重演画面","人物锚点、过程解释","‘植物采集与知识转移’抽象，样本、容器和人物动作能解释运输对象是活植物及其操作难点。此处影视画面只能说明叙事，不证明细节真实","形成中段任务线，也带来需要核实的历史节点"],
 ["字幕涉及长途海运风险和茶叶远销","船上货箱、风浪、船员和远景船队","风险具象化、情绪升级","风浪、封闭货箱和人的处境把‘航运风险’变成能感受到的空间与身体压力；同时让贸易链条有过程证据","提高中后段情绪张力，为结尾全球化意象蓄势"],
 ["结尾将茶叶与全球性风暴/世界影响重新联系","人物抱着茶箱，之后切远航船只和海面","象征、回扣","人物怀抱货物把全球商品重新变成具体重量；远航画面再把个体行动放大到世界尺度。只有地图会宏大但少情绪，只有人物则难回收世界影响","闭合标题承诺：从一片叶子回到全球历史"],
],[1.44,1.40,.91,1.76,1.64],7.0)

h("三 配画逻辑的局部规律",1)
for x in [
 "抽象名词寻找可视化代理：成分对应叶片/微观结构，贸易对应货物/港口/航线，风险对应海况和运输环境。",
 "因果链尽量用动作拍出来：采摘、加工、装箱、运输、观察、种植，比纯空镜更能说明事情怎样发生。",
 "镜头尺度随论述尺度变化：微距解释叶片，人物镜头承载选择和劳动，大全景交代路线、港口和世界系统。",
 "人物作为系统入口：代表性人物可以把制度和贸易变成行动，但角色重演不能代替资料来源。",
 "重复符号做段落连接：茶叶和蜜蜂可作为视觉回钩；每次出现仍应与本段信息相关，否则会变成装饰。",
 "情绪镜头应有因果位置：风浪、惊讶或冲突镜头最好对应真实叙事升级，不只为了画面刺激。",
]: bullet(x)

h("四 历史时间线与叙事压缩",1)
p("这条视频把多个世纪的变化剪成连续的观看旅程。剪辑上的压缩有利于节奏，但创作复述时要用时间提示区分阶段，不能让相邻镜头造成‘同一时期连续发生’的错觉。")
table(["历史节点","可核对的时间信息","剪辑/文案处理建议"],[
 ["茶进入英国并形成宫廷风尚","Royal Collection Trust称茶于1650年代出现在英国；1662年后在宫廷流行；东印度公司1664年下首个茶叶订单。","在欧洲饮茶画面前后明确17世纪背景。"],
 ["Assam本地茶树知识与早期研究","Kew记载Robert Bruce从当地Singpho人士获知Assam茶树，Charles Bruce后续培育并送检，1834年得到确认。","保留当地知识主体与英国茶业早期发展的多方贡献。"],
 ["Robert Fortune赴中国采集茶树","RHS Digital Collections记载东印度公司于1848年委托Fortune从中国采集茶树，以引入印度；Kew也说明其行动发生在Assam茶业已有早期发展的背景中。","不能写成‘Fortune第一次把茶带到印度’或‘英国直到他才知道如何种茶’。可用‘近两个世纪后，英国又把目光投向茶树本身’做时间桥。"],
],[1.42,2.45,3.1],8)
p("事实核验的叙事意义：英国早期饮茶、Assam本地茶树和Fortune的中国采集任务之间有长期历史关联，但不是一个连续事件。文案若用一个人物覆盖全部茶产业起源，会压低当地知识和早期种植工作的作用。")

h("五 复用到新视频的检查表",1)
for x in [
 "整条视频的主问题能否用一句话说清？每个章节是否推进它？",
 "每段文案旁边是否写了画面功能，而不只是素材描述？",
 "画面是在对应、补充、证明、对照、转场、象征还是留白？为什么这个关系适合此处？",
 "如果换成普通空镜，会损失什么信息、因果、情绪或尺度？",
 "历史、科学或职业事实是否核验过？生成画面会不会被误认为真实档案？",
 "跨年代、跨地点和跨因果跳转是否有明确的语言或画面桥接？",
 "结尾是否回收了开头的问题，同时提供新的理解？",
]: bullet(x)

h("资料说明",1)
p("本案例的画面与字幕观察来自用户提供的视频及分段抽帧；时间码为近似定位。历史节点核对参考：Royal Collection Trust, “Tea in the Royal Collection”；Royal Botanic Gardens, Kew, “The history of tea: From China to India”；RHS Digital Collections, “RHS Plant Collector Archive: Robert Fortune”。这些资料用于校正年代和人物背景，不代表已经逐字核验视频旁白。")

doc.core_properties.title="茶叶风暴单视频文案画面配合分析"
doc.core_properties.subject="整体叙事与局部文案画面关系拆解"
doc.core_properties.author="OpenAI"
doc.save(OUT)
print(OUT)
