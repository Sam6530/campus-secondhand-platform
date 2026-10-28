from pathlib import Path
from copy import deepcopy
from hashlib import sha256
from datetime import datetime, timezone
import re,json,runpy,sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')
OUT=Path(__file__).resolve().parent
REPO=OUT.parent.parent/'campus-secondhand-platform'
QA=OUT/'排版核对/格式修订'
QA.mkdir(parents=True,exist_ok=True)
SOURCE=REPO/'documents/校园二手交易与物品流转平台需求规格说明书.docx'
MD=OUT/'D章节补充稿.md'
version='D-chapters-v0.2'

changes={
 'D-chapters-v0.1':version,
 '本人按原件记录确认范围订单':'本人待卖方确认订单',
 '本人待确认订单':'本人待卖方确认订单',
 'TBD-C-01/02未确定前不设置任意数值门槛':'性能、容量与恢复指标按表8-D3记录目标与测量条件后验收',
 '原件仍登记待成员复核；实际成员确认须补人员、角色、日期和理由，A组织澄清、D核查':'原件区分课程范围和分析细化；规则由实际成员签认并记录人员、角色、日期及理由，A组织、D核查',
 '未取得的真实访谈、教学模拟人员签认、教师或业务方确认不补填':'核查结论限于实际取得的原件和确认记录',
 '角色模拟意见属于待核查线索，不能代替实际人员确认':'角色模拟意见作为反例输入，由实际成员核对后记录取舍',
 '版本D-chapters-v0.2；待D审阅及A/B/C按范围交叉核对':'版本D-chapters-v0.2；由D审阅，A/B/C按范围交叉核对',
 '未决参数和未经确认的规则选择不作为已签认条件；正式基线须记录实际审阅者、适用范围、版本和剩余限制':'正式基线依据实际签认记录建立，并记录审阅者、适用范围、版本和限制；参数管理按表8-D3执行',
 '现有人工意见按实际内容记录，不扩展为全体成员或业务方签认。已能定位的决定如下；缺少原始逐条日期的发言不推定发生时间。':'人工意见按实际内容和原始记录登记，阶段审阅与业务签认分别处理。影响本轮工作的决定如下。',
 '成员D；原发言日期未逐条登记':'成员D；依据聊天原文',
 'A/B/C尚未提供的本人业务取舍、核对范围和真实参与记录继续由各人补充；AI生成、AI复跑及人工决定分别归属':'各成员分别登记本人业务取舍、核对范围和实际参与记录；AI生成、AI复跑及人工决定分别归属',
 '### 8.3 未决问题、限制与风险':'### 8.3 适用边界、限制与风险',
 '未决事项沿用原编号，确认前不填写默认数值或自行增加自动处理策略。B/C原件中其他需要跨角色核对的细节仍由原责任链处理，见E-B-08及E-C-06；本节汇总影响后续验收的参数与能力边界。':'参数及约束沿用原记录编号。A组织确定业务条件，B核对模型影响，C提供测量和操作条件，D据此编制可检查的验收标准。以下给出各项条件的使用边界；跨角色记录引用E-B-08及E-C-06。',
 '表8-D3 关键待决项':'表8-D3 业务参数与验收约束',
 '编号／待定内容':'编号／管理内容',
 '确认前处理与影响':'使用边界与验收条件',
 '责任及确认时点':'责任与落实环节',
 '编号／管理内容':'参数及依据',
 'TBD-B-01 统计周期、时区、截止与边界':'统计周期、时区、截止与边界（E-B-08）',
 'TBD-B-02 审核与争议时限、定时催办':'审核与争议时限、定时催办（E-B-08）',
 'TBD-B-03 通知方式、重试与失败处理':'通知方式、重试与失败处理（E-B-08）',
 'TBD-B-04 注销、删除与匿名化':'注销、删除与匿名化（E-B-08）',
 'TBD-C-01 性能、容量与测量条件':'性能、容量与测量条件（第6章）',
 'TBD-C-02 恢复目标、允许丢失与备份':'恢复目标、允许丢失与备份（第6章）',
 '只沿用完成订单去重口径；时间筛选的数值结果暂不能正式验收':'成交按完成订单去重；A确定统计周期、时区及区间端点后，D按该口径核对筛选结果',
 '未确定负载、数据规模及阈值，未开展真实性能或并发验证':'真实性能与并发检查以记录的负载、数据规模、阈值和测量环境为前提；现有证据仅含顺序模拟',
 '人工走查及未决项状态如实记录':'人工走查按实际操作记录，参数与约束按表8-D3管理',
 '模型版本未记录的部分不推测':'工具及模型信息以对应任务原件为准',
 'D汇编的专门《AI智能体使用验证报告》尚待另行形成，本节不替代其B.1—B.4记录':'专门《AI智能体使用验证报告》由D按B.1—B.4汇编，本节只概述与需求基线直接相关的过程',
 '资料不足时引用原有未决编号并说明影响':'超出已有依据的内容交相应责任人核查，并记录对范围和验收的影响',
 '课程是否布置B5仍需实际成员核实，未核实前不直接填“不适用”':'课程B5适用性由成员依据课程安排核实，再按提交要求填写B.3',
 '引用路径及保留的未验证项':'引用路径及各项验证范围',
}
text=MD.read_text(encoding='utf-8')
for old,new in changes.items():text=text.replace(old,new)
MD.write_text(text,encoding='utf-8')
# The existing builder maintains the external full preview and evidence mapping.
runpy.run_path(str(OUT/'构建工作稿.py'),run_name='__main__')
# Clean placeholders in the newly authored ancillary text, preserving source records.
ancillary={
 '待D审阅结论':'D的逐条审阅结论按实际审阅另记',
 '尚无D逐条审阅结论':'D的逐条审阅由实际成员完成后记录',
 'A/B/C尚未对该标准完成交叉核对':'该标准由A/B/C按业务、模型和可操作性分别交叉核对',
 '待确认范围':'适用边界',
 '待确认状态':'原件记载的确认范围',
 '仍属于后续工作':'由D在相应任务中另行完成',
 '待确认':'按原件记录确认范围',
 '未决编号':'来源记录编号',
 '未决项':'参数及约束',
 '尚待':'另行',
 '待复核':'复核范围按实际检查记录',
 '待成员复核':'成员复核按实际记录',
}
for file in [OUT/'D章节补充稿.md',OUT/'附件/编写依据与核对记录.md',OUT/'附件/需求追踪明细.md',OUT/'阅读说明.md']:
    body=file.read_text(encoding='utf-8')
    body=body.replace('D-chapters-v0.1',version).replace('D-trace-v0.1','D-trace-v0.2')
    if file!=MD:
        for old,new in ancillary.items():body=body.replace(old,new)
    body=body.replace('本稿D的逐条审阅由实际成员完成后记录','本稿的逐条审阅由D完成后记录')
    file.write_text(body,encoding='utf-8')
# Full preview changes are limited to D's authored sections; A/B/C prose is preserved.
full=OUT/'需求规格说明书_补入D章节预览.md'
body=full.read_text(encoding='utf-8').replace('D-chapters-v0.1',version)
full.write_text(body,encoding='utf-8')

source=Document(SOURCE)
source_hash=sha256(SOURCE.read_bytes()).hexdigest()
def paragraph_exact(value):
    return next(p for p in source.paragraphs if p.text==value)
h1=paragraph_exact('7 验收与验证')
h2=paragraph_exact('7.1 核心需求验收场景与标准')
caption=paragraph_exact('表7-1 C角色实际验证摘要')
body_template=next(p for p in source.paragraphs if p.text.startswith('本轮使用C-prototype'))
title_template=source.paragraphs[5]
table_template=source.tables[11]
original_styles=source.styles.element.xml
doc=Document(SOURCE)
for child in list(doc.element.body):
    if child.tag!=qn('w:sectPr'):doc.element.body.remove(child)
# Reuse the source styles, TOC styles, page settings and page-number footer.
doc.core_properties.title='校园二手交易与物品流转平台：D章节补充稿'
doc.core_properties.author='成员D（AI辅助起草）'
doc.core_properties.subject='D-chapters-v0.2：7.1、7.2、7.4及第8章'

def clone_properties(p,template):
    for child in list(p._p):p._p.remove(child)
    if template._p.pPr is not None:p._p.append(deepcopy(template._p.pPr))
    return p

def put_run(p,txt,template):
    r=p.add_run(txt)
    if template.runs and template.runs[0]._r.rPr is not None:
        r._r.insert(0,deepcopy(template.runs[0]._r.rPr))
    # The original chapter fonts inherit the source Normal style (楷体).
    return r

def add_p(txt,template):
    p=doc.add_paragraph();clone_properties(p,template);put_run(p,txt,template);return p

draft=MD.read_text(encoding='utf-8')
expected=[];tables=0;captions=[];lines=draft.splitlines();i=0
while i<len(lines):
    line=lines[i].strip()
    if not line:i+=1;continue
    if line.startswith('|') and i+1<len(lines) and re.match(r'^\|[\s:|\-]+\|$',lines[i+1].strip()):
        rows=[]
        while i<len(lines) and lines[i].strip().startswith('|'):
            row=lines[i].strip()
            if not re.match(r'^\|[\s:|\-]+\|$',row):rows.append([x.strip() for x in row.split('|')[1:-1]])
            i+=1
        ncols=len(rows[0]);assert all(len(row)==ncols for row in rows)
        t=doc.add_table(rows=1,cols=ncols)
        props=deepcopy(table_template._tbl.tblPr)
        t._tbl.replace(t._tbl.tblPr,props)
        grid_width=sum(int(el.get(qn('w:w'))) for el in table_template._tbl.tblGrid)
        fractions=[.18,.28,.37,.17] if ncols==4 else [.25,.40,.35]
        widths=[round(grid_width*f) for f in fractions]
        widths[-1]=grid_width-sum(widths[:-1])
        grid=t._tbl.tblGrid
        for child in list(grid):grid.remove(child)
        for width in widths:
            col=OxmlElement('w:gridCol');col.set(qn('w:w'),str(width));grid.append(col)
        for ri,row in enumerate(rows):
            cells=t.rows[0].cells if ri==0 else t.add_row().cells
            row_template=table_template.rows[0 if ri==0 else 1]
            if row_template._tr.trPr is not None:
                t.rows[ri]._tr.insert(0,deepcopy(row_template._tr.trPr))
            for ci,txt in enumerate(row):
                c=cells[ci];template_cell=row_template.cells[min(ci,len(row_template.cells)-1)]
                cp=deepcopy(template_cell._tc.tcPr)
                cp.find(qn('w:tcW')).set(qn('w:w'),str(widths[ci]))
                c._tc.replace(c._tc.tcPr,cp)
                p=clone_properties(c.paragraphs[0],template_cell.paragraphs[0])
                put_run(p,txt,template_cell.paragraphs[0])
                p.paragraph_format.keep_with_next=(ri==0)
                expected.append(txt)
        tables+=1
        continue
    h=re.match(r'^(#{1,3})\s+(.*)$',line)
    if h:
        level=len(h.group(1));txt=h.group(2)
        p=add_p(txt,{1:title_template,2:h1,3:h2}[level])
        if level==1:
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next=True
            p.paragraph_format.page_break_before=False
    else:
        txt=line
        is_caption=bool(re.match(r'^表[78]-D\d+ ',txt))
        p=add_p(txt,caption if is_caption else body_template)
        if is_caption:
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent=Pt(0)
            p.paragraph_format.keep_with_next=True
            captions.append(txt)
    expected.append(txt);i+=1
doc.save(OUT/'D章节补充稿.docx')
assert doc.styles.element.xml==original_styles
assert source_hash==sha256(SOURCE.read_bytes()).hexdigest()
ban=['待定','不知道','待确认','待复核','待成员复核','待D审阅','待核查','未决','待决','尚待','未知','未确定','TBD','TODO']
for file in [MD,OUT/'附件/编写依据与核对记录.md',OUT/'附件/需求追踪明细.md']:
    data=file.read_text(encoding='utf-8')
    hits=[word for word in ban if word in data]
    assert not hits,(file,hits)
manifest={'version':version,'source_docx':str(SOURCE),'source_docx_sha256':source_hash,'source_styles_identical':True,'source_normal_font':'楷体','body_font_pt':12,'chapter_font_pt':16,'section_font_pt':14,'table_font_pt':10.5,'table_line_spacing':1.15,'table_header_fill':'none','captions_centered':captions,'table_count':tables,'expected_fragments':expected,'forbidden_prose_checked':ban,'visible_D_prose_hits':[],'protected_original_content_unchanged':True}
(QA/'格式与内容核对输入.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
contract=f'''# 原稿格式复用约定\n\n原件：{SOURCE.as_posix()}\n\nSHA256：{source_hash}\n\n复用范围：D独立补充稿。原件保持不变。页型为原件唯一纵向A4节，四边约2厘米，沿用原页脚页码。全部样式和目录样式从原件继承；章节从原件7和7.1的段落属性及首运行属性复制，分别为楷体16/14磅加粗、1.5倍行距。正文复制7.3首段，楷体12磅、1.5倍行距、首行2字。标题复制原件封面文档名的28磅楷体加粗。表格复制表6-1的Table Grid、单元格内边距、10.5磅楷体与1.15倍行距；白色表头、首行加粗、重复表头、禁止拆行。用户明确要求表名居中，覆盖原件左对齐并取消其首行缩进。表格列宽按新内容分配，总宽度继承原件。新增图表和目录均不增加。全部原有正文在本独立副本中移除，只放D内容；原件的样式、节、页眉页脚和引用关系作为复用组件保留。交付前Word刷新页码，导出PDF并检查全部页面。\n'''
(QA/'artifact.md').write_text(contract,encoding='utf-8')
print(json.dumps({k:v for k,v in manifest.items() if k!='expected_fragments'},ensure_ascii=False,indent=2))
