import os
import re
import sys
import time
from playwright.sync_api import sync_playwright
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

WORKSPACE_DIR = "/Users/liuruichen/My Antigravity/taofuke_activity"
SCREENSHOTS_DIR = os.path.join(WORKSPACE_DIR, "screenshots")
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)

# 映射需求章节与HTML页面
PAGES_MAP = {
    "4.1.1": {
        "title": "H5小程序-活动列表",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/4.3-activity-list-v3.html",
        "is_h5": True,
        "filename": "h5_activity_list.png",
        "action": lambda page: (page.locator("#tabType").click(), page.wait_for_timeout(500))
    },
    "4.2.2.1": {
        "title": "H5小程序-问卷活动入口页",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/5.3.1-questionnaire-entrance.html",
        "is_h5": True,
        "filename": "h5_entrance.png"
    },
    "4.2.2.2": {
        "title": "H5小程序-问卷答题进行中",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/5.3.2-questionnaire-answering.html",
        "is_h5": True,
        "filename": "h5_answering.png"
    },
    "4.2.2.3": {
        "title": "H5小程序-答题成绩页",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/5.3.3-questionnaire-main.html?submit=true",
        "is_h5": True,
        "filename": "h5_results.png",
        "action": lambda page: (
            page.evaluate("""() => {
                localStorage.setItem('taofuke_questionnaire_result', JSON.stringify({
                    questions: [
                        {id: 1, type: 'single', score: 30, correctAnswers: [2]},
                        {id: 2, type: 'multiple', score: 40, correctAnswers: [0,1,3]},
                        {id: 3, type: 'single', score: 30, correctAnswers: [1]}
                    ],
                    userAnswers: {0: [2], 1: [0,1,3], 2: [1]},
                    timeSpentSeconds: 16
                }));
            }"""),
            page.reload(),
            page.wait_for_timeout(500)
        )
    },
    "4.2.2.4": {
        "title": "H5小程序-排行榜页",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/5.3.3-questionnaire-ranking.html",
        "is_h5": True,
        "filename": "h5_ranking.png"
    },
    "4.2.2.5": {
        "title": "H5小程序-活动详情页",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/5.3.4-questionnaire-detail.html",
        "is_h5": True,
        "filename": "h5_detail.png"
    },
    "4.2.2.6": {
        "title": "H5小程序-查看答案与解析页面",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/5.3.5-questionnaire-explanation.html",
        "is_h5": True,
        "filename": "h5_explanation.png",
        "action": lambda page: (
            page.evaluate("""() => {
                localStorage.setItem('taofuke_questionnaire_result', JSON.stringify({
                    questions: [
                        {id: 1, type: 'single', score: 30, correctAnswers: [2], text: '诈骗分子常常利用短信发送含有不明链接的信息。收到此类短信，正确的做法是？', options: ['A. 好奇点击链接查看详情', 'B. 拨打短信中的“客服电话”核实', 'C. 立即删除，不点击不理会', 'D. 转发给家人朋友看看'], explanation: '不明链接往往带有木马病毒或钓鱼网站，点击极易造成个人信息泄露或资金损失。最安全的做法是直接删除，切勿点击。'},
                        {id: 2, type: 'multiple', score: 40, correctAnswers: [0,1,3], text: '遇到所谓的“公检法”人员通过电话办案，并要求转移资金到“安全账户”时，哪些做法是正确的？', options: ['A. 坚决不转账', 'B. 挂断电话，拨打110核实', 'C. 按照对方要求下载屏幕共享软件', 'D. 向身边的亲友或同事求助咨询'], explanation: '公检法机关绝不会通过电话、网络直接办案，更不存在所谓的“安全账户”。屏幕共享软件会让诈骗分子看到您的验证码和密码，极其危险！'},
                        {id: 3, type: 'single', score: 30, correctAnswers: [1], text: '在二手交易平台上，如果卖家要求您脱离平台，通过微信私下转账交易，您应该怎么做？', options: ['A. 只要商品便宜，私下转也可以', 'B. 坚决拒绝，只在官方平台内完成交易', 'C. 先付一半定金试试看'], explanation: '在二手平台交易时，脱离平台的私下交易不受平台保护。一旦发生纠纷或诈骗，极难维权。必须始终通过官方通道进行付款和沟通。'}
                    ],
                    userAnswers: {0: [2], 1: [0,1,3], 2: [1]},
                    timeSpentSeconds: 16
                }));
            }"""),
            page.reload(),
            page.wait_for_timeout(500)
        )
    },
    "4.2.2.7": {
        "title": "H5小程序-历史答题记录页",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_h5/5.3.6-questionnaire-history.html",
        "is_h5": True,
        "filename": "h5_history.png"
    },
    "5.1.1": {
        "title": "运营商后台-活动列表",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_operate/6.1.1-activity-management-list.html",
        "is_h5": False,
        "filename": "operate_activity_list.png"
    },
    "5.1.2": {
        "title": "运营商后台-创建/编辑活动",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_operate/6.1.2-activity-create.html",
        "is_h5": False,
        "filename": "operate_activity_create.png",
        "action": lambda page: (
            page.locator(".type-card:has-text('问卷答题')").click(),
            page.locator("#actionBtn").click(),
            page.locator("input[name='enableRepeatAnswer'][value='yes']").click(),
            page.locator("input[name='enableTimeLimit'][value='yes']").click(),
            page.wait_for_timeout(500)
        )
    },
    "5.1.4": {
        "title": "运营商后台-查看活动详情",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_operate/6.1.4-activity-view-questionnaire.html",
        "is_h5": False,
        "filename": "operate_activity_view.png"
    },
    "5.2": {
        "title": "运营商后台-统计分析与排行榜弹窗",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_operate/6.2-activity-statistics.html",
        "is_h5": False,
        "filename": "operate_activity_statistics.png",
        "action": lambda page: (
            page.locator("text=查看排行榜").first.click(),
            page.wait_for_timeout(500)
        )
    },
    "6.1.1": {
        "title": "平台方后台-活动列表",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_operate/7.1.1-platform-activity-management-list.html",
        "is_h5": False,
        "filename": "platform_activity_list.png"
    },
    "6.1.2": {
        "title": "平台方后台-查看活动详情",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_operate/7.1.2-platform-activity-view-questionnaire.html",
        "is_h5": False,
        "filename": "platform_activity_view.png"
    },
    "6.2": {
        "title": "平台方后台-统计分析与排行榜弹窗",
        "url": f"file://{WORKSPACE_DIR}/taofuke_activity_operate/7.2-platform-activity-statistics.html",
        "is_h5": False,
        "filename": "platform_activity_statistics.png",
        "action": lambda page: (
            page.locator("text=查看排行榜").first.click(),
            page.wait_for_timeout(500)
        )
    }
}

def capture_screenshots():
    print("开始使用 Playwright 捕获页面截图...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # 运营商/平台端上下文 (Desktop)
        desktop_context = browser.new_context(viewport={"width": 1440, "height": 900})
        # H5小程序上下文 (Mobile 模拟)
        mobile_context = browser.new_context(viewport={"width": 450, "height": 900})
        
        for section, config in PAGES_MAP.items():
            print(f"正在捕获章节 [{section}] 的页面: {config['title']}...")
            ctx = mobile_context if config['is_h5'] else desktop_context
            page = ctx.new_page()
            page.goto(config['url'])
            page.wait_for_timeout(500)
            
            # 如果有自定义交互动作，先执行
            if "action" in config:
                try:
                    config["action"](page)
                except Exception as e:
                    print(f"执行交互失败 ({config['title']}): {e}")
            
            output_path = os.path.join(SCREENSHOTS_DIR, config['filename'])
            
            # H5页面我们截取模拟器框架元素以显得更逼真；如果没有，截取全屏
            if config['is_h5']:
                locator = page.locator(".iphone-mockup")
                if locator.count() > 0:
                    locator.screenshot(path=output_path)
                else:
                    page.screenshot(path=output_path)
            else:
                page.screenshot(path=output_path)
            
            page.close()
            print(f"截图已保存到: {output_path}")
            
        browser.close()

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def generate_word_document():
    md_path = "/Users/liuruichen/Downloads/淘福客工会营销活动组件需求二期（问卷答题）.md"
    docx_path = "/Users/liuruichen/Downloads/淘福客工会营销活动组件二期（问卷答题）页面设计需求.docx"
    
    print(f"正在读取需求文档并生成 Word 文档...")
    
    doc = docx.Document()
    
    # 页面边距设置
    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)
        
    # 全局字体设置
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Microsoft YaHei'
    font.size = Pt(10.5)
    font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    # 标题段落
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("淘福客工会营销活动组件二期（问卷答题）页面设计需求")
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    
    doc.add_paragraph("本文件作为问卷答题模块的前端界面设计与排版指导，包含了 H5 小程序端、运营商后台以及平台方后台的所有设计规格和交互要求，并附带了当前开发版本的原型截图作为对照。请设计师参考原型框架进行高保真 UI 设计。\n")
    
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    in_table = False
    table_lines = []
    current_section = None
    
    def self_insert_screenshot(doc, section_key):
        """在当前位置插入对应章节的截图"""
        config = PAGES_MAP[section_key]
        img_path = os.path.join(SCREENSHOTS_DIR, config['filename'])
        if os.path.exists(img_path):
            # 使用 doc.add_picture 插入图片（比 run.add_picture 更可靠）
            width = Inches(2.2) if config['is_h5'] else Inches(5.5)
            pic_paragraph = doc.add_picture(img_path, width=width)
            # add_picture 返回的段落中的 InlineShape, 其父段落需要居中
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # 添加图注
            p_caption = doc.add_paragraph()
            p_caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_caption.add_run(f"【图：{config['title']} 原型界面对照】")
            r_cap.font.size = Pt(9)
            r_cap.font.italic = True
            r_cap.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            p_caption.paragraph_format.space_after = Pt(8)
            print(f"  ✅ 已插入截图: {config['filename']} (章节 {section_key})")
        else:
            print(f"  ⚠️ 截图文件不存在: {img_path}")

    def process_table(lines_list):
        if not lines_list:
            return
        headers = [c.strip() for c in lines_list[0].split('|')[1:-1]]
        # 移除分割线行
        data_rows = []
        for line in lines_list[2:]:
            cells = [c.strip() for c in line.split('|')[1:-1]]
            if cells:
                data_rows.append(cells)
                
        table = doc.add_table(rows=len(data_rows) + 1, cols=len(headers))
        table.style = 'Table Grid'
        
        # 填充表头
        hdr_cells = table.rows[0].cells
        for idx, header in enumerate(headers):
            hdr_cells[idx].text = header
            set_cell_background(hdr_cells[idx], "F2F2F2")
            set_cell_margins(hdr_cells[idx])
            hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
            
        # 填充数据
        for r_idx, row_data in enumerate(data_rows):
            row_cells = table.rows[r_idx + 1].cells
            for c_idx, val in enumerate(row_data):
                if c_idx < len(row_cells):
                    row_cells[c_idx].text = val
                    set_cell_margins(row_cells[c_idx])
                    
        doc.add_paragraph() # 空白间距
        
    for line in lines:
        stripped = line.strip()
        
        # 处理表格
        if stripped.startswith('|'):
            in_table = True
            table_lines.append(stripped)
            continue
        elif in_table:
            process_table(table_lines)
            table_lines = []
            in_table = False
            
        if not stripped:
            continue
            
        # 标题行匹配 - 使用 re.sub 去除开头所有 # 符号
        if stripped.startswith('#'):
            # 提取标题级别和文本
            heading_match = re.match(r'^(#{1,6})\s+(.*)', stripped)
            if not heading_match:
                continue
            level = len(heading_match.group(1))
            header_text = heading_match.group(2).strip()
            
            if level == 1:
                p = doc.add_paragraph()
                run = p.add_run(header_text)
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2A, 0x95, 0xF5)
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(6)
            elif level == 2:
                p = doc.add_paragraph()
                run = p.add_run(header_text)
                run.font.size = Pt(12)
                run.font.bold = True
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(4)
                
                # 也检查 ## 级别标题的节号（如 5.2, 6.2）
                match = re.match(r"^(\d+(?:\.\d+)+)", header_text)
                if match:
                    current_section = match.group(1)
                    if current_section in PAGES_MAP:
                        self_insert_screenshot(doc, current_section)
                        current_section = None
            else:
                # ### 或 #### 级别
                # 检测节号（如 4.1.1, 4.2.2.1, 5.1.1, 5.1.4）
                match = re.match(r"^(\d+(?:\.\d+)+)", header_text)
                if match:
                    current_section = match.group(1)
                    
                p = doc.add_paragraph()
                run = p.add_run(header_text)
                run.font.size = Pt(11)
                run.font.bold = True
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(3)
                
                # 如果进入了特定需要配图的页面章节，直接将截图插入标题下方
                if current_section in PAGES_MAP:
                    self_insert_screenshot(doc, current_section)
                    current_section = None  # 使用完毕后置空，避免重复配图
        else:
            # 普通正文及列表行
            p = doc.add_paragraph()
            # 将 markdown 中的加粗 ** 转换为 Word 加粗
            parts = re.split(r'(\*\*.*?\*\*)', stripped)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    r_text = p.add_run(part[2:-2])
                    r_text.bold = True
                else:
                    p.add_run(part)
            p.paragraph_format.space_after = Pt(2)
            
    # 如果最后还有表格未处理
    if table_lines:
        process_table(table_lines)
        
    doc.save(docx_path)
    print(f"Word文档已成功生成并保存到: {docx_path}")

if __name__ == "__main__":
    try:
        capture_screenshots()
        generate_word_document()
        print("🎉 全部任务顺利完成！")
    except Exception as e:
        print(f"❌ 运行发生错误: {e}")
        sys.exit(1)
