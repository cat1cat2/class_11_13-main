from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK
from docx.enum.table import WD_ROW_HEIGHT_RULE
from datetime import date
from pathlib import Path

OUT = Path(r'C:\Users\ncucc\Desktop\1030\個人學習任務管理系統_PRD.docx')
doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(.7); sec.bottom_margin = Inches(.65); sec.left_margin = Inches(.78); sec.right_margin = Inches(.78)

# restrained product-spec palette
GREEN='315D48'; PALE='EAF0E8'; GOLD='C49A3A'; GRAY='667064'; LIGHT='F3F5F1'; WHITE='FFFFFF'; BLACK='20251F'; RED='B96848'
styles=doc.styles
normal=styles['Normal']; normal.font.name='Microsoft JhengHei'; normal.font.size=Pt(9.2); normal.font.color.rgb=RGBColor.from_string(BLACK); normal.paragraph_format.space_after=Pt(5); normal.paragraph_format.line_spacing=1.12
for nm,size in [('Title',28),('Heading 1',17),('Heading 2',12),('Heading 3',10)]:
    s=styles[nm]; s.font.name='Microsoft JhengHei'; s.font.size=Pt(size); s.font.bold=True; s.font.color.rgb=RGBColor.from_string(BLACK); s.paragraph_format.keep_with_next=True
styles['Title'].paragraph_format.space_after=Pt(5)
styles['Heading 1'].paragraph_format.space_before=Pt(14);styles['Heading 1'].paragraph_format.space_after=Pt(6)
styles['Heading 2'].paragraph_format.space_before=Pt(8);styles['Heading 2'].paragraph_format.space_after=Pt(3)

def set_cell_shading(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)
def set_cell_margins(cell, top=80, start=100, bottom=80, end=100):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr(); mar=tcPr.first_child_found_in('w:tcMar')
    if mar is None: mar=OxmlElement('w:tcMar'); tcPr.append(mar)
    for m,v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        el=mar.find(qn('w:'+m))
        if el is None: el=OxmlElement('w:'+m); mar.append(el)
        el.set(qn('w:w'),str(v)); el.set(qn('w:type'),'dxa')
def repeat_header(row):
    trPr=row._tr.get_or_add_trPr(); el=OxmlElement('w:tblHeader');el.set(qn('w:val'),'true');trPr.append(el)
def set_run(run, bold=None, color=None, size=None):
    run.font.name='Microsoft JhengHei';run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
    if bold is not None: run.bold=bold
    if color:run.font.color.rgb=RGBColor.from_string(color)
    if size:run.font.size=Pt(size)
def table(headers, rows, widths=None, font_size=8.3):
    t=doc.add_table(rows=1,cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    if widths:
        for c,w in zip(t.columns,widths): c.width=Inches(w)
    t.style='Table Grid'; repeat_header(t.rows[0])
    for i,h in enumerate(headers):
        c=t.rows[0].cells[i];set_cell_shading(c,GREEN);set_cell_margins(c);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=c.paragraphs[0];p.paragraph_format.space_after=Pt(0);r=p.add_run(h);set_run(r,True,WHITE,font_size)
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        if widths:
            for c,w in zip(cells,widths):c.width=Inches(w)
        for i,val in enumerate(row):
            c=cells[i];c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER;set_cell_margins(c)
            if ri%2==1:set_cell_shading(c,'F7F8F5')
            p=c.paragraphs[0];p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.05
            r=p.add_run(str(val));set_run(r,False,BLACK,font_size)
    doc.add_paragraph().paragraph_format.space_after=Pt(0)
    return t
def para(text='', bold_lead=None, color=None):
    p=doc.add_paragraph()
    if bold_lead and text.startswith(bold_lead):
        r=p.add_run(bold_lead);set_run(r,True,color or BLACK)
        r=p.add_run(text[len(bold_lead):]);set_run(r,color=color)
    else:
        r=p.add_run(text);set_run(r,color=color)
    return p
def bullet(text, level=0):
    p=doc.add_paragraph(style='List Bullet' if level==0 else 'List Bullet 2');p.paragraph_format.space_after=Pt(2);r=p.add_run(text);set_run(r,size=9);return p
def number(text):
    p=doc.add_paragraph(style='List Number');p.paragraph_format.space_after=Pt(2);r=p.add_run(text);set_run(r,size=9);return p

# Header/footer
header=sec.header.paragraphs[0];header.alignment=WD_ALIGN_PARAGRAPH.RIGHT
r=header.add_run('產品需求文件  /  PERSONAL LEARNING');set_run(r,True,GRAY,8)
footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=footer.add_run('個人學習任務管理系統　·　產品需求文件　·　2026');set_run(r,False,GRAY,8)

# Cover / executive summary
p=doc.add_paragraph();p.paragraph_format.space_before=Pt(14);p.paragraph_format.space_after=Pt(5)
r=p.add_run('產品需求文件');set_run(r,True,GREEN,10)
p=doc.add_paragraph(style='Title');p.add_run('個人學習任務管理系統')
p=doc.add_paragraph();p.paragraph_format.space_after=Pt(15);r=p.add_run('Product Requirements Document  |  桌面優先 MVP');set_run(r,False,GRAY,11)
# Summary callout
call=doc.add_table(rows=1,cols=1);call.alignment=WD_TABLE_ALIGNMENT.CENTER;call.autofit=False;call.columns[0].width=Inches(6.9)
c=call.cell(0,0);set_cell_shading(c,PALE);set_cell_margins(c,150,190,150,190)
p=c.paragraphs[0];p.paragraph_format.space_after=Pt(3);r=p.add_run('產品摘要');set_run(r,True,GREEN,10)
p=c.add_paragraph();p.paragraph_format.space_after=Pt(0);r=p.add_run('為難以判斷任務先後的個人學習者，提供清楚的今日待辦清單。系統依重要程度與截止日提出優先建議，使用者可調整星等、手動排序，並以完成狀態追蹤每日進度。');set_run(r,False,BLACK,9.2)
doc.add_paragraph().paragraph_format.space_after=Pt(0)
table(['文件欄位','內容'],[
('文件版本','1.0｜初步需求版'),('狀態','MVP 需求基線；待產品／技術評估'),('主要平台','桌面網頁，寬螢幕優先；具基本響應式能力'),('產品範圍','個人學習任務管理，不含團隊協作'),('需求依據','需求訪談與已完成的互動式 HTML 原型'),('文件日期',date.today().strftime('%Y 年 %-m 月 %-d 日') if False else '2026 年 10 月 30 日')
],[1.45,5.45],8.5)

doc.add_heading('1. 產品背景與問題定義',level=1)
para('個人學習者常同時面對多項待辦，若缺少明確的優先順序，容易在任務間反覆切換，或只處理看起來緊急的事項。本產品以每日待辦為主要工作入口，將「重要程度」作為排序核心，並以截止日期補足時效性判斷。')
para('產品原則：系統負責提出可理解的推薦，使用者保有最終決定權。任務新增流程保持精簡，預計時間與備註為選填。')

doc.add_heading('2. 產品目標與成功指標',level=1)
table(['目標','衡量方式（MVP 建議）','初始目標'],[
('快速掌握今天要做什麼','首頁今日待辦區可見；量測首次進入至辨識第一項建議任務的時間','中位數 ≤ 10 秒'),
('協助建立合理優先順序','推薦排序使用率／接受推薦後完成比例','建立基線後設定目標'),
('降低管理摩擦','新增一項任務所需時間及表單放棄率','新增操作 ≤ 30 秒'),
('維持任務追蹤','每週至少完成一項任務的活躍使用者比例','上線後依基線訂目標')
],[1.45,3.85,1.6],8.1)
para('注意：以上為建議驗收方向，尚未定義正式分析工具、資料保留政策與量化營運目標。',color=GRAY)

doc.add_heading('3. 目標使用者與使用情境',level=1)
para('主要使用者：以電腦管理個人學習計畫、希望改善任務優先順序判斷的個人學習者。')
for s in ['開始一天時，快速確認今日待辦與系統建議先做的任務。','建立學習任務，標記 1–5 星重要程度與截止日，視需要補上預計時間或備註。','任務完成後標記完成；若推薦不符合當下狀況，可調整星等或拖曳任務順序。','查看全部待辦或已完成項目，回顧本週進展。'] : bullet(s)

doc.add_heading('4. 範圍界定',level=1)
table(['MVP 納入','MVP 不納入（後續評估）'],[
('今日待辦首頁、全部任務、已完成任務檢視','多使用者、團隊／共享清單與權限管理'),
('任務新增、編輯、刪除、完成／恢復','重複任務、子任務、依賴關係及專案階層'),
('重要度 1–5 星、截止日、選填預計時間與備註','完整週／月日曆與精確時段排程'),
('推薦排序、截止日排序、重要度排序、手動拖曳排序','通知推播、外部日曆同步與跨裝置帳號同步'),
('本機瀏覽器儲存；提供示範資料','雲端備份、登入、資料匯出／匯入與同步'),
('每日完成進度與簡易本週完成統計','深度學習分析、習慣養成及 AI 自動排程')
],[3.45,3.45],8.1)

doc.add_heading('5. 功能需求',level=1)
features=[
('FR-01','今日待辦首頁','首次進入預設顯示今日未完成任務、完成進度摘要與優先推薦。','Must'),
('FR-02','新增任務','任務名稱必填；重要程度預設 3 星；截止日預設今天；預計時間及備註選填。','Must'),
('FR-03','編輯任務','可修改名稱、重要度、截止日、預計時間及備註。','Must'),
('FR-04','完成狀態','可標記完成及恢復待辦；完成項目可在已完成檢視中查看。','Must'),
('FR-05','刪除任務','使用者可刪除任務；刪除前需二次確認。','Must'),
('FR-06','推薦排序','預設以截止狀況與重要度計算建議順序，並清楚標示推薦任務。','Must'),
('FR-07','手動排序','可拖曳調整任務順序；手動順序應與系統推薦模式區隔，且不覆寫星等。','Must'),
('FR-08','排序切換','支援推薦、手動、截止日期及重要程度排序。','Should'),
('FR-09','逾期辨識','對早於今日且未完成的任務顯示逾期狀態；不可錯標當日截止項目。','Must'),
('FR-10','任務檢視切換','可切換今日待辦、全部未完成任務與已完成任務。','Must'),
('FR-11','進度摘要','顯示今日完成數／今日任務數、本週完成數及基本週視圖任務提示。','Should'),
('FR-12','本機持久化','在同一瀏覽器保存任務與狀態；重新載入頁面後資料仍在。','Must'),
('FR-13','空狀態','無任務、無已完成項目時，提供明確說明與新增任務入口。','Should')]
table(['ID','功能','需求描述','優先級'],features,[.58,1.12,4.45,.75],7.65)

doc.add_heading('6. 推薦與手動排序規則',level=1)
para('推薦排序的目的為提供可解釋的起始順序，不宣稱代表唯一正確的學習策略。原型規則如下，正式產品上線前需由產品負責人確認：')
for s in ['先辨識逾期任務，再辨識今日到期任務；兩者優先於未來到期或未設定日期的任務。','在相同截止狀況下，重要度較高者優先；仍相同時，以較早截止日期優先。','對逾期任務，建議優先提醒，但如有大量逾期項目，排序仍以星等協助取捨。','使用者調整星等後，推薦結果重新計算；切換至「我的排序」或拖曳後，手動順序獨立保留。','排序模式與今日篩選是兩個不同維度：今日檢視顯示今日到期的未完成任務；全部任務檢視呈現所有未完成項目。'] : bullet(s)

table(['排序因素','定義','優先次序'],[
('截止狀況','逾期、今天到期、未來到期、未設定','第一層'),
('重要程度','使用者指定 1–5 星','同層級內第二層'),
('截止日期','日期較早者先','同星等時第三層'),
('手動位置','使用者拖曳後的位置','僅在手動排序模式採用')
],[1.35,3.95,1.6],8.2)

doc.add_heading('7. 任務資料欄位與驗證',level=1)
table(['欄位','必要性','規則／驗證','備註'],[
('任務 ID','系統產生','唯一值','用於辨識任務'),
('任務名稱','必填','去除前後空白；1–90 字','列表主要標題'),
('重要程度','必填（預設 3）','整數 1–5','1 顆星最低、5 顆星最高'),
('截止日期','目前原型提供日期欄位，需求待確認','合法日期；可未設定或預設今天','逾期判斷依使用者本地日期'),
('預計時間','選填','從支援清單選取；可清除','不做日曆時段配置'),
('備註','選填','最多 120 字（原型限制）','純文字'),
('完成狀態','系統維護','待辦／已完成','完成時間可供本週統計'),
('手動排序值','系統維護','同一清單內可比較','需定義儲存及範圍')
],[1.15,1.45,2.7,1.6],7.7)

doc.add_heading('8. 主要使用流程',level=1)
for s in ['進入系統 → 預設落在今日待辦 → 查看優先推薦與當日進度。','選擇「新增任務」→ 輸入名稱 → 選重要星等與截止日 → 可選填時間及備註 → 儲存。','在待辦列表查看建議順序 → 直接調整星等或切換手動排序並拖曳。','完成任務 → 進度與已完成清單更新 → 可於已完成頁面回顧。'] : number(s)

doc.add_heading('9. 介面與體驗原則',level=1)
for s in ['桌面優先：以側邊導覽、主任務清單與右側推薦／週視圖構成寬螢幕主畫面。','一眼辨識：今日待辦為預設入口；優先推薦需與一般清單項目區隔。','低輸入負擔：僅任務名稱為強制輸入；非必要欄位清楚標示選填。','使用者掌控：推薦為建議，不可阻止使用者修改重要度、排序或任務日期。','狀態清楚：逾期、今天到期、已完成、空清單皆須有可辨識文字與視覺狀態，不只依賴顏色。','鍵盤可用：表單欄位有標籤，對話框可關閉，按鈕需具可讀名稱；拖曳需評估鍵盤替代操作。'] : bullet(s)

doc.add_heading('10. 非功能需求',level=1)
table(['面向','MVP 需求'],[
('相容性','支援現代桌面版 Chrome、Edge、Firefox、Safari；提供基本窄螢幕響應式呈現。'),
('效能','初次載入與一般任務操作應即時回應；目標任務量至少 500 筆仍可流暢瀏覽（建議驗收）。'),
('資料保存','目前為瀏覽器 localStorage；清除瀏覽資料或換瀏覽器會遺失資料，需在產品內明確告知。'),
('隱私與安全','MVP 不需帳號；任務內容留在使用者瀏覽器，不應送出至伺服器。若未來加上分析或雲端同步，需重新評估同意與資料保護。'),
('可及性','表單標籤、鍵盤焦點、狀態文字與足夠對比；可朝 WCAG 2.1 AA 檢查。'),
('時區／日期','依使用者裝置本地日期判定今日與逾期；需測試跨日、閏日、時區與夏令時間情境。')
],[1.25,5.65],8)

doc.add_heading('11. 驗收條件',level=1)
checks=[
'首次開啟顯示今日待辦；當日任務依規則提供推薦順位。',
'可新增一筆任務，並以 1–5 星重要度儲存；名稱未填時不可送出。',
'預計時間與備註可留空；已填值可在列表或編輯表單查看。',
'可改星等、截止日期與任務文字；重新載入後資料仍保留於同一瀏覽器。',
'標記完成後任務從未完成清單移除，並可在已完成清單找到；可恢復待辦。',
'手動拖曳調整後，切換手動排序可看到更新順序；切換推薦排序依規則重新排列。',
'截止日期早於本地今日顯示逾期；今天截止項目顯示今天到期而非逾期。',
'刪除任務前出現確認；取消確認時資料不變。',
'清單為空時顯示友善空狀態，並可從介面新增任務。']
for s in checks: bullet('☐  '+s)

doc.add_heading('12. 風險、依賴與待決事項',level=1)
table(['類型','議題','建議處理'],[
('產品決策','截止日期是否必填？現有訪談確認要用截止日，但未明確確認必填性。','MVP 建議允許未設定日期；新增表單預設今天並可清除，請產品負責人確認。'),
('產品決策','推薦排序中逾期與高重要度的權重關係。','確認逾期任務是否一律置頂，並以代表性任務案例走查排序。'),
('產品決策','「今日待辦」是否包含今天建立、但沒有截止日的任務。','定義今日任務／排定方式；本原型以截止日等於今天篩選。'),
('技術風險','localStorage 僅限單一瀏覽器，沒有跨裝置同步與備份。','MVP 加醒目說明；若有多裝置需求，規劃登入、同步和備份。'),
('原型落差','刪除使用瀏覽器原生確認；本週統計由完成時間計算。','正式版本改為一致的產品對話框，補上測試與資料規則。'),
('無障礙風險','目前拖曳排序缺少明確鍵盤操作替代。','增加上移／下移操作或其他可及性調整。'),
('指標依賴','目前未選定事件分析與使用者回饋管道。','若需追蹤成功指標，先界定隱私政策、事件清單和同意方式。')
],[1.05,2.95,2.9],7.5)

doc.add_heading('附錄 A　原型對照',level=1)
para('互動式原型檔案：index.html。原型已涵蓋主要桌面首頁、今日／全部／已完成切換、任務 CRUD、星等調整、截止日提示、推薦與手動排序、進度摘要及瀏覽器本機儲存。原型主要用於需求討論，不等同正式產品規格；其預設推薦權重、截止日預設值、示範資料與資料保存方式均須在開發前定案。')

# Avoid splitting rows and keep title hierarchy tidy
for t in doc.tables:
    for row in t.rows:
        trPr=row._tr.get_or_add_trPr();el=OxmlElement('w:cantSplit');trPr.append(el)
# Core metadata
props=doc.core_properties;props.title='個人學習任務管理系統產品需求文件';props.subject='桌面優先 MVP 產品需求文件';props.author='OpenAI';props.keywords='PRD, 學習任務, 個人管理'
doc.save(OUT)
print(OUT)
