from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION
from pathlib import Path

out=Path(r'C:\Users\ncucc\Desktop\10-30\個人行程與任務管理系統_PRD.docx')
doc=Document()
sec=doc.sections[0]
sec.top_margin=Inches(.72);sec.bottom_margin=Inches(.72);sec.left_margin=Inches(.78);sec.right_margin=Inches(.78)
styles=doc.styles
for stylename in ['Normal','Title','Heading 1','Heading 2','Heading 3']:
    st=styles[stylename];st.font.name='Aptos';st._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei')
styles['Normal'].font.size=Pt(10.5);styles['Normal'].font.color.rgb=RGBColor(42,52,48)
styles['Normal'].paragraph_format.space_after=Pt(5)
styles['Title'].font.size=Pt(27);styles['Title'].font.bold=True;styles['Title'].font.color.rgb=RGBColor(0,0,0)
for sn,size in [('Heading 1',17),('Heading 2',13),('Heading 3',11)]:
    st=styles[sn];st.font.size=Pt(size);st.font.bold=True;st.font.color.rgb=RGBColor(0,0,0);st.paragraph_format.space_before=Pt(11);st.paragraph_format.space_after=Pt(5);st.paragraph_format.keep_with_next=True

def shade(cell,color):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd');shd.set(qn('w:fill'),color);tcPr.append(shd)
def borders(table,color='D9E1DC'):
    tblPr=table._tbl.tblPr;bd=OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        e=OxmlElement('w:'+edge);e.set(qn('w:val'),'single');e.set(qn('w:sz'),'5');e.set(qn('w:color'),color);bd.append(e)
    tblPr.append(bd)
def set_cell(cell,text,bold=False,color=None,size=9.2):
    cell.text='';p=cell.paragraphs[0];p.paragraph_format.space_after=Pt(1);p.paragraph_format.space_before=Pt(1)
    r=p.add_run(str(text));r.bold=bold;r.font.name='Aptos';r._element.rPr.rFonts.set(qn('w:eastAsia'),'Microsoft JhengHei');r.font.size=Pt(size)
    if color:r.font.color.rgb=RGBColor.from_string(color)
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

def table(headers,rows,widths=None):
    t=doc.add_table(rows=1,cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    if widths:
        for c,w in zip(t.columns,widths):c.width=Inches(w)
    for i,h in enumerate(headers):set_cell(t.rows[0].cells[i],h,True,'FFFFFF',9);shade(t.rows[0].cells[i],'426E5B')
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        if widths:
            for c,w in zip(cells,widths):c.width=Inches(w)
        for i,val in enumerate(row):
            set_cell(cells[i],val,size=9)
            if ri%2:shade(cells[i],'F3F6F3')
    borders(t)
    doc.add_paragraph().paragraph_format.space_after=Pt(2)
    return t

def bullet(text):
    p=doc.add_paragraph(style='List Bullet');p.paragraph_format.space_after=Pt(2);p.add_run(text);return p

def numbered(text):
    p=doc.add_paragraph(style='List Number');p.paragraph_format.space_after=Pt(2);p.add_run(text);return p

p=doc.add_paragraph(style='Title');p.add_run('個人行程與任務管理系統')
p=doc.add_paragraph();r=p.add_run('產品需求文件 PRD');r.bold=True;r.font.size=Pt(14);r.font.color.rgb=RGBColor(66,110,91)
p=doc.add_paragraph();p.paragraph_format.space_after=Pt(12);p.add_run('版本 1.0  ｜  日期 2026 年 10 月 30 日  ｜  狀態 初版需求').font.size=Pt(9)
p=doc.add_paragraph('本文件定義一套以手機為優先的個人行程與任務管理系統，協助單一使用者規劃每日行程、追蹤待辦期限、降低遺漏，並透過多科目學習目標掌握長期進度。此 PRD 根據目前討論與前端原型整理；通知服務與資料同步等尚待技術決策。')

doc.add_heading('1 產品概述',level=1)
doc.add_paragraph('產品暫名「日常 Daymark」，服務個人使用者，提供今日總覽、行程、任務、學習目標、提醒偏好與個人化介面。產品先以手機瀏覽器為主要使用情境，桌面端需保持可用。')
table(['項目','說明'],[
('產品目標','讓使用者快速知道今天要做什麼、何時到期，以及長期學習目標進展如何。'),
('主要使用者','單一個人使用者；不含多人協作、共享行事曆或權限管理。'),
('主要平台','手機優先的響應式 Web；桌面瀏覽器支援。'),
('提醒需求','希望 Gmail 與 LINE 在任務期限前一天提醒；需有提醒時間與管道設定。'),
('目前原型限制','目前為前端互動原型，資料保存於瀏覽器本機；Gmail／LINE 通知未實際串接。')],[1.35,5.55])

doc.add_heading('2 問題與成功指標',level=1)
doc.add_paragraph('使用者需要把行程、待辦和長期學習進度放在同一處管理。若今日資訊不易掃描、期限缺乏提醒、目標進度無法持續更新，使用者容易遺漏待辦或失去長期追蹤動力。')
table(['指標','初期目標／衡量方式'],[
('快速建立','新增一筆任務的核心流程不超過 1 分鐘；透過原型測試觀察完成率與時間。'),
('每日可見性','首頁清楚呈現當日行程、未完成待辦與目標摘要。'),
('完成追蹤','使用者可標記任務完成並看到當日待辦數更新。'),
('通知可用性','若正式版串接通知，使用者能選擇管道、提前時間，並確認連線狀態。'),
('目標維護','使用者可建立、修改、刪除不同科目的目標並追蹤進度與期限。')],[1.45,5.45])

doc.add_heading('3 範圍與版本規劃',level=1)
doc.add_heading('3.1 初版範圍',level=2)
for x in ['今日首頁：每日名言、日期、當日待辦摘要與行程列表。','任務管理：建立、編輯、完成任務；設定分類、優先級、說明、日期、時間、提醒與重複頻率。','行程管理：新增具日期／時間的行程並按時間排列。','學習目標：可管理多筆不同科目的目標；設定目標名稱、科目、備註、目前進度、目標總量與期限。','提醒偏好：Gmail、LINE 管道選擇與提醒提前時間設定；正式發送需後端與授權。','個人化：提供多組介面配色並保存使用者選擇。','資料持久化：原型階段使用瀏覽器本機儲存；正式版需定義帳號與跨裝置同步策略。']:
    bullet(x)
doc.add_heading('3.2 暫不納入',level=2)
for x in ['多人共用、團隊工作區與權限管理。','Gmail／LINE 實際通知服務的部署、授權流程與推播可靠性保證（另列技術整合項目）。','原生 iOS／Android App、離線多裝置同步、外部行事曆雙向同步。','複雜目標拆解、依賴關係、子任務與分析報表。']:
    bullet(x)

doc.add_heading('4 使用者情境與主要流程',level=1)
for x in ['早上查看首頁，掃描每日名言、今日行程與待辦。','快速新增任務，填寫名稱、分類、優先級、期限及提醒偏好。','編輯既有任務或標記完成，首頁待辦統計即時更新。','建立多個學習目標，例如日文、數學或設計；分別更新進度與期限。','查看提醒管道設定，選擇 Gmail／LINE 與提前通知時間。','切換介面配色，偏好在同一瀏覽器重新開啟後保留。']:
    numbered(x)

doc.add_heading('5 功能需求',level=1)
requirements=[
('FR-01','今日首頁','顯示當地日期、每日一句名言、今日行程、今日未完成任務數與今日任務。','必須'),
('FR-02','每日名言','每日按日期選出一句名言；同一天重新載入維持一致，日期變更時更新。','必須'),
('FR-03','新增任務','可輸入標題、分類、優先級、說明、日期、可選時間、提醒時間與重複頻率。','必須'),
('FR-04','編輯／完成任務','可從任務卡進入編輯；可標記完成或還原；變更後保存並更新列表。','必須'),
('FR-05','任務排序與提示','任務列表應能辨識一般／重要／緊急；期限與優先級提供視覺提示。','應有'),
('FR-06','行程管理','可新增具日期與時間的行程；當日行程依時間排序並顯示標題。','必須'),
('FR-07','多學習目標','可新增多筆不同科目目標，每筆包含科目、名稱、說明、目前進度、總量及期限。','必須'),
('FR-08','學習目標維護','可編輯或刪除目標；以目前進度／總量計算百分比，顯示距期限天數或逾期狀態。','必須'),
('FR-09','提醒設定','可選擇 Gmail、LINE、提醒提前時間；系統須在未完成串接時明確標示為示意。','必須'),
('FR-10','配色偏好','至少提供森林綠、海岸藍、薰衣草、暖陽橘；使用者選擇保存在本機。','應有'),
('FR-11','資料保存','原型用本機儲存保存任務、行程、目標及配色；正式版須防止意外資料遺失。','必須')]
table(['ID','功能','需求說明','優先級'],requirements,[.62,1.23,4.25,.8])

doc.add_heading('6 資料欄位建議',level=1)
table(['資料物件','主要欄位'],[
('任務','taskId、title、category、priority、description、dueDate、time（選填）、reminder、repeat、status、createdAt、updatedAt'),
('行程','eventId、title、date、startTime、endTime（待確認）、location／note（選填）'),
('學習目標','goalId、subject、title、description、currentProgress、targetTotal、unit（建議新增）、deadline、status'),
('提醒偏好','channel（Gmail／LINE）、enabled、leadTime、deliveryTime、connectionStatus'),
('使用者偏好','theme、locale、timeZone')],[1.35,5.55])
doc.add_paragraph('建議在學習目標新增「進度單位」（如堂、章、題、分鐘），以免不同科目使用不同量尺時產生理解歧義；此欄位尚待產品確認。')

doc.add_heading('7 介面與互動原則',level=1)
for x in ['手機優先、響應式版面；主要操作區與按鈕適合觸控。','首頁資訊層級以「今天」優先，避免長期目標與提醒遮蔽即時待辦。','新增與編輯表單提供明確標籤、必填提示、合理預設值及欄位驗證。','不同任務優先級不能只靠顏色區分，需搭配文字標籤。','刪除目標等不可逆操作須二次確認；完成狀態可還原。','配色切換需維持文字對比與狀態辨識，正式實作需進行無障礙檢查。']:
    bullet(x)

doc.add_heading('8 非功能需求',level=1)
table(['類別','需求'],[
('可用性','核心頁面在常見手機瀏覽器可操作，表單與導覽不依賴 hover。'),
('效能','首頁資料載入與主要互動應保持即時；目標數量增加時列表仍可捲動使用。'),
('可及性','使用語意化 HTML、鍵盤可操作、表單有可讀標籤與足夠色彩對比。'),
('隱私','任務及個人目標屬個人資訊；正式版須說明儲存位置、存取與刪除方式。'),
('相容性','支援近期主流手機與桌面瀏覽器；正式版需明確列出最低版本。'),
('可靠性','正式提醒需處理授權失效、退信、重試、時區與重複通知等情況。')],[1.2,5.7])

doc.add_heading('9 提醒與整合考量',level=1)
doc.add_paragraph('需求提出透過 Gmail 與 LINE 在期限前一天提醒。前端原型目前只能保存設定與呈現狀態，不會實際寄送。正式版需後端排程、使用者授權、憑證安全保存、時區處理、送達／失敗紀錄與取消通知規則。LINE 可用的官方訊息通道與 Gmail 發送方式，應在技術設計階段確認當期平台政策及帳號條件。')
table(['待決策','建議方向'],[
('提醒觸發時間','預設期限前一天上午 9:00，讓使用者可調整；避免在已完成任務上提醒。'),
('多管道策略','允許同時啟用 Gmail 與 LINE；需定義兩者同時送達或失敗時的行為。'),
('連線狀態','提供連線／未連線／授權過期狀態，避免介面顯示已設定但實際無法發送。'),
('無期限任務','不安排期限型提醒；可另行支援使用者指定提醒時間。')],[1.5,5.4])

doc.add_heading('10 驗收條件',level=1)
for x in ['新使用者可新增任務並保存必要欄位；重新載入原型後資料仍存在於同一瀏覽器。','任務可被編輯並保存；完成／還原操作會更新首頁統計。','可建立至少三個不同科目的學習目標，且各目標的進度、期限互不影響。','目標百分比以目前進度除以目標總量計算，超過總量時顯示不高於 100%。','任務與目標到期日顯示正確剩餘天數或逾期狀態。','配色選擇套用於頁面並於同一瀏覽器重新載入後保留。','每日名言在同一天保持一致，隔日可更新。','未完成郵件／LINE 後端整合前，介面不得宣稱提醒已實際送達。']:
    bullet(x)

doc.add_heading('11 風險、假設與依賴',level=1)
table(['類型','說明'],[
('假設','目前以單人、手機優先、繁體中文使用者為主；預設提醒提前一天。'),
('風險','純本機儲存可能因清除瀏覽器資料或更換裝置而遺失；需雲端同步或匯出備份策略。'),
('依賴','Gmail 與 LINE 實際發送依賴外部 API、帳號授權及後端服務。'),
('風險','不同目標進度可能使用不同單位，未指定單位會使百分比缺乏語意。'),
('風險','任務重複頻率設定需明確定義重複任務生成與完成行為。')],[1.1,5.8])

doc.add_heading('12 待確認事項',level=1)
for x in ['學習目標的進度單位是否需要自訂（堂、章、題、分鐘等）？','行程是否需要結束時間、地點及與外部行事曆同步？','提醒時間是否固定前一天上午 9:00，還是每筆任務可個別選擇時刻？','任務重複應採「自動建立下一筆」或同一筆循環顯示？','是否需要帳號登入與跨手機／電腦同步？','每日名言是否接受內建語錄清單，或需使用者自訂／收藏？']:
    bullet(x)

doc.add_heading('附錄 原型現況',level=1)
doc.add_paragraph('目前前端原型位於 index.html，包含每日名言、今日行程與待辦、多科目學習目標、任務詳細設定、配色切換及提醒設定示意。此文件描述產品需求與後續建議，不代表外部通知、帳號登入或雲端同步已完成。')

# Footer with document name and page number field
for section in doc.sections:
    fp=section.footer.paragraphs[0];fp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    rr=fp.add_run('個人行程與任務管理系統 PRD  |  ');rr.font.size=Pt(8);rr.font.color.rgb=RGBColor(110,120,114)
    fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');fp._p.append(fld)

doc.core_properties.title='個人行程與任務管理系統產品需求文件'
doc.core_properties.subject='PRD：手機優先的個人行程、任務、學習目標與提醒'
doc.core_properties.author='OpenAI'
doc.save(out)
print(out)
