import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import time
import pandas as pd
import os
import smtplib
from email.message import EmailMessage
import re

def normalize_dept_id(val) -> str:
    if pd.isna(val):
        return ""
    try:
        num = float(str(val).strip())
        if num.is_integer():
            return str(int(num))
        return str(num).strip()
    except Exception:
        s = str(val).strip()
        if s.endswith(".0"):
            s = s[:-2]
        if "." in s and s.replace(".", "", 1).isdigit():
            s = s.split(".", 1).strip()
        return s

def normalize_pubdate(pubdate_text: str) -> str:
    if not pubdate_text:
        return ""
    s = pubdate_text.strip()
    if len(s) >= 10 and s == "-" and s == "-":
        return s[:10]
    fmts = ["%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S %Z", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"]
    for fmt in fmts:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    return s

def send_email_report(file_path, target_date_str, has_data=True):
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_PASSWORD")
    receiver_email = os.environ.get("RECEIVER_EMAIL")

    if not sender_email or not sender_password or not receiver_email:
        print("⚠️ ตั้งค่า Email Secrets ไม่ครบ")
        return

    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    
    if has_data:
        msg['Subject'] = f'🚨 [A/B Test] ประกาศจัดซื้อใหม่ (แบบดึงด้วยรหัสหลัก 4 หลัก) - {target_date_str}'
        body_content = "แนบไฟล์รายงานการดึงข้อมูลแบบหว่านแห (เฉพาะรหัสหน่วยงานหลัก) รบกวนเปรียบเทียบกับระบบเดิมครับ\n\nขอบคุณครับ"
        msg.set_content(body_content)
        with open(file_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=os.path.basename(file_path))
    else:
        msg['Subject'] = f'ℹ️ [A/B Test] ไม่พบข้อมูล (แบบดึงด้วยรหัสหลัก) - {target_date_str}'
        msg.set_content("ตรวจสอบแล้ว ไม่พบประกาศจัดซื้อใหม่ที่ตรงคีย์เวิร์ดครับ")

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        print("✅ ส่งอีเมลสำเร็จ!")
    except Exception as e:
        print(f"❌ ส่งอีเมลล้มเหลว: {e}")

def fetch_main_dept_egp():
    input_file = "Gov_Main_List.xlsx"
    if not os.path.exists(input_file):
        print(f"❌ ไม่พบไฟล์: {input_file}")
        return

    try:
        df_gold = pd.read_excel(input_file, dtype=str)
    except Exception:
        return

    ofm_keywords = [
        "กระดาษ", "หมึกพิมพ์", "โทนเนอร์", "วัสดุสำนักงาน", "อุปกรณ์สำนักงาน", "แฟ้ม", "คอมพิวเตอร์", "โน้ตบุ๊ก", "ปริ้นเตอร์",
        "เฟอร์นิเจอร์", "โต๊ะ", "เก้าอี้", "ตู้", "เครื่องปรับอากาศ", "แอร์", "กระดาษชำระ", "ถุงขยะ"
    ] 
    
    now = datetime.now()
    if now.hour < 12:
        target_date_str = (now - timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        target_date_str = now.strftime("%Y-%m-%d")

    all_projects = []
    total_rows = len(df_gold)
    print(f"🚀 เริ่มทำงานแบบหว่านแห รหัสหลัก {total_rows} หน่วยงาน")

    for index, row in df_gold.iterrows():
        # ⚠️ แก้ไขจุดนี้ให้ตรงกับหัวคอลัมน์ใน Excel ของคุณ (B และ C)
        main_dept = str(row.get("ชื่อหน่วยงาน", "")).strip()
        main_dept_id = normalize_dept_id(row.get("รหัสหน่วยงาน", None))
        
        # เติม 0 ให้ครบ 4 หลัก
        if main_dept_id and main_dept_id.isdigit() and len(main_dept_id) < 4:
            main_dept_id = main_dept_id.zfill(4)

        if not main_dept_id or main_dept_id.lower() == "nan":
            continue

        # ยิงแค่รหัสหลักเท่านั้น ไม่มี deptsubId เจือปน
        url = f"http://process3.gprocurement.go.th/EPROCRssFeedWeb/egpannouncerss.xml?methodId=16&anounceType=D0&deptId={main_dept_id}"
        
        print(f"📡 [{index + 1}/{total_rows}] ค้นหา: {main_dept} (รหัส: {main_dept_id})")
        
        try:
            response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
            xml_text = response.content.decode("cp874", errors="replace").replace("Windows-874", "utf-8")
            root = ET.fromstring(xml_text)
            
            for item in root.findall("./channel/item"):
                pubDate = normalize_pubdate(item.findtext("pubDate", default="").strip())
                if pubDate != target_date_str: continue
                    
                title = item.findtext("title", default="").strip()
                link = item.findtext("link", default="").strip()
                desc = item.findtext("description", default="").strip()
                
                budget = "ไม่ระบุ"
                budget_match = re.search(r'(งบประมาณ|ราคากลาง|วงเงิน)[^\d]*([\d,]+\.?\d*)', desc)
                if budget_match: budget = budget_match.group(2) + " บาท"

                matched_kw = [kw for kw in ofm_keywords if kw and kw in title]
                if not matched_kw: continue
                    
                all_projects.append({
                    "วันที่ประกาศ": pubDate,
                    "ชื่อหน่วยงานหลัก": main_dept,
                    "รหัสหน่วยงานหลัก": main_dept_id,
                    "ชื่อโครงการจัดซื้อ": title,
                    "งบประมาณ (ประเมิน)": budget,
                    "ความน่าสนใจ": f"★ ({', '.join(matched_kw)})",
                    "ลิงก์เอกสาร (TOR)": link
                })
        except Exception:
            pass
        time.sleep(1.5) # หน่วงเวลา 1.5 วิ ป้องกันโดนแบน

    if all_projects:
        df_result = pd.DataFrame(all_projects)
        output_file = f"eGP_Main_Result_{target_date_str.replace('-','')}.xlsx"
        df_result.to_excel(output_file, index=False)
        send_email_report(output_file, target_date_str, has_data=True)
    else:
        send_email_report(None, target_date_str, has_data=False)

if __name__ == "__main__":
    fetch_main_dept_egp()
