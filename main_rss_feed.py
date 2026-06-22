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
        msg['Subject'] = f'🚨 ประกาศจัดซื้อใหม่ - {target_date_str}'
        body_content = "แนบไฟล์รายงานการดึงข้อมูล e-GP\n\nขอบคุณครับ"
        msg.set_content(body_content)
        with open(file_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=os.path.basename(file_path))
    else:
        msg['Subject'] = f'ℹ️ ไม่พบข้อมูล - {target_date_str}'
        msg.set_content("ตรวจสอบแล้ว ไม่พบประกาศจัดซื้อใหม่ใน e-GP ที่ตรงคีย์เวิร์ดครับ")

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        print("✅ ส่งอีเมลรายงานประจำวันสำเร็จ!")
    except Exception as e:
        print(f"❌ ส่งอีเมลล้มเหลว: {e}")

# ==========================================
# ส่วนที่เพิ่มใหม่: ระบบส่งอีเมลสรุปรายเดือน
# ==========================================
def send_monthly_summary_report(file_path, month_display):
    sender_email = os.environ.get("SENDER_EMAIL")
    sender_password = os.environ.get("SENDER_PASSWORD")
    receiver_email = os.environ.get("RECEIVER_EMAIL")

    if not sender_email or not sender_password or not receiver_email:
        return

    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f'📊 สรุปประกาศจัดซื้อสะสมเดือน {month_display}'
    
    body_content = (
        f"เรียน ทีมงาน,\n\n"
        f"รายงานสรุปข้อมูลประกาศจัดซื้อจัดจ้างสะสม ประจำเดือน {month_display}\n\n"
        f"ขอบคุณครับ"
    )
    msg.set_content(body_content)
    
    if file_path and os.path.exists(file_path):
        with open(file_path, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename=os.path.basename(file_path))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(sender_email, sender_password)
            smtp.send_message(msg)
        print("✅ ส่งอีเมลสรุปรายเดือนสำเร็จ!")
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการส่งอีเมลรายเดือน: {e}")

def check_and_process_monthly_report():
    today = datetime.now()
    # ตรวจสอบว่าวันนี้คือวันที่ 1 หรือไม่
    if today.day == 1:
        first_of_this_month = today.replace(day=1)
        last_day_of_prev_month = first_of_this_month - timedelta(days=1)
        prev_month_str = last_day_of_prev_month.strftime('%Y%m')
        prev_month_display = last_day_of_prev_month.strftime('%m/%Y')
        
        monthly_file = os.path.join("Backup", f"eGP_Monthly_Main_{prev_month_str}.xlsx")
        if os.path.exists(monthly_file):
            send_monthly_summary_report(monthly_file, prev_month_display)
# ==========================================

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
        # หมวดกระดาษและเครื่องเขียน (ใช้คำมาตรฐานราชการ)
        "กระดาษ", "กระดาษถ่ายเอกสาร", "กระดาษสี", "กระดาษต่อเนื่อง", "กระดาษโพสต์อิท", "กระดาษโน้ต", "กระดาษสติกเกอร์", "กระดาษโฟโต้", 
        "เอ 4", "เอสี่", "ขนาด A4", "แฟ้ม", "แฟ้มเอกสาร", "แฟ้มสันกว้าง", "แฟ้มโชว์เอกสาร", "ซองจดหมาย", "ซองเอกสาร", "ซองน้ำตาล", 
        "เครื่องเขียน", "ปากกา", "ปากกาเน้นข้อความ", "น้ำยาลบคำผิด", "สมุด", "สมุดบันทึก", "กระดาน", "ไวท์บอร์ด", "กระดานดำ", "บอร์ดนิทรรศการ", 
        "ที่เย็บกระดาษ", "ลวดเย็บกระดาษ", "คลิปหนีบกระดาษ", "เทปกาว", "กาว", "แท่นตัดกระดาษ", "ตรายาง", "ป้ายชื่อ",

        # หมวดอุปกรณ์และเครื่องใช้สำนักงาน (คำหมวดวัสดุ/ครุภัณฑ์สำนักงาน)
        "วัสดุสำนักงาน", "อุปกรณ์สำนักงาน", "ครุภัณฑ์สำนักงาน", "เครื่องทำลายเอกสาร", "เครื่องเคลือบบัตร", "เครื่องคิดเลข", "เครื่องเข้าเล่ม", 
        "เครื่องเจาะกระดาษ", "เครื่องตอกบัตร", "เครื่องสแกนลายนิ้วมือ", "เครื่องนับธนบัตร", "ตู้จดหมาย",

        # หมวดหมึกพิมพ์และวัสดุสิ้นเปลือง
        "หมึกพิมพ์", "โทนเนอร์", "ตลับหมึก", "หมึกเติม", "ดรัม", "ผ้าหมึก", "ตลับผงหมึก", "หมึกเครื่องพิมพ์",

        # หมวดไอที คอมพิวเตอร์ และระบบเครือข่าย (อิงเกณฑ์ราคากลาง ICT)
        "คอมพิวเตอร์", "คอมพิวเตอร์ส่วนบุคคล", "คอมพิวเตอร์พกพา", "โน้ตบุ๊ก", "โน๊ตบุ๊ค", "แท็บเล็ต", 
        "เครื่องพิมพ์", "ปริ้นเตอร์", "เครื่องสแกนเนอร์", "เครื่องฉายภาพ", "โปรเจคเตอร์", "จอภาพ", "จอแสดงผล", "จอมอนิเตอร์", 
        "เครื่องสำรองไฟฟ้า", "เครื่องสำรองไฟ", "เมาส์", "คีย์บอร์ด", "แป้นพิมพ์", "ฮาร์ดดิสก์", "หน่วยจัดเก็บข้อมูล", "อุปกรณ์บันทึกข้อมูล", 
        "แฟลชไดร์ฟ", "ยูเอสบี", "หน่วยความจำ", "แผงวงจรหลัก", 
        "อุปกรณ์เครือข่าย", "อุปกรณ์จัดเส้นทาง", "เราเตอร์", "อุปกรณ์กระจายสัญญาณ", "สวิตช์", "จุดกระจายสัญญาณ", "สายแลน", "สายสัญญาณ", 
        "เครื่องแม่ข่าย", "เซิร์ฟเวอร์", "สมาร์ททีวี", "โทรทัศน์", "ปลั๊กพ่วง", "หูฟัง", "ไมโครโฟน", "กล้องเว็บแคม", "ระบบประชุมทางไกล", 
        "ซอฟต์แวร์", "ระบบปฏิบัติการ", "โปรแกรมสำนักงาน", "โปรแกรมป้องกันไวรัส", "เครื่องรูดบัตร", "เครื่องอ่านบาร์โค้ด", "ลิ้นชักเก็บเงิน", "เครื่องบันทึกเงินสด", "ครุภัณฑ์คอมพิวเตอร์",

        # หมวดเฟอร์นิเจอร์สำนักงาน
        "เฟอร์นิเจอร์", "ครุภัณฑ์", "โต๊ะ", "โต๊ะทำงาน", "โต๊ะสำนักงาน", "โต๊ะคอมพิวเตอร์", "โต๊ะประชุม", "โต๊ะอเนกประสงค์", "โต๊ะพับ", "โต๊ะผู้บริหาร", 
        "เก้าอี้", "เก้าอี้สำนักงาน", "เก้าอี้ผู้บริหาร", "เก้าอี้เพื่อสุขภาพ", "เก้าอี้พับ", "เก้าอี้จัดเลี้ยง", "โซฟา", 
        "ตู้เก็บเอกสาร", "ตู้เหล็ก", "ตู้ล็อคเกอร์", "ตู้เซฟ", "ตู้นิรภัย", "บานเลื่อน", "ชั้นวางของ", "ชั้นเหล็ก", "พาร์ทิชั่น", "ฉากกั้น", "รถเข็น",

        # หมวดอุปกรณ์ทำความสะอาดและงานแม่บ้าน
        "อุปกรณ์ทำความสะอาด", "น้ำยาทำความสะอาด", "น้ำยาล้างห้องน้ำ", "น้ำยาล้างจาน", "น้ำยาถูพื้น", 
        "ถุงพลาสติก", "ถุงขยะ", "ถังขยะ", "ไม้กวาด", "ไม้ถูพื้น", "ม็อบ", "เครื่องขัดพื้น", "เครื่องดูดฝุ่น", "เครื่องฉีดน้ำ", 
        "กระดาษชำระ", "กระดาษเช็ดมือ", "กระดาษเช็ดหน้า", "สบู่เหลว", "เจลล้างมือ", "แอลกอฮอล์", "หน้ากากอนามัย", "ถุงมือยาง", "ถุงมือพลาสติก", "กล่องพลาสติก",

        # หมวดเครื่องใช้ไฟฟ้าและส่วนกลาง
        "เครื่องใช้ไฟฟ้า", "เครื่องปรับอากาศ", "พัดลม", "พัดลมระบายอากาศ", "พัดลมอุตสาหกรรม", "เครื่องทำน้ำอุ่น", "เครื่องฟอกอากาศ", 
        "เครื่องกดน้ำ", "เครื่องทำน้ำเย็น", "ตู้เย็น", "ไมโครเวฟ", "กระติกน้ำร้อน", "กาต้มน้ำ", "น้ำดื่ม", "เครื่องชงกาแฟ", 
        "แก้วน้ำ", "แก้วกระดาษ", "กล่องข้าว", "จานชาม", "ช้อนส้อม",

        # หมวดอุปกรณ์ความปลอดภัยและเครื่องมือช่าง
        "กล้องวงจรปิด", "อุปกรณ์ความปลอดภัย", "รองเท้านิรภัย", "รองเท้าเซฟตี้", "หมวกนิรภัย", "แว่นตานิรภัย", "เสื้อสะท้อนแสง", "กรวยจราจร", "ป้ายเตือน", 
        "เครื่องมือช่าง", "สว่าน", "เครื่องเจียร", "ตลับเมตร", "กล่องเครื่องมือ", "บันไดอลูมิเนียม", "เทปพันสายไฟ", "ปั๊มน้ำ", "มอเตอร์", "เครื่องปั่นไฟ", "สายยาง"
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
        main_dept = str(row.get("ชื่อหน่วยงาน", "")).strip()
        main_dept_id = normalize_dept_id(row.get("รหัสหน่วยงาน", None))
        
        if main_dept_id and main_dept_id.isdigit() and len(main_dept_id) < 4:
            main_dept_id = main_dept_id.zfill(4)

        if not main_dept_id or main_dept_id.lower() == "nan":
            continue

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
        time.sleep(1.5)

    print("\n" + "=" * 60)
    if all_projects:
        df_result = pd.DataFrame(all_projects)
        output_file = f"eGP_Main_Result_{target_date_str.replace('-','')}.xlsx"
        df_result.to_excel(output_file, index=False)
        print(f"🎉 พบประกาศที่ตรงคีย์เวิร์ด {len(all_projects)} รายการ")
        
        # ==========================================
        # ส่วนที่เพิ่มใหม่: ระบบบันทึกไฟล์ Backup รายเดือน
        # ==========================================
        backup_dir = "Backup"
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
            
        current_month_str = datetime.now().strftime('%Y%m')
        monthly_file = os.path.join(backup_dir, f"eGP_Monthly_Main_{current_month_str}.xlsx")
        
        if os.path.exists(monthly_file):
            try:
                df_old = pd.read_excel(monthly_file)
                df_combined = pd.concat([df_old, df_result], ignore_index=True)
                # ตัดข้อมูลที่ซ้ำกันออก เผื่อรันซ้ำ
                df_combined.drop_duplicates(subset=["ชื่อโครงการจัดซื้อ", "รหัสหน่วยงานหลัก"], keep="last", inplace=True)
                df_combined.to_excel(monthly_file, index=False)
                print(f"💾 อัปเดตข้อมูลสะสมเข้าไฟล์รายเดือนเรียบร้อย")
            except Exception as e:
                print(f"⚠️ ไม่สามารถอัปเดตไฟล์รายเดือนได้: {e}")
        else:
            df_result.to_excel(monthly_file, index=False)
            print(f"📁 เริ่มสร้างไฟล์คลังข้อมูลประจำเดือนชุดใหม่เรียบร้อย")
        # ==========================================
            
        send_email_report(output_file, target_date_str, has_data=True)
    else:
        print("🏁 ไม่พบประกาศจัดซื้อใหม่ที่ตรงคีย์เวิร์ด")
        send_email_report(None, target_date_str, has_data=False)
        
    # สั่งเช็คว่าวันนี้เป็นวันที่ 1 หรือไม่ ถ้าใช่ให้ส่งอีเมลของเดือนก่อน
    check_and_process_monthly_report()
    print("=" * 60)

if __name__ == "__main__":
    fetch_main_dept_egp()
