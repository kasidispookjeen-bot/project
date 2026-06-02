import datetime
import io
import pandas as pd
import plotly.express as px
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import matplotlib.pyplot as plt
import seaborn as sns

# ตั้งค่าหน้าเว็บให้ดูเป็นแบบกว้างและทันสมัย
st.set_page_config(
    page_title="ระบบตรวจสอบสต็อกสินค้าอัจฉริยะ", 
    page_icon="📦", 
    layout="wide"
)

# ตกแต่งสไตล์ CSS เบื้องต้นให้ปุ่มและองค์ประกอบดูสวยงามขึ้น
st.markdown("""
    <style>
    .main-title { font-size: 32px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #4B5563; margin-bottom: 25px; }
    .stButton>button { border-radius: 8px; font-weight: bold; }
    .reportview-container .main .block-container{ padding-top: 2rem; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📦 ระบบตรวจสอบและส่งออกสต็อกสินค้าปรับรุ่น</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">ระบบคัดกรองผลงานพนักงาน: ไม่รวมสถานี MASTER, ตัดรายการซ้ำตาม SN, สรุปยอดรวม (Total) และส่งออกไฟล์ Excel พร้อมรูปกราฟ</p>', unsafe_allow_html=True)

# --------------------------------------------------
# ระบบฐานข้อมูลจำลอง (ประวัติยอดรวมสะสม)
# --------------------------------------------------
if "history_log" not in st.session_state:
    st.session_state.history_log = []

# แยกส่วนอัปโหลดไฟล์ไว้ด้านบนสุดอย่างเด่นชัด
with st.container():
    uploaded_files = st.file_uploader(
        "📥 ลากและวางไฟล์ Excel ของคุณที่นี่ (เลือกอัปโหลดพร้อมกันได้หลายไฟล์)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

if uploaded_files:
    all_file_data = []
    file_names_list = []

    for uploaded_file in uploaded_files:
        try:
            xls = pd.ExcelFile(uploaded_file)
            file_names_list.append(uploaded_file.name)

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

                # ตรวจสอบโครงสร้างคอลัมน์สำคัญตามหน้างานจริง
                if (
                    "DATE" in df.columns
                    and "EMP ID" in df.columns
                    and "STATION" in df.columns
                    and "SN" in df.columns
                ):
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    # เติมวันที่ลงมาในกรณีช่องว่างเว้นไว้ (Forward Fill)
                    df_clean["DATE"] = df_clean["DATE"].ffill()

                    # ตัดช่องว่างเพื่อป้องกันข้อมูลเพี้ยน
                    df_clean["EMP ID"] = df_clean["EMP ID"].astype(str).str.strip()
                    df_clean["STATION"] = df_clean["STATION"].astype(str).str.strip()
                    df_clean["SN"] = df_clean["SN"].astype(str).str.strip()
                    df_clean["DATE_STR"] = df_clean["DATE"].astype(str).str.strip()

                    # แก้ปัญหาการพิมพ์รหัสพนักงานสลับ ตัว O กับ เลข 0
                    df_clean["EMP ID"] = df_clean["EMP ID"].str.replace("O", "0", case=False)

                    # 🚨 ไม่เอาสถานีทดสอบที่เป็น MASTER
                    df_clean = df_clean[df_clean["STATION"].str.upper() != "MASTER"]

                    all_file_data.append(df_clean[["DATE_STR", "EMP ID", "STATION", "SN"]])
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        # รวมข้อมูลทั้งหมดจากทุกไฟล์เข้าด้วยกัน
        raw_combined_df = pd.concat(all_file_data, ignore_index=True)

        # 🎯 [ตัดตัวซ้ำ] คัดตามรหัสสินค้า (SN) จริง เพื่อได้ยอดผลงานเนื้องานแท้ๆ ไม่บวมตามรอบอุณหภูมิ
        combined_df = raw_combined_df.drop_duplicates(subset=["SN"], keep="first")

        if not combined_df.empty:
            # คำนวณสรุปข้อมูลแยกตามรูปแบบตารางเรฟเฟอเรนซ์
            daily_summary_df = (
                combined_df.groupby(["DATE_STR", "EMP ID", "STATION"])["SN"]
                .count()
                .reset_index()
            )
            daily_summary_df.columns = ["วันที่", "รหัสพนักงาน", "รุ่นสินค้า (STATION)", "จำนวนจริง (ตัว)"]
            
            total_sum = daily_summary_df["จำนวนจริง (ตัว)"].sum()
            unique_emps = daily_summary_df["รหัสพนักงาน"].nunique()

            # สร้างโครงสร้างแท็บเพื่อแบ่งโซนข้อมูลให้สะอาดตา
            tab1, tab2 = st.tabs(["📊 สรุปผลการอัปโหลดรอบนี้", "📜 คลังประวัติยอดรวมสะสมระยะยาว"])

            with tab1:
                st.markdown("### 📈 ภาพรวมรายงานปัจจุบัน")
                
                # แสดงผลกล่องตัวเลขสไตล์ Dashboard หรูหรา
                col1, col2, col3 = st.columns(3)
                col1.metric("📦 ยอดปรับรุ่นรวมสุทธิ (Total)", f"{total_sum:,} ตัว")
                col2.metric("👤 พนักงานที่ปฏิบัติงาน", f"{unique_emps} คน")
                col3.metric("📂 จำนวนไฟล์ที่ประมวลผล", f"{len(uploaded_files)} ไฟล์")

                st.markdown("---")
                
                # แบ่งหน้าจอซ้าย-ขวา (ซ้ายแสดงตารางและปุ่มโหลด / ขวาแสดงกราฟ)
                left_col, right_col = st.columns([1, 1])

                with left_col:
                    st.markdown("#### 📋 ตารางสรุปผลงานรายวัน (ตามรูปแบบรายงาน)")
                    
                    # ตกแต่งข้อมูลตารางดิบเพื่อความสวยงามบนเว็บ
                    st.dataframe(daily_summary_df, use_container_width=True, hide_index=True)

                    # 📥 ฟังก์ชันประกอบตารางสรุปที่มีแถว Total ด้านล่าง + ฝังรูปภาพกราฟ
                    export_df = daily_summary_df.copy()
                    total_row = pd.DataFrame([{
                        "วันที่": "Total", "รหัสพนักงาน": "", "รุ่นสินค้า (STATION)": "", "จำนวนจริง (ตัว)": total_sum
                    }])
                    export_df = pd.concat([export_df, total_row], ignore_index=True)

                    # สร้างรูปกราฟลงในหน่วยความจำชั่วคราวเพื่อส่งเข้าไฟล์ Excel
                    fig_img_buf = io.BytesIO()
                    plt.figure(figsize=(6, 4))
                    sns.barplot(data=daily_summary_df, x="รหัสพนักงาน", y="จำนวนจริง (ตัว)", hue="รุ่นสินค้า (STATION)", palette="Set2")
                    plt.title("Daily Production Summary (Actual Work)")
                    plt.tight_layout()
                    plt.savefig(fig_img_buf, format="png", dpi=150)
                    fig_img_buf.seek(0)
                    plt.close()

                    # เขียนสคริปต์ openpyxl เพื่อประกอบร่างโครงสร้าง Excel ขั้นสูง
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                        export_df.to_excel(writer, index=False, sheet_name="Summary_Report")
                        worksheet = writer.sheets["Summary_Report"]
                        
                        # แทรกรูปกราฟไว้ที่ตำแหน่งคอลัมน์ F (เซลล์ F2) เคียงข้างตารางตัวเลขแบบสวยงาม
                        try:
                            img = OpenpyxlImage(fig_img_buf)
                            worksheet.add_image(img, "F2")
                        except:
                            pass
                    excel_buffer.seek(0)

                    # ปุ่มดาวน์โหลดเด่นชัดสีเขียวสดใส
                    st.download_button(
                        label="🟢 ส่งออกรายงานสรุปเป็นไฟล์ Excel (มี Total + รูปกราฟ)",
                        data=excel_buffer,
                        file_name=f"รายงานสต็อกปรับรุ่นจริง_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    # ปุ่มกดสำหรับบันทึกงานเข้าฐานข้อมูลคลังสะสม
                    if st.button("📥 บันทึกข้อมูลชุดนี้เข้าสู่คลังประวัติสะสม", use_container_width=True):
                        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        files_string = ", ".join(file_names_list)

                        for _, row in daily_summary_df.iterrows():
                            st.session_state.history_log.append({
                                "เวลาที่บันทึกระบบ": current_time,
                                "จากไฟล์ทั้งหมด": files_string,
                                "วันที่ทำงาน": row["วันที่"],
                                "รหัสพนักงาน": row["รหัสพนักงาน"],
                                "รุ่นสินค้า (STATION)": row["รุ่นสินค้า (STATION)"],
                                "จำนวนรวมสะสม (ตัว)": row["จำนวนจริง (ตัว)"],
                            })
                        st.toast("บันทึกข้อมูลเข้าคลังประวัติเรียบร้อยแล้ว!", icon="💾")
                        st.rerun()

                with right_col:
                    st.markdown("#### 📊 กราฟวิเคราะห์สถิติการทำงาน")
                    # สร้างกราฟInteractiveที่ผู้ใช้สามารถเอาเมาส์ไปชี้ส่องดูเลขได้ลื่นไหล
                    fig = px.bar(
                        daily_summary_df,
                        x="วันที่",
                        y="จำนวนจริง (ตัว)",
                        color="รุ่นสินค้า (STATION)",
                        facet_col="รหัสพนักงาน",
                        barmode="group",
                        color_discrete_sequence=px.colors.qualitative.Pastel
                    )
                    fig.update_layout(margin=dict(t=30, b=10, l=10, r=10))
                    st.plotly_chart(fig, use_container_width=True)

            with tab2:
                st.markdown("### 📜 คลังเก็บประวัติข้อมูลสต็อกสะสมระยะยาว")
                if st.session_state.history_log:
                    history_df = pd.DataFrame(st.session_state.history_log)
                    total_accumulated = history_df["จำนวนรวมสะสม (ตัว)"].sum()

                    # แสดงผลยอดสะสมตั้งแต่อดีตแบบตระการตา
                    st.metric("ยอดนับสินค้าสะสมรวมทั้งหมดในระบบ", f"{total_accumulated:,} ตัว")
                    st.dataframe(history_df, use_container_width=True, hide_index=True)

                    # ปุ่มสำหรับส่งออกเฉพาะประวัติสะสมทั้งหมดเป็นไฟล์ Excel
                    buffer_history = io.BytesIO()
                    with pd.ExcelWriter(buffer_history, engine="openpyxl") as writer:
                        history_df.to_excel(writer, index=False, sheet_name="History_Log")
                    buffer_history.seek(0)

                    st.download_button(
                        label="🔵 ดาวน์โหลดคลังประวัติสะสมทั้งหมดออกเป็นไฟล์ Excel",
                        data=buffer_history,
                        file_name=f"คลังประวัติสต็อกสะสมรวม_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("🗑️ ล้างข้อมูลประวัติสะสมทั้งหมด", type="primary", use_container_width=True):
                        st.session_state.history_log = []
                        st.rerun()
                else:
                    st.info("ℹ️ ปัจจุบันคลังประวัติสะสมยังคงว่างอยู่ คุณสามารถกดปุ่มบันทึกจากแท็บแรกเพื่อจัดเก็บข้อมูลได้")
        else:
            st.warning("⚠️ ไม่พบข้อมูลสินค้าอื่นนอกเหนือจากสถานี MASTER เลย")
    else:
        st.error("❌ ไม่พบโครงสร้างข้อมูลคอลัมน์ที่กำหนด กรุณาตรวจสอบหัวตารางไฟล์ Excel อีกครั้ง")
