import datetime
import io  # เพิ่มส่วนประกอบสำหรับจัดการแปลงข้อมูลเป็นไฟล์ให้ดาวน์โหลด
import pandas as pd
import plotly.express as px
import streamlit as st

# ตั้งค่าหน้าเว็บให้ดูกว้างและมินิมอลขึ้น
st.set_page_config(
    page_title="ระบบสต็อกสินค้าจริง", page_icon="📦", layout="wide"
)

# ตกแต่ง UI ให้มินิมอล สบายตา และจัดระเบียบองค์ประกอบ
st.markdown("""
    <style>
    /* จัดการระยะขอบหน้าเว็บ */
    .block-container { padding-top: 2.5rem; padding-bottom: 2rem; max-width: 1200px; }
    /* สไตล์หัวข้อหลัก */
    .main-title { font-size: 28px; font-weight: 700; color: #1E293B; margin-bottom: 4px; }
    .sub-title { font-size: 14px; color: #64748B; margin-bottom: 24px; }
    /* ปรับแต่งปุ่มกดดาวน์โหลดและบันทึก */
    .stButton>button, .stDownloadButton>button {
        border-radius: 6px !important;
        font-weight: 500 !important;
    }
    /* แยกกรอบพื้นที่อัปโหลดให้ดูสะอาด */
    .stFileUploader { background-color: #F8FAFC; border: 1px dashed #E2E8F0; padding: 12px; border-radius: 8px; }
    /* สไตล์สำหรับตัวเลข Total สรุปผลยอด */
    .total-box { background-color: #F1F5F9; border-left: 4px solid #10B981; padding: 12px 16px; border-radius: 6px; margin-top: 15px; margin-bottom: 15px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📦 ระบบตรวจสอบและส่งออกประวัติสต็อก (เฉพาะงานจริง)</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">ระบบนับยอดชิ้นงานผลิตจริงตามรหัสสินค้า (SN): Automatic Count System </p>', unsafe_allow_html=True)

# --------------------------------------------------
# ระบบฐานข้อมูลจำลอง (ประวัติยอดรวมสะสม)
# --------------------------------------------------
if "history_log" not in st.session_state:
    st.session_state.history_log = []

# ส่วนอัปโหลดไฟล์
uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ Excel ของคุณมาวางที่นี่ (เลือกได้หลายไฟล์พร้อมกัน)",
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

                # ตรวจสอบโครงสร้างคอลัมน์สำคัญ
                if (
                    "EMP ID" in df.columns
                    and "STATION" in df.columns
                    and "SN" in df.columns
                ):
                    # 1. คลีนข้อมูลเบื้องต้น ลบแถวว่างและแถวหัวตารางที่เกินมาออก
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    # 2. ตัดช่องว่าง (Space) หน้า-หลังข้อความทั้งหมด
                    df_clean["EMP ID"] = (
                        df_clean["EMP ID"].astype(str).str.strip()
                    )
                    df_clean["STATION"] = (
                        df_clean["STATION"].astype(str).str.strip()
                    )
                    df_clean["SN"] = df_clean["SN"].astype(str).str.strip()

                    # 🚨 แก้ปัญหาพนักงานพิมพ์รหัสสลับ ตัว O กับ เลข 0
                    df_clean["EMP ID"] = (
                        df_clean["EMP ID"]
                        .str.replace("O", "0", case=False)
                    )

                    # 🚨 กรองตัดสถานีที่เป็น MASTER หรือ master ออกไปเลย 100%
                    df_clean = df_clean[
                        df_clean["STATION"].str.upper() != "MASTER"
                    ]

                    all_file_data.append(df_clean[["EMP ID", "STATION", "SN"]])
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        # รวมข้อมูลดิบจากทุกไฟล์เข้าด้วยกัน
        raw_combined_df = pd.concat(all_file_data, ignore_index=True)

        # 🎯 [ตัดตัวซ้ำ] ยุบแถวที่ SN ซ้ำกันให้เหลือชิ้นเดียว ได้ยอดเนื้องานที่ปรับจริง
        combined_df = raw_combined_df.drop_duplicates(subset=["SN"], keep="first")

        if not combined_df.empty:
            # คำนวณสรุปยอดรวมสุทธิแยกรายบุคคล (เฉพาะสินค้าจริง)
            emp_total_df = (
                combined_df.groupby("EMP ID")["SN"].count().reset_index()
            )
            emp_total_df.columns = ["รหัสพนักงาน (EMP ID)", "จำนวนรวมแท้จริง (ตัว)"]

            # คำนวณยอดรวมทั้งหมด (Total Sum) ในรอบปัจจุบัน
            current_total_sum = emp_total_df["จำนวนรวมแท้จริง (ตัว)"].sum()

            # --------------------------------------------------
            # ✨ ส่วนการแสดงผลป๊อปอัปและข้อความสรุปรายคน
            # --------------------------------------------------
            popup_message = "🔔 สรุปผลยอดนับสินค้าจริง: \n"
            for _, row in emp_total_df.iterrows():
                popup_message += f"- พนักงาน {row['รหัสพนักงาน (EMP ID)']} ปรับได้ {row['จำนวนรวมแท้จริง (ตัว)']} ตัว\n"
            st.toast(popup_message, icon="📊")

            st.success(
                f"🎉 รวมข้อมูลสำเร็จทั้งหมด {len(uploaded_files)} ไฟล์! (คัดเฉพาะชิ้นงานจริงที่ปรับรุ่นสำเร็จ)"
            )

            # 📦 ปรับปรุง UI หน้าสรุปปัจจุบัน: แบ่งฝั่งแสดงผลอย่างมีระเบียบ (ตาราง+ดาวน์โหลด อยู่ซ้าย / กราฟ อยู่ขวา)
            col_left, col_right = st.columns([4, 5])

            with col_left:
                st.markdown("### 📋 ข้อมูลสรุปพนักงาน")
                
                # ลิสต์สรุปพนักงานแยกบรรทัดสไตล์มินิมอล
                for _, row in emp_total_df.iterrows():
                    st.markdown(
                        f"👤 รหัสพนักงาน: **{row['รหัสพนักงาน (EMP ID)']}** ➡️ ยอดปรับจริงรวม **{row['จำนวนรวมแท้จริง (ตัว)']:,}** ตัว"
                    )

                # แสดงกล่องผลรวม Total รวมของทุกคนในหน้านี้
                st.markdown(
                    f'<div class="total-box"><b>📊 ยอดรวมทั้งหมดประจำรอบนี้ (Total):</b> <span style="color:#10B981; font-size:18px; font-weight:700;">{current_total_sum:,}</span> ตัว</div>', 
                    unsafe_allow_html=True
                )
                
                st.markdown("---")
                
                # ✨ [ฟังก์ชันใหม่] ระบบเลือกช่วงวันที่ปฏิบัติงาน (Date Range)
                st.write("**📅 ป้อนช่วงข้อมูลวันที่สำหรับรายงาน Excel**")
                
                # ผู้ใช้สามารถคลิกและลากเลือกช่วงวันที่ได้ เช่น วันที่ 25 พ.ค. ถึง 1 มิ.ย.
                date_range = st.date_input(
                    "เลือกช่วงวันที่ปฏิบัติงาน",
                    value=(datetime.date.today() - datetime.timedelta(days=7), datetime.date.today())
                )
                
                # แปลงช่วงวันที่เลือกออกมาเป็นข้อความปั๊มหัวกระดาษและชื่อไฟล์
                if isinstance(date_range, tuple) and len(date_range) == 2:
                    start_date, end_date = date_range
                    date_string = f"{start_date.strftime('%Y-%m-%d')} - {end_date.strftime('%Y-%m-%d')}"
                elif isinstance(date_range, tuple) and len(date_range) == 1:
                    start_date = date_range[0]
                    date_string = f"{start_date.strftime('%Y-%m-%d')}"
                else:
                    date_string = datetime.date.today().strftime('%Y-%m-%d')
                
                # --------------------------------------------------
                # 📥 ระบบส่งออกเป็นไฟล์ Excel (.xlsx) แบบมีช่วงวันที่ + มี Total
                # --------------------------------------------------
                st.write("**📥 ดาวน์โหลดรายงานสรุป**")
                
                # โครงสร้างตารางตามที่คุณต้องการ: วันที่ | รหัสพนักงาน | จำนวน
                excel_df = pd.DataFrame({
                    "วันที่": date_string,
                    "รหัสพนักงาน": emp_total_df["รหัสพนักงาน (EMP ID)"],
                    "จำนวนรวมแท้จริง (ตัว)": emp_total_df["จำนวนรวมแท้จริง (ตัว)"]
                })
                
                # เพิ่มบรรทัดสรุปรวม (Total) ท้ายตารางใน Excel
                total_row = pd.DataFrame([{
                    "วันที่": "Total",
                    "รหัสพนักงาน": "",
                    "จำนวนรวมแท้จริง (ตัว)": current_total_sum
                }])
                excel_final_df = pd.concat([excel_df, total_row], ignore_index=True)

                # แปลงข้อมูลให้กลายเป็นไฟล์ Excel ในหน่วยความจำ (Buffer)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    excel_final_df.to_excel(writer, index=False, sheet_name="Summary")
                buffer.seek(0)

                # สร้างปุ่มสำหรับกดดาวน์โหลดไฟล์ Excel ออกไปนอกระบบ
                st.download_button(
                    label="🟢 ดาวน์โหลดรายงานเป็นไฟล์ Excel (ระบุช่วงวันที่ + Total)",
                    data=buffer,
                    file_name=f"สรุปยอดปรับรุ่นจริง_({date_string.replace(' ', '')}).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสม
                if st.button("📥 บันทึกชุดนี้ลงประวัติสะสม", use_container_width=True):
                    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    files_string = ", ".join(file_names_list)

                    for _, row in emp_total_df.iterrows():
                        st.session_state.history_log.append(
                            {
                                "เวลาที่บันทึกระบบ": current_time,
                                "จากไฟล์ทั้งหมด": files_string,
                                "รหัสพนักงาน (EMP ID)": row["รหัสพนักงาน (EMP ID)"],
                                "จำนวนรวมสะสม (ตัว)": row["จำนวนรวมแท้จริง (ตัว)"],
                            }
                        )
                    st.toast("บันทึกข้อมูลเข้าคลังประวัติเรียบร้อยแล้ว!")
                    st.rerun()

            with col_right:
                # 3. แสดงกราฟแท่งเปรียบเทียบยอดงานจริงรายคน (ปรับโทนสีพาสเทลให้มินิมอล)
                fig = px.bar(
                    emp_total_df,
                    x="รหัสพนักงาน (EMP ID)",
                    y="จำนวนรวมแท้จริง (ตัว)",
                    color="รหัสพนักงาน (EMP ID)",
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    title="สัดส่วนจำนวนชิ้นงานผลิตจริงของพนักงาน",
                )
                fig.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    margin=dict(t=40, b=0, l=0, r=0)
                )
                st.plotly_chart(fig, use_container_width=True)

            # ตารางรายละเอียดสินค้าแยกตามสถานีจริงแบบไม่ซ้ำ
            st.markdown("<br>", unsafe_allow_html=True)
            with st.expander("🔍 คลิกเพื่อเรียกดูรายละเอียดแยกตามพนักงานและรุ่นสินค้าเพิ่มเติม"):
                detail_df = (
                    combined_df.groupby(["EMP ID", "STATION"])["SN"]
                    .count()
                    .reset_index()
                )
                detail_df.columns = [
                    "รหัสพนักงาน",
                    "รุ่นสินค้า/สถานี",
                    "จำนวนจริง (ตัว)",
                ]
                st.dataframe(detail_df, use_container_width=True, hide_index=True)

        else:
            st.warning("⚠️ ไม่พบข้อมูลสินค้าอื่นนอกเหนือจากสถานี MASTER เลย")
    else:
        st.error(
            "❌ ไม่พบโครงสร้างข้อมูลที่ถูกต้อง กรุณาตรวจสอบหัวตารางไฟล์ Excel"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสมย้อนหลัง
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)
    total_accumulated = history_df["จำนวนรวมสะสม (ตัว)"].sum()

    col_metric, col_btn = st.columns([3, 7])
    with col_metric:
        st.metric("ยอดนับสะสมรวมในระบบ (ไม่รวม MASTER)", f"{total_accumulated:,} ตัว")
    
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    # 📥 ปุ่มส่งออกไฟล์ Excel สำหรับตารางประวัติสะสมย้อนหลังระยะยาว
    buffer_history = io.BytesIO()
    with pd.ExcelWriter(buffer_history, engine="openpyxl") as writer:
        history_df.to_excel(writer, index=False, sheet_name="History_Log")
    buffer_history.seek(0)

    col_dl, col_cl = st.columns(2)
    with col_dl:
        st.download_button(
            label="🔵 ดาวน์โหลดคลังประวัติสะสมทั้งหมด (Excel)",
            data=buffer_history,
            file_name=f"คลังประวัติสต็อกสะสม_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_cl:
        if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด", type="primary", use_container_width=True):
            st.session_state.history_log = []
            st.rerun()
else:
    st.info("ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง")
