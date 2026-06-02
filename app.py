import datetime
import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบอ่านไฟล์นับสต็อกแบบหลายไฟล์", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบและบันทึกประวัติสต็อก (อัปโหลดหลายไฟล์)")
st.write(
    "คุณสามารถเลือกหรือลากไฟล์ Excel หลายๆ ไฟล์มาวางพร้อมกันได้ ระบบจะรวมยอดให้อัตโนมัติ"
)

# --------------------------------------------------
# ระบบฐานข้อมูลจำลอง (ประวัติยอดรวมสะสม)
# --------------------------------------------------
if "history_log" not in st.session_state:
    st.session_state.history_log = []

# ส่วนอัปโหลดไฟล์ (เปิดใช้งาน accept_multiple_files=True)
uploaded_files = st.file_uploader(
    "เลือกไฟล์ Excel ของคุณ (เลือกได้ทีละหลายไฟล์)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

# ตรวจสอบว่ามีการอัปโหลดไฟล์เข้ามาอย่างน้อย 1 ไฟล์ไหม
if uploaded_files:
    all_file_data = []
    file_names_list = []

    # วนลูปอ่านทีละไฟล์ที่อัปโหลดเข้ามา
    for uploaded_file in uploaded_files:
        try:
            xls = pd.ExcelFile(uploaded_file)
            file_names_list.append(uploaded_file.name)

            # วนลูปอ่านทุกชีตในไฟล์นั้นๆ
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

                required_cols = ["EMP ID", "STATION", "SN"]
                if "EMP ID" in df.columns and "STATION" in df.columns and "SN" in df.columns:
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    df_clean["EMP ID"] = df_clean["EMP ID"].astype(str).str.strip()
                    df_clean["STATION"] = df_clean["STATION"].astype(str).str.strip()

                    all_file_data.append(df_clean[["EMP ID", "STATION", "SN"]])
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    # ถ้ารวมข้อมูลจากทุกไฟล์แล้วมีข้อมูลที่ใช้งานได้
    if all_file_data:
        combined_df = pd.concat(all_data := all_file_data, ignore_index=True)

        # คำนวณสรุปรวมของทุกไฟล์ที่อัปโหลดเข้ามาในรอบนี้
        summary_df = (
            combined_df.groupby(["EMP ID", "STATION"])["SN"]
            .count()
            .reset_index()
        )
        summary_df.columns = [
            "รหัสพนักงาน (EMP ID)",
            "สินค้า/สถานี (STATION)",
            "จำนวนที่นับได้ (ตัว)",
        ]

        st.success(f"✅ อ่านข้อมูลสำเร็จทั้งหมด {len(uploaded_files)} ไฟล์!")

        # แสดงตารางผลลัพธ์รวมของไฟล์ชุดนี้
        st.subheader("📊 ผลการนับสต็อกจากชุดไฟล์ที่อัปโหลดปัจจุบัน")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        total_items = len(combined_df)
        unique_staff = combined_df["EMP ID"].nunique()

        st.write(
            f"💡 ชุดไฟล์นี้มีพนักงานทำงานรวม **{unique_staff} คน** ยอดนับรวมทั้งหมด **{total_items} ตัว**"
        )

        # ปุ่มกดบันทึกข้อมูลเข้าประวัติรวมสะสม
        if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            files_string = ", ".join(file_names_list)

            # บันทึกข้อมูลลงในประวัติย้อนหลัง
            for _, row in summary_df.iterrows():
                st.session_state.history_log.append(
                    {
                        "เวลาที่บันทึก": current_time,
                        "จากไฟล์ทั้งหมด": files_string,
                        "รหัสพนักงาน (EMP ID)": row["รหัสพนักงาน (EMP ID)"],
                        "สินค้า/สถานี (STATION)": row["สินค้า/สถานี (STATION)"],
                        "จำนวนที่นับได้ (ตัว)": row["จำนวนที่นับได้ (ตัว)"],
                    }
                )
            st.toast("บันทึกข้อมูลเข้าประวัติรวมเรียบร้อยแล้ว!")
            st.rerun()
    else:
        st.error(
            "❌ ไฟล์ที่อัปโหลดเข้ามาไม่มีคอลัมน์ 'EMP ID', 'STATION' หรือ 'SN' เลย กรุณาตรวจสอบหัวตาราง"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสม (Historical Dashboard)
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสม (Historical Data)")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)

    # แสดง KPI ยอดสะสมรวมตั้งแต่เปิดเว็บมา
    total_accumulated_count = history_df["จำนวนที่นับได้ (ตัว)"].sum()

    col1, col2 = st.columns(2)
    with col1:
        st.metric("บันทึกข้อมูลไปแล้วทั้งหมด", f"{len(history_df)} รายการกลุ่ม")
    with col2:
        st.metric("ยอดนับสินค้าสะสมรวมทุกไฟล์", f"{total_accumulated_count:,} ตัว")

    st.write("**ตารางประวัติการบันทึกข้อมูลสะสมย้อนหลัง:**")
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    # ปุ่มล้างประวัติ
    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info(
        "ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง กรุณาอัปโหลดไฟล์แล้วกดปุ่มบันทึกด้านบน"
    )
