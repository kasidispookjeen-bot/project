import datetime
import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบอ่านไฟล์นับสต็อกอัจฉริยะ", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบและบันทึกประวัติสต็อก (อัปโหลดหลายไฟล์)")
st.write(
    "อัปโหลดไฟล์ Excel หลายๆ ไฟล์พร้อมกัน ระบบจะรวมยอดและแสดงสรุปผลแยกพนักงานให้ทันที"
)

# --------------------------------------------------
# ระบบฐานข้อมูลจำลอง (ประวัติยอดรวมสะสม)
# --------------------------------------------------
if "history_log" not in st.session_state:
    st.session_state.history_log = []

# ส่วนอัปโหลดไฟล์
uploaded_files = st.file_uploader(
    "เลือกหรือลากไฟล์ Excel ของคุณมาวางที่นี่ (เลือกได้หลายไฟล์)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

# ตรวจสอบเมื่อมีการอัปโหลดไฟล์เข้ามา
if uploaded_files:
    all_file_data = []
    file_names_list = []

    for uploaded_file in uploaded_files:
        try:
            xls = pd.ExcelFile(uploaded_file)
            file_names_list.append(uploaded_file.name)

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

                # ตรวจสอบคอลัมน์สำคัญ
                if (
                    "EMP ID" in df.columns
                    and "STATION" in df.columns
                    and "SN" in df.columns
                ):
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    df_clean["EMP ID"] = (
                        df_clean["EMP ID"].astype(str).str.strip()
                    )
                    df_clean["STATION"] = (
                        df_clean["STATION"].astype(str).str.strip()
                    )

                    all_file_data.append(df_clean[["EMP ID", "STATION", "SN"]])
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        combined_df = pd.concat(all_data := all_file_data, ignore_index=True)

        # คำนวณสรุปรวมแยกตามพนักงานและสถานี เพื่อนำไปแสดงในตาราง
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

        # คำนวณยอดรวมสุทธิแยกรายบุคคล (พนักงาน 1 คน นับได้รวมกี่ตัวจากทุกสินค้า)
        emp_total_df = (
            combined_df.groupby("EMP ID")["SN"].count().reset_index()
        )

        # --------------------------------------------------
        # ✨ ส่วนป๊อปอัปแจ้งเตือนและแสดงยอดแยกรายบุคคล
        # --------------------------------------------------
        # 1. ทำป๊อปอัปข้อความเด้งเตือน (Toast Notification) สรุปยอดพนักงานแต่ละคน
        popup_message = "🔔 สรุปผลการอัปโหลด: \n"
        for _, row in emp_total_df.iterrows():
            popup_message += f"- พนักงานรหัส {row['EMP ID']} นับได้รวม {row['SN']} ตัว\n"

        st.toast(popup_message, icon="📊")

        # 2. ทำกล่องแจ้งเตือนสีเขียวเด่นๆ (Success Box) แยกพนักงานให้เห็นชัดเจนด้านบนสุด
        st.success(f"🎉 อัปโหลดสำเร็จ {len(uploaded_files)} ไฟล์! สรุปยอดรวมแยกพนักงาน:")

        # แสดงเป็นกล่องข้อความแยกบรรทัดให้อ่านง่าย
        for _, row in emp_total_df.iterrows():
            st.markdown(
                f"👤 รหัสพนักงาน: **{row['EMP ID']}** ➡️ นับโปรดัคได้รวมทั้งหมด **{row['SN']:,}** ตัว"
            )

        # --------------------------------------------------
        # ตารางผลลัพธ์แบบละเอียดด้านล่าง
        # --------------------------------------------------
        st.markdown("---")
        st.subheader("📊 ตารางแสดงผลแยกตามสถานี/โปรดัค")
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        # ปุ่มกดบันทึกข้อมูลเข้าประวัติรวมสะสม
        if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            files_string = ", ".join(file_names_list)

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
            "❌ ไฟล์ที่อัปโหลดเข้ามาไม่มีโครงสร้างคอลัมน์ที่ต้องการ กรุณาตรวจสอบหัวตาราง"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสม (Historical Dashboard)
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสม (Historical Data)")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)
    total_accumulated_count = history_df["จำนวนที่นับได้ (ตัว)"].sum()

    col1, col2 = st.columns(2)
    with col1:
        st.metric("บันทึกข้อมูลไปแล้วทั้งหมด", f"{len(history_df)} รายการกลุ่ม")
    with col2:
        st.metric("ยอดนับสินค้าสะสมรวมทุกไฟล์", f"{total_accumulated_count:,} ตัว")

    st.write("**ตารางประวัติการบันทึกข้อมูลสะสมย้อนหลัง:**")
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info(
        "ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง กรุณาอัปโหลดไฟล์แล้วกดปุ่มบันทึกด้านบน"
    )
