import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบอ่านไฟล์นับสต็อกรายสัปดาห์", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบและบันทึกประวัติสต็อกรายสัปดาห์")
st.write(
    "อัปโหลดไฟล์ Excel หลายๆ ไฟล์พร้อมกัน ระบบจะรวมยอดแยกตามรายสัปดาห์ให้ทันที"
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

if uploaded_files:
    all_file_data = []
    file_names_list = []

    for uploaded_file in uploaded_files:
        try:
            xls = pd.ExcelFile(uploaded_file)
            file_names_list.append(uploaded_file.name)

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

                # ตรวจสอบคอลัมน์สำคัญ (รอบนี้ดึงคอลัมน์ DATE มาคำนวณสัปดาห์ด้วย)
                if (
                    "DATE" in df.columns
                    and "EMP ID" in df.columns
                    and "STATION" in df.columns
                    and "SN" in df.columns
                ):
                    # คลีนข้อมูลเบื้องต้น
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    df_clean["EMP ID"] = (
                        df_clean["EMP ID"].astype(str).str.strip()
                    )
                    df_clean["STATION"] = (
                        df_clean["STATION"].astype(str).str.strip()
                    )

                    # จัดการเรื่องวันที่ (แปลงเป็น datetime ของ pandas)
                    # ffill() ช่วยเติมวันที่ในแถวที่ปล่อยว่างไว้ใน Excel
                    df_clean["DATE"] = df_clean["DATE"].ffill()
                    df_clean["DATE_PARSED"] = pd.to_datetime(
                        df_clean["DATE"], errors="coerce", dayfirst=True
                    )

                    # สร้างชื่อสัปดาห์ เช่น "ปี 2026 - สัปดาห์ที่ 21"
                    df_clean["สัปดาห์"] = df_clean["DATE_PARSED"].dt.strftime(
                        "ปี %Y - สัปดาห์ที่ %U"
                    )
                    # หากแถวไหนไม่มีวันที่ ให้ระบุว่า ไม่ระบุวันที่
                    df_clean["สัปดาห์"] = df_clean["สัปดาห์"].fillna(
                        "ไม่ระบุวันที่ในไฟล์"
                    )

                    all_file_data.append(
                        df_clean[["สัปดาห์", "EMP ID", "STATION", "SN"]]
                    )
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        combined_df = pd.concat(all_file_data, ignore_index=True)

        # คำนวณสรุปยอดรวมแยกพนักงานตามสัปดาห์
        weekly_emp_df = (
            combined_df.groupby(["สัปดาห์", "EMP ID"])["SN"]
            .count()
            .reset_index()
        )
        weekly_emp_df.columns = [
            "สัปดาห์ที่ทำงาน",
            "รหัสพนักงาน (EMP ID)",
            "จำนวนรวมที่นับได้ (ตัว)",
        ]

        # --------------------------------------------------
        # ✨ แสดงผลป๊อปอัปแจ้งเตือนกล่องเขียวรายบุคคล
        # --------------------------------------------------
        st.success(f"🎉 อัปโหลดสำเร็จ {len(uploaded_files)} ไฟล์!")

        # 1. แสดงการ์ดสรุปยอดแยกตามสัปดาห์และรายบุคคล
        st.subheader("🗓️ สรุปยอดรวมแยกตามสัปดาห์")
        for _, row in weekly_emp_df.iterrows():
            st.markdown(
                f"📅 **{row['สัปดาห์ที่ทำงาน']}** ➡️ พนักงานรหัส: **{row['รหัสพนักงาน (EMP ID)']}** นับได้รวมทั้งหมด **{row['จำนวนรวมที่นับได้ (ตัว)']}** ตัว"
            )

        # 2. เพิ่มกราฟแท่งแสดงให้เห็นภาพชัดเจน (Visual Chart)
        fig = px.bar(
            weekly_emp_df,
            x="สัปดาห์ที่ทำงาน",
            y="จำนวนรวมที่นับได้ (ตัว)",
            color="รหัสพนักงาน (EMP ID)",
            barmode="group",
            title="กราฟเปรียบเทียบยอดนับสต็อกของพนักงานในแต่ละสัปดาห์",
        )
        st.plotly_chart(fig, use_container_width=True)

        # 3. แสดงตารางดีเทลแยกตามสถานีสินค้าแบบละเอียด
        with st.expander("🔍 คลิกเพื่อดูตารางรายละเอียดแยกตามประเภทสินค้า"):
            detail_df = (
                combined_df.groupby(["สัปดาห์", "EMP ID", "STATION"])["SN"]
                .count()
                .reset_index()
            )
            detail_df.columns = [
                "สัปดาห์",
                "รหัสพนักงาน",
                "สินค้า/สถานี",
                "จำนวน (ตัว)",
            ]
            st.dataframe(detail_df, use_container_width=True, hide_index=True)

        # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสม
        if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            files_string = ", ".join(file_names_list)

            # นำข้อมูลกลุ่มสัปดาห์ไปบันทึกเก็บไว้
            for _, row in weekly_emp_df.iterrows():
                st.session_state.history_log.append(
                    {
                        "เวลาที่บันทึกระบบ": current_time,
                        "จากไฟล์": files_string,
                        "สัปดาห์": row["สัปดาห์ที่ทำงาน"],
                        "รหัสพนักงาน (EMP ID)": row["รหัสพนักงาน (EMP ID)"],
                        "จำนวนรวมสะสม (ตัว)": row["จำนวนรวมที่นับได้ (ตัว)"],
                    }
                )
            st.toast("บันทึกข้อมูลเข้าคลังประวัติเรียบร้อยแล้ว!")
            st.rerun()
    else:
        st.error(
            "❌ ไฟล์ที่อัปโหลดไม่มีโครงสร้างคอลัมน์หลัก (DATE, EMP ID, STATION, SN) กรุณาตรวจสอบหัวตาราง"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสมย้อนหลัง (Historical Data)
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)
    total_accumulated = history_df["จำนวนรวมสะสม (ตัว)"].sum()

    st.metric("ยอดนับสินค้าสะสมรวมในระบบทั้งหมด", f"{total_accumulated:,} ตัว")
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info(
        "ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง กรุณาอัปโหลดไฟล์แล้วกดปุ่มบันทึกด้านบน"
    )
