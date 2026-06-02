import datetime
import pandas as pd
import plotly.express as px
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบสต็อกอัจฉริยะ (ไม่นับ MASTER)", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบสต็อกรายสัปดาห์ (ไม่รวมยอด MASTER)")
st.write(
    "อัปโหลดไฟล์ Excel หลายๆ ไฟล์พร้อมกัน ระบบจะตัดสินค้าซ้ำ และ **ไม่นับสถานี MASTER** ให้โดยอัตโนมัติ"
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

                # ตรวจสอบโครงสร้างคอลัมน์สำคัญ
                if (
                    "DATE" in df.columns
                    and "EMP ID" in df.columns
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

                    # 🚨 [จุดสำคัญที่เพิ่มเข้ามา] ทำการคัดออก! ไม่เอาสถานีที่เป็น 'MASTER' หรือ 'master'
                    df_clean = df_clean[
                        df_clean["STATION"].str.upper() != "MASTER"
                    ]

                    # 3. จัดการเติมวันที่ในกรณีที่ Excel ปล่อยเว้นว่างไว้ในแถวถัดๆ มา
                    df_clean["DATE"] = df_clean["DATE"].ffill()
                    df_clean["DATE_PARSED"] = pd.to_datetime(
                        df_clean["DATE"], errors="coerce", dayfirst=True
                    )
                    
                    # 4. แปลงวันที่เป็นเลขสัปดาห์ของปี
                    df_clean["สัปดาห์"] = df_clean["DATE_PARSED"].dt.strftime(
                        "ปี %Y - สัปดาห์ที่ %U"
                    )
                    df_clean["สัปดาห์"] = df_clean["สัปดาห์"].fillna(
                        "ไม่ระบุวันที่ในไฟล์"
                    )

                    all_file_data.append(
                        df_clean[["สัปดาห์", "EMP ID", "STATION", "SN"]]
                    )
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        # รวมข้อมูลดิบจากทุกไฟล์เข้าด้วยกัน
        raw_combined_df = pd.concat(all_file_data, ignore_index=True)

        # 🎯 ตัดรายการที่ "หมายเลขสินค้า (SN)" ซ้ำกันออก ให้เหลือชิ้นเดียว
        combined_df = raw_combined_df.drop_duplicates(subset=["SN"], keep="first")

        if not combined_df.empty:
            # คำนวณสรุปยอดรวมแยกพนักงานตามสัปดาห์ (ไม่มี MASTER และไม่มีของซ้ำ)
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
            # ✨ ส่วนการแสดงผลสรุปบนหน้าเว็บ
            # --------------------------------------------------
            st.success(
                f"🎉 รวมข้อมูลสำเร็จ! (คัดกรองตัวซ้ำและคัดแยกสถานี MASTER ออกให้เรียบร้อยแล้ว)"
            )

            # แสดงข้อความสรุปยอดที่แท้จริงแยกรายบุคคล
            st.subheader("🗓️ ยอดสรุปการทำงานรายสัปดาห์ (เฉพาะตัวสินค้าจริง)")
            for _, row in weekly_emp_df.iterrows():
                st.markdown(
                    f"📅 **{row['สัปดาห์ที่ทำงาน']}** ➡️ พนักงานรหัส: **{row['รหัสพนักงาน (EMP ID)']}** นับสินค้าจริงได้รวม **{row['จำนวนรวมที่นับได้ (ตัว)']}** ตัว"
                )

            # แสดงกราฟแท่งวิเคราะห์ยอด
            fig = px.bar(
                weekly_emp_df,
                x="สัปดาห์ที่ทำงาน",
                y="จำนวนรวมที่นับได้ (ตัว)",
                color="รหัสพนักงาน (EMP ID)",
                barmode="group",
                title="กราฟแสดงยอดนับสต็อกจริงในแต่ละสัปดาห์ (ไม่นับ MASTER)",
            )
            st.plotly_chart(fig, use_container_width=True)

            # ตารางรายละเอียดแยกตามประเภทตัวสินค้าจริง
            with st.expander("🔍 คลิกเพื่อดูตารางรายชื่อสินค้า (SN) ทั้งหมดแบบละเอียด"):
                detail_df = (
                    combined_df.groupby(["สัปดาห์", "EMP ID", "STATION"])["SN"]
                    .count()
                    .reset_index()
                )
                detail_df.columns = [
                    "สัปดาห์",
                    "รหัสพนักงาน",
                    "สินค้า/สถานี",
                    "จำนวนจริง (ตัว)",
                ]
                st.dataframe(detail_df, use_container_width=True, hide_index=True)

            # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสม
            if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                files_string = ", ".join(file_names_list)

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
            st.warning("⚠️ หลังจากตัดสถานี MASTER ออกแล้ว ไม่พบข้อมูลตัวงานอื่นในไฟล์เลย")
    else:
        st.error(
            "❌ ไม่พบโครงสร้างข้อมูลที่ถูกต้อง กรุณาตรวจสอบหัวตารางไฟล์ Excel"
        )

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสมย้อนหลัง
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง (ตัวงานจริง)")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)
    total_accumulated = history_df["จำนวนรวมสะสม (ตัว)"].sum()

    st.metric("ยอดนับสินค้าสะสมรวมในระบบทั้งหมด (ไม่รวม MASTER)", f"{total_accumulated:,} ตัว")
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info("ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง")
