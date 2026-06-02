import datetime
import io  # เพิ่มส่วนประกอบสำหรับจัดการแปลงข้อมูลเป็นไฟล์ให้ดาวน์โหลด
import pandas as pd
import plotly.express as px
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบสต็อกสินค้าจริง", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบและส่งออกประวัติสต็อก (เฉพาะงานจริง)")
st.write(
    "ระบบนับยอดชิ้นงานผลิตจริงตามรหัสสินค้า (SN): ไม่รวม MASTER, ตัดตัวซ้ำ, แก้รหัสพิมพ์สลับ ตัว O ➡️ เลข 0 และส่งออกไฟล์ Excel ได้"
)

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

            # ลิสต์สรุปพนักงานแยกบรรทัดให้อ่านง่าย
            for _, row in emp_total_df.iterrows():
                st.markdown(
                    f"👤 รหัสพนักงาน: **{row['รหัสพนักงาน (EMP ID)']}** ➡️ ยอดปรับสินค้าจริงรวมทั้งหมด **{row['จำนวนรวมแท้จริง (ตัว)']:,}** ตัว"
                )

            # --------------------------------------------------
            # 📥 [ฟังก์ชันใหม่] ระบบส่งออกเป็นไฟล์ Excel (.xlsx)
            # --------------------------------------------------
            st.markdown("### 📥 ดาวน์โหลดรายงานสรุป")
            
            # แปลงข้อมูลในตาราง emp_total_df ให้กลายเป็นไฟล์ Excel ในหน่วยความจำ (Buffer)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                emp_total_df.to_excel(writer, index=False, sheet_name="Summary")
            buffer.seek(0)

            # สร้างปุ่มสำหรับกดดาวน์โหลดไฟล์ Excel ออกไปนอกระบบ
            st.download_button(
                label="🟢 คลิกที่นี่เพื่อดาวน์โหลดรายงานสรุปเป็นไฟล์ Excel",
                data=buffer,
                file_name=f"สรุปยอดปรับรุ่นจริง_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # 3. แสดงกราฟแท่งเปรียบเทียบยอดงานจริงรายคน
            st.markdown("---")
            fig = px.bar(
                emp_total_df,
                x="รหัสพนักงาน (EMP ID)",
                y="จำนวนรวมแท้จริง (ตัว)",
                color="รหัสพนักงาน (EMP ID)",
                title="กราฟเปรียบเทียบจำนวนชิ้นงานผลิตจริงของพนักงาน (รอบปัจจุบัน)",
            )
            st.plotly_chart(fig, use_container_width=True)

            # ตารางรายละเอียดสินค้าแยกตามสถานีจริงแบบไม่ซ้ำ
            with st.expander("🔍 คลิกเพื่อดูตารางรายละเอียดแยกตามรหัสพนักงานและรุ่นสินค้า"):
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

            # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสม
            if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
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
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง (เฉพาะตัวงานจริงเท่านั้น)")

if st.session_state.history_log:
    history_df = pd.DataFrame(st.session_state.history_log)
    total_accumulated = history_df["จำนวนรวมสะสม (ตัว)"].sum()

    st.metric("ยอดนับสินค้าสะสมรวมในระบบทั้งหมด (ไม่รวม MASTER)", f"{total_accumulated:,} ตัว")
    st.dataframe(history_df, use_container_width=True, hide_index=True)

    # 📥 [ฟังก์ชันใหม่] ปุ่มส่งออกไฟล์ Excel สำหรับตารางประวัติสะสมย้อนหลังระยะยาว
    buffer_history = io.BytesIO()
    with pd.ExcelWriter(buffer_history, engine="openpyxl") as writer:
        history_df.to_excel(writer, index=False, sheet_name="History_Log")
    buffer_history.seek(0)

    st.download_button(
        label="🔵 คลิกที่นี่เพื่อดาวน์โหลดคลังประวัติสะสมทั้งหมดเป็นไฟล์ Excel",
        data=buffer_history,
        file_name=f"คลังประวัติสต็อกสะสม_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.write("") # เว้นช่องไฟ

    if st.button("🗑️ ล้างประวัติยอดรวมทั้งหมด"):
        st.session_state.history_log = []
        st.rerun()
else:
    st.info("ℹ️ ยังไม่มีประวัติยอดรวมสะสมในคลัง")
