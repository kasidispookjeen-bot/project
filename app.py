import datetime
import io
import pandas as pd
import plotly.express as px
import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import matplotlib.pyplot as plt
import seaborn as sns

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="ระบบสต็อกสินค้าจริง (Export Excel สมบูรณ์แบบ)", page_icon="📦", layout="centered"
)

st.title("📦 ระบบตรวจสอบสต็อกสินค้า (ส่งออก Excel พร้อมรูปกราฟและยอด Total)")
st.write(
    "อัปโหลดไฟล์ Excel เพื่อดูสรุปยอดผลิตแยกตามวัน (คัดเฉพาะงานจริง ไม่รวม MASTER, ตัดตัวซ้ำ, แก้ตัว O เป็นเลข 0 และส่งออก Excel มีรูปกราฟ)"
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
                    "DATE" in df.columns
                    and "EMP ID" in df.columns
                    and "STATION" in df.columns
                    and "SN" in df.columns
                ):
                    df_clean = df[df["EMP ID"].notna()]
                    df_clean = df_clean[df_clean["EMP ID"] != "EMP ID"]

                    # เติมวันที่ในแถวที่เว้นว่างไว้ (ลากสูตรลงมา)
                    df_clean["DATE"] = df_clean["DATE"].ffill()

                    # ตัดช่องว่างหน้า-หลังข้อความทั้งหมด
                    df_clean["EMP ID"] = df_clean["EMP ID"].astype(str).str.strip()
                    df_clean["STATION"] = df_clean["STATION"].astype(str).str.strip()
                    df_clean["SN"] = df_clean["SN"].astype(str).str.strip()
                    df_clean["DATE_STR"] = df_clean["DATE"].astype(str).str.strip()

                    # แก้ปัญหาพนักงานพิมพ์รหัสสลับ ตัว O กับ เลข 0
                    df_clean["EMP ID"] = df_clean["EMP ID"].str.replace("O", "0", case=False)

                    # คัดออก! ไม่เอาสถานีที่เป็น 'MASTER' หรือ 'master'
                    df_clean = df_clean[df_clean["STATION"].str.upper() != "MASTER"]

                    all_file_data.append(df_clean[["DATE_STR", "EMP ID", "STATION", "SN"]])
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการอ่านไฟล์ {uploaded_file.name}: {e}")

    if all_file_data:
        # รวมข้อมูลดิบจากทุกไฟล์เข้าด้วยกัน
        raw_combined_df = pd.concat(all_file_data, ignore_index=True)

        # 🎯 ตัดรายการที่ "หมายเลขสินค้า (SN)" ซ้ำกันออก ให้เหลือชิ้นเดียวจริง
        combined_df = raw_combined_df.drop_duplicates(subset=["SN"], keep="first")

        if not combined_df.empty:
            # คำนวณสรุปยอดแยกตาม: วันที่ -> พนักงาน -> รุ่นสินค้า (STATION) แบบในรูปตัวอย่าง
            daily_summary_df = (
                combined_df.groupby(["DATE_STR", "EMP ID", "STATION"])["SN"]
                .count()
                .reset_index()
            )
            daily_summary_df.columns = [
                "วันที่",
                "รหัสพนักงาน",
                "รุ่นสินค้า (STATION)",
                "จำนวนจริง (ตัว)",
            ]

            # --------------------------------------------------
            # ✨ ส่วนการแสดงผลบนหน้าจอเว็บ
            # --------------------------------------------------
            st.success(f"🎉 รวมข้อมูลสำเร็จทั้งหมด {len(uploaded_files)} ไฟล์!")

            # แสดงลิสต์สรุปรายวันแยกบรรทัดให้อ่านง่าย
            total_sum = daily_summary_df["จำนวนจริง (ตัว)"].sum()
            for _, row in daily_summary_df.iterrows():
                st.markdown(
                    f"📅 วันที่: **{row['วันที่']}** | 👤 พนักงาน: **{row['รหัสพนักงาน']}** | 🛠️ รุ่น: ` {row['รุ่นสินค้า (STATION)']} ` ➡️ ได้จริง **{row['จำนวนจริง (ตัว)']}** ตัว"
                )
            st.markdown(f"📊 **ยอดรวมทั้งหมด (Total): {total_sum:,} ตัว**")

            # --------------------------------------------------
            # 📥 [ฟังก์ชันไฮไลต์] ระบบสร้างไฟล์ Excel แอด Total + ฝังรูปภาพกราฟ
            # --------------------------------------------------
            st.markdown("### 📥 ส่งออกข้อมูลเป็นไฟล์ Excel")

            # 1. เตรียมตารางสำหรับส่งออก (เพิ่มแถว Total ท้ายตารางแบบในรูปเรฟ)
            export_df = daily_summary_df.copy()
            total_row = pd.DataFrame([{
                "วันที่": "Total",
                "รหัสพนักงาน": "",
                "รุ่นสินค้า (STATION)": "",
                "จำนวนจริง (ตัว)": total_sum
            }])
            export_df = pd.concat([export_df, total_row], ignore_index=True)

            # 2. สร้างกราฟเป็นรูปภาพในหน่วยความจำเพื่อนำไปฝังใน Excel
            fig_img_buf = io.BytesIO()
            plt.figure(figsize=(6, 4))
            sns.barplot(data=daily_summary_df, x="รหัสพนักงาน", y="จำนวนจริง (ตัว)", hue="รุ่นสินค้า (STATION)")
            plt.title("Daily Production Summary")
            plt.tight_layout()
            plt.savefig(fig_img_buf, format="png")
            fig_img_buf.seek(0)
            plt.close()

            # 3. ใช้ openpyxl สร้างไฟล์และจัดหน้า
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                export_df.to_excel(writer, index=False, sheet_name="Summary_Report")
                
                # เข้าถึงสไตล์เพื่อตบแต่งหน้าตาและใส่รูปภาพ
                workbook = writer.book
                worksheet = writer.sheets["Summary_Report"]
                
                # ทำการแทรกรูปภาพกราฟลงในเซลล์ F2 (ข้างๆ ตารางสรุป)
                try:
                    img = OpenpyxlImage(fig_img_buf)
                    worksheet.add_image(img, "F2")
                except Exception as img_err:
                    pass # หากระบบฝังรูปภาพติดขัดในบาง Environment ให้ทำงานต่อได้

            excel_buffer.seek(0)

            # ปุ่มดาวน์โหลด Excel
            st.download_button(
                label="🟢 คลิกตรงนี้เพื่อส่งออกไฟล์ Excel (มีตาราง + ยอด Total + รูปกราฟ)",
                data=excel_buffer,
                file_name=f"รายงานสรุปสต็อกจริง_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # 4. แสดงกราฟแบบส่องดูบนเว็บ สวยงามคมชัดเหมือนเดิม
            st.markdown("---")
            fig = px.bar(
                daily_summary_df,
                x="วันที่",
                y="จำนวนจริง (ตัว)",
                color="รุ่นสินค้า (STATION)",
                facet_col="รหัสพนักงาน",
                barmode="group",
                title="กราฟสรุปยอดผลิตจริงแยกตามวัน พนักงาน และรุ่นสินค้า",
            )
            st.plotly_chart(fig, use_container_width=True)

            # ปุ่มกดบันทึกข้อมูลเข้าคลังประวัติสะสม
            if st.button("📥 บันทึกข้อมูลชุดนี้ลงประวัติยอดรวมสะสม"):
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                files_string = ", ".join(file_names_list)

                for _, row in daily_summary_df.iterrows():
                    st.session_state.history_log.append(
                        {
                            "เวลาที่บันทึกระบบ": current_time,
                            "จากไฟล์ทั้งหมด": files_string,
                            "วันที่ทำงาน": row["วันที่"],
                            "รหัสพนักงาน": row["รหัสพนักงาน"],
                            "รุ่นสินค้า (STATION)": row["รุ่นสินค้า (STATION)"],
                            "จำนวนรวมสะสม (ตัว)": row["จำนวนจริง (ตัว)"],
                        }
                    )
                st.toast("บันทึกข้อมูลเข้าคลังประวัติเรียบร้อยแล้ว!")
                st.rerun()
        else:
            st.warning("⚠️ ไม่พบข้อมูลสินค้าอื่นนอกเหนือจากสถานี MASTER เลย")
    else:
        st.error("❌ ไม่พบโครงสร้างข้อมูลที่ถูกต้อง กรุณาตรวจสอบหัวตารางไฟล์ Excel")

# --------------------------------------------------
# ส่วนที่ 2: หน้าต่างประวัติยอดรวมสะสมย้อนหลัง
# --------------------------------------------------
st.markdown("---")
st.subheader("📜 คลังประวัติยอดรวมสะสมย้อนหลัง (เฉพาะตัวงานจริง)")

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
