# -*- coding: utf-8 -*-
"""
将 PPTX 转换为 PDF（通过 PowerPoint COM 接口）
需要系统已安装 Microsoft Office PowerPoint
"""
import os
import sys
import time

# 文件名映射：原文件名 -> 输出PDF名（英文名，避免编码问题）
MAPPING = {
    "12.12市场秩序-水电气收费监管解决方案V5【待评审】.pptx": "01_water_gas_electricity.pdf",
    "【新】智慧行政审批标准化解决方案.pptx": "02_smart_admin_approval.pdf",
    "【新】非现场监管-食品全产业链溯源监管解决方案1.pptx": "03_food_traceability.pdf",
    "【新】非现场监管-食品安全互联网+AI监管解决方案(地市版).pptx": "04_food_ai_safety.pdf",
    "【新】非现场监管——成品油综合监管解决方案（v2）.pptx": "05_refined_oil.pdf",
    "市场监管行政执法解决方案-0315 V6.pptx": "06_administrative_enforcement.pdf",
    "海南市场监管局方案_海南省工业产品质量安全监管信息系统建设方案 (备注).pptx": "07_industrial_product_quality.pdf",
    "网络直播监测解决方案7.23.pdf": "08_livestream_monitoring.pdf",
    "药品监管-智慧药检解决方案V1.pptx": "09_drug_inspection.pdf",
    "质量监管\u201c一站式\u201d服务平台解决方案(培训)-0702-中移集成解决方案部市场监管行业.pdf": "10_nqi_one_stop.pdf",
}

SRC_DIR = r"C:\Users\maqi\Desktop\方案"
DST_DIR = r"c:\Users\maqi\Desktop\trae\word备注解答工具设计\pdfs"


def convert_pptx_to_pdf():
    try:
        import comtypes
        import comtypes.client
    except ImportError:
        print(" installing comtypes...")
        os.system(f"{sys.executable} -m pip install comtypes -q")
        import comtypes
        import comtypes.client

    os.makedirs(DST_DIR, exist_ok=True)

    print(" starting PowerPoint...")
    comtypes.CoInitialize()
    try:
        ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
        # PowerPoint 必须可见才能转 PDF（Office 安全限制），设为最小化
        try:
            ppt_app.WindowState = 2  # ppWindowMinimized
        except Exception:
            pass
        try:
            ppt_app.Visible = 1
        except Exception:
            pass

        # ppFixedFormatTypePDF = 2
        # ppFixedFormatIntentPrint = 2
        for src_name, dst_name in MAPPING.items():
            src_path = os.path.join(SRC_DIR, src_name)
            dst_path = os.path.join(DST_DIR, dst_name)

            if not os.path.exists(src_path):
                print(f"  [MISS] source not found: {src_name}")
                continue

            if os.path.exists(dst_path):
                print(f"  [SKIP] {dst_name} (already exists)")
                continue

            if src_name.lower().endswith(".pdf"):
                # 已经是PDF，直接复制
                import shutil
                shutil.copy2(src_path, dst_path)
                print(f"  [COPY] {dst_name}")
                continue

            print(f"  [CONVERT] {src_name} -> {dst_name}")
            try:
                # 打开演示文稿（WithWindow=False）
                pres = ppt_app.Presentations.Open(src_path, ReadOnly=True, Untitled=False, WithWindow=False)
                # 保存为PDF
                pres.SaveAs(dst_path, 32)  # ppSaveAsPDF = 32
                pres.Close()
                print(f"    [OK]")
            except Exception as e:
                print(f"    [ERROR] {e}")
                # 如果失败，尝试用ExportAsFixedFormat
                try:
                    pres = ppt_app.Presentations.Open(src_path, ReadOnly=True, Untitled=False, WithWindow=False)
                    pres.ExportAsFixedFormat(dst_path, 2)  # ppFixedFormatTypePDF=2
                    pres.Close()
                    print(f"    [OK via ExportAsFixedFormat]")
                except Exception as e2:
                    print(f"    [ERROR2] {e2}")

        try:
            ppt_app.Quit()
        except Exception:
            pass
        print(" ALL_DONE")
    finally:
        comtypes.CoUninitialize()


if __name__ == "__main__":
    convert_pptx_to_pdf()
