import os
import fitz
import win32com.client as win32
from pypdf import PdfReader, PdfWriter
import psutil




# 경로 설정
excel_folder = r"C:\Users\올담에듀\Desktop\피드백 파일 생성기\단어 엑셀 input"
output_pdf_dir = r"C:\Users\올담에듀\Desktop\피드백 파일 생성기\단어 엑셀 PDF 파일 output"
output_img_dir = r"C:\Users\올담에듀\Desktop\피드백 파일 생성기\단어 엑셀 이미지 파일 output"

os.makedirs(output_pdf_dir, exist_ok=True)
os.makedirs(output_img_dir, exist_ok=True)


def save_and_quit_excel():
    excel = None
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        return


    if excel is not None:
        for wb in excel.Workbooks:
            try:
                print(f"저장 중: {wb.Name}")
                wb.Save()
            except Exception as e:
                print(f"저장 실패: {wb.Name} - {e}")

        excel.Quit()
        print("엑셀 종료 완료")

save_and_quit_excel()



# 폴더 내 엑셀파일 읽기
excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
pdf_paths = []
file_count = 0

try:
    for file_name in os.listdir(excel_folder):
        if file_name.endswith(".xlsx") and not file_name.startswith("~$"):
            excel_path = os.path.join(excel_folder, file_name)
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            file_count += 1

            for idx, sheet in enumerate(wb.Sheets, start=1):
                try:
                    print(f"[{file_name}] 시트 처리 중: {sheet.Name}")

                    # 인쇄 범위
                    sheet.PageSetup.PrintArea = "A1:AX73" # 변경 원할 시 이거 수정!!!

                    # 인쇄 설정
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = 1

                    sheet.PageSetup.LeftMargin = excel.InchesToPoints(0.1)
                    sheet.PageSetup.RightMargin = excel.InchesToPoints(0.1)
                    sheet.PageSetup.TopMargin = excel.InchesToPoints(0.1)
                    sheet.PageSetup.BottomMargin = excel.InchesToPoints(0.1)

                    sheet.PageSetup.CenterHorizontally = True
                    sheet.PageSetup.CenterVertically = True
                    sheet.PageSetup.Orientation = 1

                    # 파일 저장
                    prefix = f"{file_count}-{idx}. "
                    pdf_filename = f"{prefix}{sheet.Name}.pdf"
                    pdf_path = os.path.join(output_pdf_dir, pdf_filename)
                    sheet.ExportAsFixedFormat(0, pdf_path)
                    pdf_paths.append(pdf_path)

                except Exception as e:
                    print(f"Error exporting {sheet.Name}: {e}")

            wb.Close(SaveChanges=False)

finally:
    excel.Quit()


# PDF - 이미지 변환
for pdf_path in pdf_paths:
    doc = fitz.open(pdf_path)

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    page = doc.load_page(0)
    pix = page.get_pixmap()

    img_name = f"{base_name}.jpg"
    img_path = os.path.join(output_img_dir, img_name)
    pix.save(img_path)

print("모든 시트 PDF 저장 및 이미지 변환 완료.")

# PDF 병합
merged_pdf_path = os.path.join(output_pdf_dir, "최종 병합.pdf")
writer = PdfWriter()

for pdf_path in pdf_paths:
    reader = PdfReader(pdf_path)

    for page in reader.pages:
        writer.add_page(page)

with open(merged_pdf_path, "wb") as f_out:
    writer.write(f_out)

print(f"모든 PDF를 '{merged_pdf_path}'로 병합 완료.")
