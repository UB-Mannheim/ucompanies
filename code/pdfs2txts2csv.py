import pdfplumber
from PIL import Image
import io
import numpy as np
import os
import csv
from tqdm import tqdm
import pandas as pd


def detect_check_mark(image):
    grayscale_image = image.convert("L")
    image_array = np.array(grayscale_image)
    average_pixel_value = np.mean(image_array)
    if average_pixel_value < 170:  # Checked mark will have lower average pixel value (darker)
        return "CHECKED_MARK"  # Checked check mark
    else:
        return "UNCHECKED_MARK"  # Unchecked check mark

def extract_fields_from_pdf(pdf_path):
    results = {
        'Application Number': '',
        'Applicant Name': '',
        'Tax Number or Passport': '',
        'Gender': '',
        'Location': '',
        'Main Economic Activity': '',
        'Company Registration Date': '',
        'Phone': '',
        'Email': '',
        'Website': '',
        'Employees on 31.12.2021': '',
        'Current Employees': '',
        'Actual Business Location': '',
        'Contact Person': '',
        'Received State Relocation Assistance': '',
        'Received Other Assistance': '',
        'Damages - Damage to company premises': 'No',
        'Damages - Damage to production': 'No',
        'Damages - Layoff of employees': 'No',
        'Damages - Problems with operation': 'No',
        'Damages - Relocation of company': 'No',
        'Damages - Shortage of raw materials': 'No',
        'Damages - Logistic problems': 'No',
        'Damages - Contract terminations': 'No',
        'Damages - Other': 'No',
        'Operational Status': '',
        'Grant Usage Plans': '',
        'Grant Amount (UAH)': '',
        'Specific Use of Grant': '',
        'Grant Impact': '',
        'Activity - Production': 'No',
        'Activity - Services': 'No',
        'Activity - Agriculture': 'No',
        'Activity - Trade': 'No',
        'Activity - Other': 'No',
        'Relocation - Already moved to another safer region of Ukraine': 'No',
        'Relocation - Plan to move to another safer region of Ukraine': 'No',
        'Relocation - Plan to move abroad': 'No',
        'Relocation - Plan to return to the previous location if the war ends': 'No',
        'Relocation - Decision not yet made, will monitor the situation and decide later': 'No',
        'Relocation - Other': 'No',
        'Grant Impact - Help restore production/services': 'No',
        'Grant Impact - Increase production capacity': 'No',
        'Grant Impact - Help expand customer base': 'No',
        'Grant Impact - Increase personnel capacity': 'No',
        'Grant Impact - Increase profitability': 'No',
        'Grant Impact - Other': 'No',
        'Date': '',
        'Signature': ''
    }

    damage_fields = {
        "Пошкодження приміщень підприємства": 'Damages - Damage to company premises',
        "Пошкодження виробництва": 'Damages - Damage to production',
        "Звільнення співробітників": 'Damages - Layoff of employees',
        "Проблеми з експлуатацією виробничих потужностей та обладнання": 'Damages - Problems with operation',
        "Переїзд підприємства в інший регіон": 'Damages - Relocation of company',
        "Відсутність або дефіцит сировини для виробництва": 'Damages - Shortage of raw materials',
        "Проблеми з логістикою": 'Damages - Logistic problems',
        "Припинення або розірвання через бойові дії контрактів на поставку товарів, робіт та послуг за кордон": 'Damages - Contract terminations',
        "Інше": 'Damages - Other'
    }

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page_num, page in enumerate(pdf.pages):
            text_blocks = []

            # Extract text lines and bounding boxes
            for word in page.extract_words():
                text = word['text']
                bbox = [word['x0'], word['top'], word['x1'], word['bottom']]
                text_blocks.append({
                    "text": text.strip(),
                    "bbox": bbox  # Store bounding box coordinates
                })

            check_marks = []

            # Extract images and positions (only from the second page onwards)
            if page_num > 0:
                for img in page.images:
                    img_bbox = [img['x0'], img['top'], img['x1'], img['bottom']]
                    img_obj = page.within_bbox(img_bbox).to_image()
                    img_bytes = io.BytesIO()
                    img_obj.save(img_bytes, format='PNG')
                    image = Image.open(img_bytes)

                    check_mark_text = detect_check_mark(image)  # Use the image content to detect check mark

                    check_marks.append((check_mark_text, img_bbox))

            # Place check marks based on bounding boxes
            for check_mark_text, bbox_coords in check_marks:
                check_mark_center_y = (bbox_coords[1] + bbox_coords[3]) / 2
                matched = False
                for block in text_blocks:
                    text_bbox = block["bbox"]
                    if text_bbox[1] <= check_mark_center_y <= text_bbox[3]:
                        block["text"] = check_mark_text + " " + block["text"]
                        block["check_mark"] = check_mark_text
                        matched = True
                        break

            # Sort text blocks by their vertical position and group them by lines
            text_blocks.sort(key=lambda block: block["bbox"][1])
            current_y = None
            line_text = ""
            page_text = ""

            for block in text_blocks:
                text = block["text"]
                y_position = block["bbox"][1]
                if current_y is None or abs(current_y - y_position) > 5:  # New line detected
                    if line_text:
                        page_text += line_text.strip() + "\n"
                    line_text = text + " "
                    current_y = y_position
                else:
                    line_text += text + " "
            if line_text:
                page_text += line_text.strip() + "\n"

            full_text += page_text.strip() + "\n"

        # Extract specific fields from the full text
        lines = full_text.split("\n")
        for i, line in enumerate(lines):
            if 'Заява №' in line:
                results['Application Number'] = line.split("№")[1].strip()
            elif 'Прізвище, ім’я, по батькові' in line or 'Найменування юридичної особи' in line:
                results['Applicant Name'] = lines[i+1].strip()
            elif 'Реєстраційний номер' in line or 'Ідентифікаційний код' in line:
                results['Tax Number or Passport'] = lines[i+1].strip()
            elif 'Відомості щодо статі' in line or 'Відомості щодо статі керівника' in line:
                results['Gender'] = lines[i+1].strip()
            elif 'Місцезнаходження' in line:
                results['Location'] = " ".join(lines[i+1:i+4]).strip()
            elif 'Основний вид економічної діяльності' in line:
                results['Main Economic Activity'] = lines[i+1].strip()
            elif 'Дата реєстрації компанії' in line:
                results['Company Registration Date'] = lines[i+1].strip()
            elif 'Телефон' in line:
                results['Phone'] = lines[i+1].strip()
            elif 'Адреса електронної пошти' in line:
                results['Email'] = lines[i+1].strip()
            elif 'Вебсайт' in line:
                results['Website'] = lines[i+1].strip()
            elif 'Кількість найманих працівників' in line:
                if i + 1 < len(lines):
                    employees_line = lines[i + 2]
                    employee_numbers = employees_line.split()
                    if len(employee_numbers) >= 2:
                        results['Employees on 31.12.2021'] = employee_numbers[0].strip()
                        results['Current Employees'] = employee_numbers[1].strip()
            elif 'Фактичне місце провадження господарської діяльності' in line:
                results['Actual Business Location'] = " ".join(lines[i+1:i+4]).strip()
            elif 'Контактна особа' in line:
                results['Contact Person'] = lines[i+1].strip()
            elif 'Відомості щодо отримання державної допомоги' in line:
                results['Received State Relocation Assistance'] = lines[i+2].strip()
            elif 'Відомості щодо отримання будь-якої іншої державної допомоги' in line:
                results['Received Other Assistance'] = lines[i+2].strip()
            elif 'Збитки, які зазнали у зв’язку із веденням бойових дій' in line:
                for j in range(i + 1, i + 10):
                    if j < len(lines):
                        for damage_text, damage_field in damage_fields.items():
                            if damage_text in lines[j]:
                                if "UNCHECKED_MARK" in lines[j]:
                                    results[damage_field] = 'No'
                                elif "CHECKED_MARK" in lines[j]:
                                    results[damage_field] = 'Yes'
            elif 'Інформація стосовно фактичного ведення господарської діяльності' in line:
                results['Operational Status'] = lines[i+1].strip()
            elif 'Потреби та плани щодо використання допомоги' in line:
                results['Grant Usage Plans'] = lines[i+1].strip()
            elif 'Сума, яку плануєте витратити на це, грн' in line:
                results['Grant Amount (UAH)'] = lines[i+1].strip()
            elif 'На що саме ви плануєте витратити цю суму' in line:
                results['Specific Use of Grant'] = lines[i+1].strip()
            elif 'Ефект або вплив від отриманої допомоги' in line:
                results['Grant Impact'] = " ".join(lines[i+1:i+5]).strip()
            elif 'Сфера діяльності' in line:
                results['Activity - Production'] = 'No' if "UNCHECKED_MARK Виробництво" in full_text else 'Yes'
                results['Activity - Services'] = 'No' if "UNCHECKED_MARK Надання послуг" in full_text else 'Yes'
                results['Activity - Agriculture'] = 'No' if "UNCHECKED_MARK Сільське господарство" in full_text else 'Yes'
                results['Activity - Trade'] = 'No' if "UNCHECKED_MARK Торгівля" in full_text else 'Yes'
                results['Activity - Other'] = 'No' if "UNCHECKED_MARK Інше" in full_text else 'Yes'
            elif 'Наміри щодо розміщення підприємства' in line:
                results['Relocation - Already moved to another safer region of Ukraine'] = 'No' if "UNCHECKED_MARK Уже переїхав в інший більш безпечний регіон України" in full_text else 'Yes'
                results['Relocation - Plan to move to another safer region of Ukraine'] = 'No' if "UNCHECKED_MARK Планую переїхати в інший більш безпечний регіон України" in full_text else 'Yes'
                results['Relocation - Plan to move abroad'] = 'No' if "UNCHECKED_MARK Планую перемістити підприємство за кордон" in full_text else 'Yes'
                results['Relocation - Plan to return to the previous location if the war ends'] = 'No' if "UNCHECKED_MARK Планую повернутися на попереднє місце перебування, якщо закінчиться війна" in full_text else 'Yes'
                results['Relocation - Decision not yet made, will monitor the situation and decide later'] = 'No' if "UNCHECKED_MARK Рішення ще не прийнято, буду стежити за розвитком ситуації, а потім вирішу" in full_text else 'Yes'
                results['Relocation - Other'] = 'No' if "UNCHECKED_MARK Інше" in full_text else 'Yes'
            elif 'Вплив гранту на бізнес протягом наступних 6 місяців' in line:
                results['Grant Impact - Help restore production/services'] = 'No' if "UNCHECKED_MARK Допоможе відновити виробництво/надання послуг" in full_text else 'Yes'
                results['Grant Impact - Increase production capacity'] = 'No' if "UNCHECKED_MARK Збільшить виробничі потужності" in full_text else 'Yes'
                results['Grant Impact - Help expand customer base'] = 'No' if "UNCHECKED_MARK Допоможе розширенню клієнтської бази" in full_text else 'Yes'
                results['Grant Impact - Increase personnel capacity'] = 'No' if "UNCHECKED_MARK Збільшить кадровий потенціал" in full_text else 'Yes'
                results['Grant Impact - Increase profitability'] = 'No' if "UNCHECKED_MARK Збільшить прибутковість" in full_text else 'Yes'
                results['Grant Impact - Other'] = 'No' if "UNCHECKED_MARK Інше" in full_text else 'Yes'
            elif 'Інше' in line:
                if i + 1 < len(lines):
                    signature_line = lines[i + 1].strip()
                    if ' ' in signature_line:
                        results['Date'] = signature_line.split()[0].strip()
                        results['Signature'] = " ".join(signature_line.split()[1:]).strip()

    return results

def process_pdfs_in_directory(directory_path, output_csv):
    all_results = []
    fields = [
        'Application Number', 'Applicant Name', 'Tax Number or Passport', 'Gender', 'Location',
        'Main Economic Activity', 'Company Registration Date', 'Phone', 'Email', 'Website',
        'Employees on 31.12.2021', 'Current Employees', 'Actual Business Location', 'Contact Person',
        'Received State Relocation Assistance', 'Received Other Assistance',
        'Damages - Damage to company premises', 'Damages - Damage to production',
        'Damages - Layoff of employees', 'Damages - Problems with operation',
        'Damages - Relocation of company', 'Damages - Shortage of raw materials',
        'Damages - Logistic problems', 'Damages - Contract terminations', 'Damages - Other',
        'Operational Status', 'Grant Usage Plans', 'Grant Amount (UAH)', 'Specific Use of Grant', 'Grant Impact',
        'Activity - Production', 'Activity - Services', 'Activity - Agriculture', 'Activity - Trade', 'Activity - Other',
        'Relocation - Already moved to another safer region of Ukraine', 'Relocation - Plan to move to another safer region of Ukraine',
        'Relocation - Plan to move abroad', 'Relocation - Plan to return to the previous location if the war ends',
        'Relocation - Decision not yet made, will monitor the situation and decide later', 'Relocation - Other',
        'Grant Impact - Help restore production/services', 'Grant Impact - Increase production capacity',
        'Grant Impact - Help expand customer base', 'Grant Impact - Increase personnel capacity',
        'Grant Impact - Increase profitability', 'Grant Impact - Other', 'Date', 'Signature'
    ]

    for i, filename in enumerate(tqdm(os.listdir(directory_path))):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory_path, filename)
            result = extract_fields_from_pdf(pdf_path)
            all_results.append(result)

    # Write results to CSV
    with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        for row in all_results:
            writer.writerow(row)
    
    df = pd.DataFrame(all_results, columns=fields)
    # Ensure column names are compliant with Stata's requirements
    df.columns = [col.replace(" ", "_").replace("-", "_").replace(".", "").replace("/", "_")[:32] for col in df.columns]
    # Convert all string columns to be compatible with latin-1 encoding
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(lambda x: ''.join([i if ord(i) < 256 else '?' for i in x]) if isinstance(x, str) else x)
    # Save to Stata format with utf-8 encoding
    df.to_stata('../data/extracted_structured_data/328_applforms.dta')

def main():
    pdf_directory = "../data/raw_pdfs/"  # Update this path to your directory containing PDF files
    output_csv = "../data/extracted_structured_data/328_applforms.csv"  # Update this path to your desired output CSV file
    process_pdfs_in_directory(pdf_directory, output_csv)
    print(f"Results written to {output_csv}")

if __name__ == "__main__":
    main()
