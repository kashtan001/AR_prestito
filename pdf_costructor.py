#!/usr/bin/env python3
"""
PDF Constructor API для генерации документов Intesa Sanpaolo
Поддерживает: contratto, garanzia, carta, compensazione
"""

from io import BytesIO
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

# Импортируем типы для аннотаций
from typing import List, Dict, Any


def format_money(amount: float) -> str:
    """Форматирование суммы БЕЗ знака € (он уже есть в HTML)"""
    return f"{amount:,.2f}".replace(',', ' ')


def format_date() -> str:
    """Получение текущей даты в итальянском формате"""
    return datetime.now().strftime("%d/%m/%Y")


def format_money_it(amount: float) -> str:
    """Сумма для итальянского текста: 10 000,00 (без € в шаблоне)."""
    return f"{amount:,.2f}".replace(',', ' ').replace('.', ',')


def monthly_payment(amount: float, months: int, annual_rate: float) -> float:
    """Аннуитетный расчёт ежемесячного платежа"""
    r = (annual_rate / 100) / 12
    if r == 0:
        return round(amount / months, 2)
    num = amount * r * (1 + r) ** months
    den = (1 + r) ** months - 1
    return round(num / den, 2)


def calculate_amortization_schedule(amount: float, months: int, rate: float) -> List[Dict[str, Any]]:
    """Расчет графика погашения"""
    schedule = []
    balance = amount
    monthly_rate = (rate / 100) / 12
    payment = monthly_payment(amount, months, rate)
    
    for i in range(1, months + 1):
        interest = balance * monthly_rate
        principal = payment - interest
        
        # Корректировка последнего платежа
        if i == months:
            principal = balance
            payment = principal + interest
        
        balance -= principal
        if balance < 0.01: balance = 0
        
        schedule.append({
            'month': i,
            'payment': payment,
            'interest': interest,
            'principal': principal,
            'balance': balance
        })
    return schedule


def generate_amortization_html(schedule: List[Dict[str, Any]]) -> str:
    """Генерация HTML таблицы амортизации"""
    rows = ""
    for item in schedule:
        rows += f"""
<tr class="c7">
<td class="c5" style="border: 1pt solid #666666; padding: 3pt; text-align: center;"><span class="c3">{item['month']}</span></td>
<td class="c5" style="border: 1pt solid #666666; padding: 3pt; text-align: right;"><span class="c9 c8">&euro; {format_money(item['payment'])}</span></td>
<td class="c5" style="border: 1pt solid #666666; padding: 3pt; text-align: right;"><span class="c9 c8">&euro; {format_money(item['interest'])}</span></td>
<td class="c5" style="border: 1pt solid #666666; padding: 3pt; text-align: right;"><span class="c9 c8">&euro; {format_money(item['principal'])}</span></td>
<td class="c5" style="border: 1pt solid #666666; padding: 3pt; text-align: right;"><span class="c9 c8">&euro; {format_money(item['balance'] if item['balance'] > 0 else 0)}</span></td>
</tr>
"""
    
    table = f"""
<table class="c18" style="width: 100%; border-collapse: collapse; margin: 10pt 0;">
<tr class="c7">
<td class="c4" style="border: 1pt solid #666666; padding: 5pt; text-align: center; font-weight: 700;"><span class="c3">Mese</span></td>
<td class="c4" style="border: 1pt solid #666666; padding: 5pt; text-align: center; font-weight: 700;"><span class="c3">Rata</span></td>
<td class="c4" style="border: 1pt solid #666666; padding: 5pt; text-align: center; font-weight: 700;"><span class="c3">Interessi</span></td>
<td class="c4" style="border: 1pt solid #666666; padding: 5pt; text-align: center; font-weight: 700;"><span class="c3">Capitale</span></td>
<td class="c4" style="border: 1pt solid #666666; padding: 5pt; text-align: center; font-weight: 700;"><span class="c3">Residuo</span></td>
</tr>
{rows}
</table>
"""
    return table


def generate_signatures_table() -> str:
    """
    Генерирует две наложенные друг на друга таблицы:
    1) Таблица с печатью (seal.png) в первой колонке
    2) Таблица с подписями (sing_1.png и sing_2.png) - сдвинута вправо
    Изображения встраиваются как base64 для гарантированной загрузки в weasyprint.
    """
    import os
    import base64

    base_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()

    def image_to_base64(filename: str) -> str | None:
        img_path = os.path.join(base_dir, filename)
        if os.path.exists(img_path):
            with open(img_path, 'rb') as f:
                img_base64 = base64.b64encode(f.read()).decode('utf-8')
            mime_type = 'image/png' if filename.lower().endswith('.png') else 'image/jpeg'
            return f"data:{mime_type};base64,{img_base64}"
        return None

    sing_1_data = image_to_base64('sing_1.png')
    sing_2_data = image_to_base64('sing_2.png')
    seal_data = image_to_base64('seal.png')

    if not all([sing_1_data, sing_2_data, seal_data]):
        print("⚠️  Не все изображения найдены для таблицы подписей/печати (sing_1.png, sing_2.png, seal.png)")
        return ''

    # Таблица 1: Печать в 1-й колонке (нижний слой)
    seal_table = f'''
<table class="signatures-table-base">
<tr>
<td style="width: 33.33%;">
<img src="{seal_data}" alt="Sigillo Mediatore" class="seal-img" style="display: block; margin: 0 auto;" />
</td>
<td style="width: 33.33%;"></td>
<td style="width: 33.33%;"></td>
</tr>
</table>
'''

    # Таблица 2: Подписи в 1-й и 2-й колонках (верхний слой, сдвинута вправо)
    signatures_table = f'''
<table class="signatures-table-overlay">
<tr>
<td style="width: 33.33%;">
<img src="{sing_2_data}" alt="Firma Banca" class="sing-img" style="display: block; margin: 0 auto;" />
</td>
<td style="width: 33.33%;">
<img src="{sing_1_data}" alt="Firma Mediatore" class="sing-img" style="display: block; margin: 0 auto;" />
</td>
<td style="width: 33.33%;"></td>
</tr>
</table>
'''

    return f'''
<div class="signatures-tables-wrapper">
{seal_table}
{signatures_table}
</div>
'''


def generate_contratto_pdf(data: dict) -> BytesIO:
    """
    API функция для генерации PDF договора
    """
    # Рассчитываем платеж если не задан
    if 'payment' not in data:
        data['payment'] = monthly_payment(data['amount'], data['duration'], data['tan'])
    
    # Генерируем таблицу амортизации
    schedule = calculate_amortization_schedule(data['amount'], data['duration'], data['tan'])
    data['amortization_table'] = generate_amortization_html(schedule)
    
    html = fix_html_layout('contratto')
    return _generate_pdf_with_images(html, 'contratto', data)


def generate_garanzia_pdf(name: str) -> BytesIO:
    """
    API функция для генерации PDF гарантийного письма
    """
    html = fix_html_layout('garanzia')
    return _generate_pdf_with_images(html, 'garanzia', {'name': name})


def generate_carta_pdf(data: dict) -> BytesIO:
    """
    API функция для генерации PDF письма о карте
    """
    # Рассчитываем платеж если не задан
    if 'payment' not in data:
        data['payment'] = monthly_payment(data['amount'], data['duration'], data['tan'])
    
    html = fix_html_layout('carta')
    return _generate_pdf_with_images(html, 'carta', data)


def generate_approvazione_pdf(data: dict) -> BytesIO:
    """
    API функция для генерации PDF письма об одобрении кредита
    """
    html = fix_html_layout('approvazione')
    return _generate_pdf_with_images(html, 'approvazione', data)


def generate_compensazione_pdf(data: dict) -> BytesIO:
    html = fix_html_layout('compensazione')
    return _generate_pdf_with_images(html, 'compensazione', data)


def _generate_pdf_with_images(html: str, template_name: str, data: dict) -> BytesIO:
    """Внутренняя функция для генерации PDF с изображениями"""
    try:
        from weasyprint import HTML
        
        # Заменяем XXX на реальные данные
        if template_name in ['contratto', 'carta', 'garanzia', 'approvazione', 'compensazione']:
            replacements = []
            if template_name == 'contratto':
                # Calculate amortization summary
                monthly_rate = (data['tan'] / 100) / 12
                total_payments = data['payment'] * data['duration']
                overpayment = total_payments - data['amount']

                replacements = [
                    ('XXX', data['name']),  # имя клиента (первое)
                    ('XXX', format_money(data['amount'])),  # сумма кредита
                    ('XXX', f"{data['tan']:.2f}%"),  # TAN
                    ('XXX', f"{data['taeg']:.2f}%"),  # TAEG  
                    ('XXX', f"{data['duration']} mesi"),  # срок
                    ('XXX', format_money(data['payment'])),  # платеж
                    ('11/06/2025', format_date()),  # дата
                    ('XXX', data['name']),  # имя в подписи
                ]

                # Вставляем таблицу амортизации
                html = html.replace('{{AMORTIZATION_TABLE}}', data.get('amortization_table', ''))
                
                # Вставляем данные для пункта 6
                html = html.replace('PAYMENT_SCHEDULE_MONTHLY_RATE', f"{monthly_rate:.10f}")
                html = html.replace('PAYMENT_SCHEDULE_MONTHLY_PAYMENT', f"&euro; {format_money(data['payment'])}")
                html = html.replace('PAYMENT_SCHEDULE_TOTAL_PAYMENTS', f"&euro; {format_money(total_payments)}")
                html = html.replace('PAYMENT_SCHEDULE_OVERPAYMENT', f"&euro; {format_money(overpayment)}")

                # Добавляем класс к разделу 7 для принудительного разрыва страницы
                import re
                html = re.sub(
                    r'(<div style="page-break-before: always;"></div>)',
                    r'',
                    html
                )
                # Ищем параграф с "7. Firme" и добавляем разрыв страницы ПЕРЕД пунктирной линией
                html = re.sub(
                    r'(<p class="c2">\s*<span class="c1">-{10,}</span>\s*</p>)(\s*<p class="c2">\s*<span class="c12 c6">7\. Firme</span>\s*</p>)',
                    r'<p class="c2 section-7-firme"><span class="c1">------------------------------------------</span></p>\2',
                    html
                )
                print("✅ Раздел 7 'Firme' будет начинаться с новой страницы")

                # Таблица с подписями и печатью, вставляем после 7-го пункта
                signatures_table = generate_signatures_table()
                html = html.replace('<!-- SIGNATURES_TABLE_PLACEHOLDER -->', signatures_table)
                print("💉 Изображения подписей внедрены через signatures_table")
                
                # ПОСЛЕ вставки таблицы - заменяем XXX на данные
                for old, new in replacements:
                    html = html.replace(old, new, 1)  # заменяем по одному
                
                # Уже обработали contratto, пропускаем общий блок замен
                replacements = []
            elif template_name == 'carta':
                replacements = [
                    ('XXX', data['name']),
                    ('XXX', format_money(data['amount'])),
                    ('XXX', f"{data['tan']:.2f}%"),
                    ('XXX', f"{data['duration']} mesi"),
                    ('XXX', format_money(data['payment'])),
                ]
            elif template_name == 'garanzia':
                replacements = [
                    ('XXX', data['name']),
                ]
            elif template_name == 'approvazione':
                replacements = [
                    ('XXX', data['name']),
                    ('XXX', format_money(data['amount'])),
                    ('XXX', f"{data['tan']:.2f}%"),
                ]
            elif template_name == 'compensazione':
                nm = data['name'].strip()
                name_display = nm if nm.endswith(',') else nm + ','
                replacements = [
                    ('XXX', format_date()),
                    ('XXX', name_display),
                    ('XXX', format_money_it(data['commission'])),
                    ('XXX', format_money_it(data['indemnity'])),
                ]
            
            for old, new in replacements:
                html = html.replace(old, new, 1)
        
        # Конвертируем HTML в PDF
        pdf_bytes = HTML(string=html, base_url='.').write_pdf()
        
        # НАКЛАДЫВАЕМ ИЗОБРАЖЕНИЯ ЧЕРЕЗ REPORTLAB
        return _add_images_to_pdf(pdf_bytes, template_name)
            
    except Exception as e:
        print(f"Ошибка генерации PDF: {e}")
        raise


def _add_images_to_pdf(pdf_bytes: bytes, template_name: str) -> BytesIO:
    """Добавляет изображения на PDF через ReportLab"""
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from PyPDF2 import PdfReader, PdfWriter
        from PIL import Image
        
        # Создаем overlay с изображениями
        overlay_buffer = BytesIO()
        overlay_canvas = canvas.Canvas(overlay_buffer, pagesize=A4)
        
        # Размер ячейки для расчета сдвигов
        cell_width_mm = 210/25  # 8.4mm
        cell_height_mm = 297/35  # 8.49mm
        
        if template_name == 'garanzia':
            # Логика для garanzia (без изменений)
            company_img = Image.open("company.png")
            company_width_mm = company_img.width * 0.264583
            company_height_mm = company_img.height * 0.264583
            
            company_scaled_width = company_width_mm / 1.33
            company_scaled_height = company_height_mm / 1.33
            
            row_27 = (27 - 1) // 25
            col_27 = (27 - 1) % 25
            
            x_27_center = (col_27 + 5 + 0.5 + 1.25 + 1 + 1/3 - 3 - 0.5) * cell_width_mm * mm
            y_27_center = (297 - (row_27 + 0.5 + 1 + 0.5) * cell_height_mm) * mm
            
            x_27 = x_27_center - (company_scaled_width * mm / 2)
            y_27 = y_27_center - (company_scaled_height * mm / 2)
            
            overlay_canvas.drawImage("company.png", x_27, y_27, 
                                   width=company_scaled_width*mm, height=company_scaled_height*mm,
                                   mask='auto', preserveAspectRatio=True)
            
            seal_img = Image.open("seal.png")
            seal_width_mm = seal_img.width * 0.264583
            seal_height_mm = seal_img.height * 0.264583
            
            seal_scaled_width = seal_width_mm / 5
            seal_scaled_height = seal_height_mm / 5
            
            row_590 = (590 - 1) // 25
            col_590 = (590 - 1) % 25
            
            x_590_center = (col_590 + 0.5) * cell_width_mm * mm
            y_590_center = (297 - (row_590 + 0.5) * cell_height_mm) * mm
            
            x_590 = x_590_center - (seal_scaled_width * mm / 2)
            y_590 = y_590_center - (seal_scaled_height * mm / 2)
            
            overlay_canvas.drawImage("seal.png", x_590, y_590, 
                                   width=seal_scaled_width*mm, height=seal_scaled_height*mm,
                                   mask='auto', preserveAspectRatio=True)
            
            sing1_img = Image.open("sing_1.png")
            sing1_width_mm = sing1_img.width * 0.264583
            sing1_height_mm = sing1_img.height * 0.264583
            
            sing1_scaled_width = sing1_width_mm / 5
            sing1_scaled_height = sing1_height_mm / 5
            
            row_593 = (593 - 1) // 25
            col_593 = (593 - 1) % 25
            
            x_593_center = (col_593 + 0.5) * cell_width_mm * mm
            y_593_center = (297 - (row_593 + 0.5) * cell_height_mm) * mm
            
            x_593 = x_593_center - (sing1_scaled_width * mm / 2)
            y_593 = y_593_center - (sing1_scaled_height * mm / 2)
            
            overlay_canvas.drawImage("sing_1.png", x_593, y_593, 
                                   width=sing1_scaled_width*mm, height=sing1_scaled_height*mm,
                                   mask='auto', preserveAspectRatio=True)
            
            overlay_canvas.save()
            print("🖼️ Добавлены изображения для garanzia через ReportLab API")
        
        elif template_name in ['carta', 'approvazione', 'compensazione']:
            # Логика для carta, approvazione и compensazione (как carta)
            img = Image.open("company.png")
            img_width_mm = img.width * 0.264583
            img_height_mm = img.height * 0.264583
            
            scaled_width = (img_width_mm / 2) * 1.44
            scaled_height = (img_height_mm / 2) * 1.44
            
            row_52 = (52 - 1) // 25 + 1
            col_52 = (52 - 1) % 25 + 1
            
            x_52 = (col_52 * cell_width_mm - 0.5 * cell_width_mm - (1/6) * cell_width_mm + 0.25 * cell_width_mm) * mm
            y_52 = (297 - (row_52 * cell_height_mm + cell_height_mm) + 0.5 * cell_height_mm + 0.25 * cell_height_mm - 1 * cell_height_mm) * mm
            
            overlay_canvas.drawImage("company.png", x_52, y_52, 
                                   width=scaled_width*mm, height=scaled_height*mm, 
                                   mask='auto', preserveAspectRatio=True)
            
            # Печать и подпись: для carta — печать ещё на 3 клетки вниз (665→740), подпись 768
            seal_cell = 740 if template_name == 'carta' else 590
            sign_cell = 768 if template_name == 'carta' else 593
            if template_name == 'compensazione':
                seal_cell, sign_cell = 740, 768
            
            seal_img = Image.open("seal.png")
            seal_width_mm = seal_img.width * 0.264583
            seal_height_mm = seal_img.height * 0.264583
            
            seal_scaled_width = seal_width_mm / 5
            seal_scaled_height = seal_height_mm / 5
            
            row_seal = (seal_cell - 1) // 25
            col_seal = (seal_cell - 1) % 25
            
            x_seal_center = (col_seal + 0.5) * cell_width_mm * mm
            y_seal_center = (297 - (row_seal + 0.5) * cell_height_mm) * mm
            
            x_seal = x_seal_center - (seal_scaled_width * mm / 2)
            y_seal = y_seal_center - (seal_scaled_height * mm / 2)
            
            overlay_canvas.drawImage("seal.png", x_seal, y_seal, 
                                   width=seal_scaled_width*mm, height=seal_scaled_height*mm,
                                   mask='auto', preserveAspectRatio=True)
            
            # Подпись (для carta уже со сдвигом на 3 клетки вниз через sign_cell)
            sing1_img = Image.open("sing_1.png")
            sing1_width_mm = sing1_img.width * 0.264583
            sing1_height_mm = sing1_img.height * 0.264583
            
            sing1_scaled_width = sing1_width_mm / 5
            sing1_scaled_height = sing1_height_mm / 5
            
            row_sign = (sign_cell - 1) // 25
            col_sign = (sign_cell - 1) % 25
            
            x_sign_center = (col_sign + 0.5) * cell_width_mm * mm
            y_sign_center = (297 - (row_sign + 0.5) * cell_height_mm) * mm
            
            x_sign = x_sign_center - (sing1_scaled_width * mm / 2)
            y_sign = y_sign_center - (sing1_scaled_height * mm / 2)
            
            overlay_canvas.drawImage("sing_1.png", x_sign, y_sign, 
                                   width=sing1_scaled_width*mm, height=sing1_scaled_height*mm,
                                   mask='auto', preserveAspectRatio=True)
            
            overlay_canvas.save()
            msg = "печать и подпись на 3 клетки вниз" if template_name == 'carta' else "изображения"
            print(f"🖼️ Добавлены изображения для {template_name} через ReportLab API ({msg})")
        
        elif template_name == 'contratto':
            # Страница 1 - добавляем company.png
            img = Image.open("company.png")
            img_width_mm = img.width * 0.264583
            img_height_mm = img.height * 0.264583
            
            scaled_width = (img_width_mm / 2) * 1.44
            scaled_height = (img_height_mm / 2) * 1.44
            
            row_52 = (52 - 1) // 25 + 1
            col_52 = (52 - 1) % 25 + 1
            
            x_52 = (col_52 * cell_width_mm - 0.5 * cell_width_mm - (1/6) * cell_width_mm + 0.25 * cell_width_mm) * mm
            y_52 = (297 - (row_52 * cell_height_mm + cell_height_mm) + 0.5 * cell_height_mm + 0.25 * cell_height_mm - 1 * cell_height_mm) * mm
            
            overlay_canvas.drawImage("company.png", x_52, y_52, 
                                   width=scaled_width*mm, height=scaled_height*mm, 
                                   mask='auto', preserveAspectRatio=True)
            
            # Добавляем logo.png на страницу 1
            logo_img = Image.open("logo.png")
            logo_width_mm = logo_img.width * 0.264583
            logo_height_mm = logo_img.height * 0.264583
            
            logo_scaled_width = logo_width_mm / 9
            logo_scaled_height = logo_height_mm / 9
            
            row_71 = (71 - 1) // 25
            col_71 = (71 - 1) % 25
            
            x_71 = (col_71 - 2 + 4 - 1.5) * cell_width_mm * mm
            y_71 = (297 - (row_71 * cell_height_mm + cell_height_mm) - 0.25 * cell_height_mm - 1 * cell_height_mm) * mm
            
            overlay_canvas.drawImage("logo.png", x_71, y_71, 
                                   width=logo_scaled_width*mm, height=logo_scaled_height*mm,
                                   mask='auto', preserveAspectRatio=True)
            
            overlay_canvas.showPage()
            
            # Страница 2 и далее - добавляем только logo.png (подписи и печать теперь в HTML)
            overlay_canvas.drawImage("logo.png", x_71, y_71, 
                                   width=logo_scaled_width*mm, height=logo_scaled_height*mm,
                                   mask='auto', preserveAspectRatio=True)
                
            overlay_canvas.save()
            print("🖼️ Добавлены логотипы для contratto через ReportLab API")
            print("💉 Подписи и печать теперь в HTML-таблице, не через ReportLab")
        
        # Объединяем PDF с overlay
        overlay_buffer.seek(0)
        base_pdf = PdfReader(BytesIO(pdf_bytes))
        overlay_pdf = PdfReader(overlay_buffer)
        
        writer = PdfWriter()
        
        # carta/approvazione — только одна страница; лишние от WeasyPrint отбрасываем
        pages_to_use = base_pdf.pages[:1] if template_name in ['carta', 'approvazione', 'compensazione'] else base_pdf.pages
        
        # Накладываем изображения на каждую страницу
        for i, page in enumerate(pages_to_use):
            if i < len(overlay_pdf.pages):
                page.merge_page(overlay_pdf.pages[i])
            writer.add_page(page)
        
        final_buffer = BytesIO()
        writer.write(final_buffer)
        final_buffer.seek(0)
        
        print(f"✅ PDF с изображениями создан через API! Размер: {len(final_buffer.getvalue())} байт")
        return final_buffer
        
    except Exception as e:
        print(f"❌ Ошибка наложения изображений через API: {e}")
        buf = BytesIO(pdf_bytes)
        buf.seek(0)
        return buf


def fix_html_layout(template_name='contratto'):
    """Исправляем HTML для корректного отображения"""
    
    html_file = f'{template_name}.html'
    with open(html_file, 'r', encoding='utf-8') as f:
        html = f.read()
    
    if template_name == 'garanzia':
        # Логика для garanzia (без изменений)
        import re
        html = re.sub(r'<img[^>]*>', '', html)
        html = re.sub(r'<span[^>]*overflow:[^>]*>[^<]*</span>', '<br><br>', html)
        
        css_fixes = """
    <style>
    @page {
        size: A4;
        margin: 1cm;
        border: 4pt solid #a52b4c;
        padding: 0;
    }
    .c8 { padding: 0 2cm !important; max-width: none !important; }
    * { page-break-after: avoid !important; page-break-inside: avoid !important; page-break-before: avoid !important; }
    @page:nth(2) { display: none !important; }
    </style>
    """
        html = html.replace('</head>', f'{css_fixes}</head>')
        return html
    
    elif template_name in ['carta', 'approvazione', 'compensazione']:
        # Логика для carta: убираем фиксированную высоту строки .c10 (789.9pt),
        # иначе контент не помещается на A4 и WeasyPrint создаёт лишнюю 2-ю страницу
        css_fixes = """
    <style>
    @page {
        size: A4;
        margin: 1cm;
        border: 2pt solid #a52b4c;
        padding: 0;
    }
    body { font-family: "Roboto Mono", monospace; font-size: 9pt; line-height: 1.0; margin: 0; padding: 0 2cm; overflow: hidden; }
    * { page-break-after: avoid !important; page-break-inside: avoid !important; page-break-before: avoid !important; overflow: hidden !important; }
    .c10 { height: auto !important; }
    .c12, .c9, .c20, .c22, .c8 { border: none !important; padding: 2pt !important; margin: 0 !important; width: 100% !important; max-width: none !important; }
    .c12 { max-width: none !important; padding: 0 !important; margin: 0 !important; width: 100% !important; height: auto !important; overflow: hidden !important; border: none !important; }
    .c6, .c0, .c2, .c3 { margin: 1pt 0 !important; padding: 0 !important; text-align: left !important; width: 100% !important; line-height: 1.0 !important; overflow: hidden !important; }
    table { margin: 1pt 0 !important; padding: 0 !important; width: 100% !important; font-size: 9pt !important; border-collapse: collapse !important; }
    td, th { padding: 1pt !important; margin: 0 !important; font-size: 9pt !important; line-height: 1.0 !important; }
    .c15, .c1, .c16, .c6 { background-color: transparent !important; background: none !important; }
    ul, ol, li { margin: 0 !important; padding: 0 !important; line-height: 1.0 !important; }
    h1, h2, h3, h4, h5, h6 { margin: 2pt 0 !important; padding: 0 !important; font-size: 10pt !important; line-height: 1.0 !important; }
    </style>
    """
        if template_name == 'compensazione':
            css_fixes += """
    <style>
    @page:nth(2) { display: none !important; }
    body.c9.doc-content {
        padding-top: 7em !important;
        font-family: "Courier New", Courier, monospace !important;
        font-size: 11pt !important;
        line-height: 1.15 !important;
    }
    body.c9.doc-content td.c8 {
        overflow: visible !important;
    }
    body.c9.doc-content td.c8 p,
    body.c9.doc-content td.c8 span {
        overflow: visible !important;
    }
    body.c9.doc-content td.c8 span.comp-title {
        font-family: Arial, Helvetica, sans-serif !important;
        font-weight: 700 !important;
        font-size: 13pt !important;
    }
    body.c9.doc-content td.c8 span:not(.comp-title) {
        font-family: "Courier New", Courier, monospace !important;
        font-size: 11pt !important;
        line-height: 1.15 !important;
    }
    body.c9.doc-content span.c4 {
        font-weight: 700 !important;
    }
    body.c9.doc-content span.c5 {
        font-weight: 400 !important;
    }
    body.c9.doc-content p.comp-bullet {
        margin: 6pt 0 8pt 0 !important;
        padding-left: 1.35em !important;
        text-indent: -1.35em !important;
    }
    body.c9.doc-content p.comp-quote {
        margin: 0 0 10pt 0 !important;
        padding-left: 2em !important;
        text-indent: 0 !important;
    }
    body.c9.doc-content p.comp-line-data {
        margin-bottom: 3pt !important;
    }
    body.c9.doc-content p.comp-line-gentile {
        margin-bottom: 6pt !important;
    }
    body.c9.doc-content p.comp-saluti {
        margin-top: 12pt !important;
    }
    </style>
    """
    else:
        # Для contratto - ТЕКУЧИЙ КОНТЕНТ (без ограничений страниц)
        css_fixes = """
    <style>
    @page {
        size: A4;
        margin: 1cm;
        border: 4pt solid #a52b4c;
        padding: 0;
    }
    
    body {
        font-family: "Roboto Mono", monospace;
        font-size: 10pt;
        line-height: 1.0;
        margin: 0;
        padding: 0 2cm;
    }
    
    /* Убираем рамки из HTML-контейнеров для предотвращения двойных рамок */
    .c17, .c23 { border: none !important; padding: 0 !important; }
    
    /* Убираем рамки */
    .c20 { border: none !important; padding: 3mm !important; margin: 0 !important; }
    
    /* Разрешаем перенос страниц */
    .page-break { page-break-before: always; }
    
    p { margin: 2pt 0 !important; padding: 0 !important; line-height: 1.0 !important; }
    div { margin: 0 !important; padding: 0 !important; }
    table { margin: 3pt 0 !important; font-size: 10pt !important; }
    .c22 { max-width: none !important; padding: 0 !important; margin: 0 !important; border: none !important; }
    .c14, .c25 { margin-left: 0 !important; }
    .c15 { font-size: 14pt !important; margin: 4pt 0 !important; font-weight: 700 !important; }
    .c10 { font-size: 12pt !important; margin: 3pt 0 !important; font-weight: 700 !important; }
    .c6:empty { height: 0pt !important; margin: 0 !important; padding: 0 !important; }
    .c3 { margin: 1pt 0 !important; }
    .c1, .c16 { background-color: transparent !important; background: none !important; }

    /* ТАБЛИЦА С ПОДПИСЯМИ И ПЕЧАТЬЮ */
    .signatures-tables-wrapper {
        position: relative !important;
        width: 100% !important;
        margin-top: 15pt !important;
        margin-bottom: 10pt !important;
        page-break-inside: avoid !important;
    }

    .signatures-table-base {
        width: 100% !important;
        border-collapse: collapse !important;
        border: none !important;
        background: transparent !important;
        position: relative !important;
        z-index: 10 !important;
    }

    .signatures-table-overlay {
        width: 100% !important;
        border-collapse: collapse !important;
        border: none !important;
        background: transparent !important;
        position: absolute !important;
        top: 0 !important;
        left: 25mm !important;
        z-index: 20 !important;
    }

    .signatures-table-base td, .signatures-table-overlay td {
        border: none !important;
        padding: 10pt !important;
        background: transparent !important;
        vertical-align: bottom !important;
        text-align: center !important;
    }

    /* Разрыв страницы перед пунктом 7 */
    .section-7-firme {
        page-break-before: always !important;
        margin-top: 0 !important;
    }

    /* Печать - увеличена на 30% */
    .seal-img {
        display: block !important;
        margin: 0 auto !important;
        max-width: 97.5mm !important;
        max-height: 42.25mm !important;
        width: auto !important;
        height: auto !important;
    }

    /* Подписи - увеличены на 60% */
    .sing-img {
        display: block !important;
        margin: 0 auto !important;
        max-width: 80mm !important;
        max-height: 32mm !important;
        width: auto !important;
        height: auto !important;
    }
    </style>
    """
    
    html = html.replace('</head>', f'{css_fixes}</head>')
    
    import re
    
    if template_name == 'contratto':
        # 1. Удаляем только старые блоки изображений между разделами, но НЕ удаляем новые подписи
        middle_images_pattern = r'<p class="c3"><span style="overflow: hidden[^>]*><img alt="" src="images/image1\.png"[^>]*></span><span style="overflow: hidden[^>]*><img alt="" src="images/image2\.png"[^>]*></span><span style="overflow: hidden[^>]*><img alt="" src="images/image4\.png"[^>]*></span></p>'
        html = re.sub(middle_images_pattern, '', html)
        
        # 2. Убираем пустые div в конце
        html = re.sub(r'<div><p class="c6 c18"><span class="c7 c23"></span></p></div>$', '', html)
        html = re.sub(r'<p class="c3 c6"><span class="c7 c12"></span></p>$', '', html)
        html = re.sub(r'<p class="c6 c24"><span class="c7 c12"></span></p>$', '', html)
        
        # 3. Убираем лишние высоты
        html = html.replace('class="c13"', 'class="c13" style="height: auto !important;"')
        html = html.replace('class="c19"', 'class="c19" style="height: auto !important;"')
        
    elif template_name in ['carta', 'approvazione', 'compensazione']:
        # Логика для carta (очистка изображений)
        logo_pattern = r'<p class="c12"><span style="overflow: hidden[^>]*><img alt="" src="images/image1\.png"[^>]*></span></p>'
        html = re.sub(logo_pattern, '', html)
        seal_pattern = r'<span style="overflow: hidden[^>]*><img alt="" src="images/image2\.png"[^>]*></span>'
        html = re.sub(seal_pattern, '', html)
        signature_pattern = r'<span style="overflow: hidden[^>]*><img alt="" src="images/image3\.png"[^>]*></span>'
        html = re.sub(signature_pattern, '', html)
        
        html = re.sub(r'<div><p class="c6 c18"><span class="c7 c23"></span></p></div>', '', html)
        html = re.sub(r'<p class="c3 c6"><span class="c7 c12"></span></p>', '', html)
        html = re.sub(r'<p class="c6 c24"><span class="c7 c12"></span></p>', '', html)
        html = re.sub(r'<p class="c6"><span class="c7"></span></p>', '', html)
        
        html = html.replace('class="c13"', 'class="c13" style="height: auto !important;"')
        html = html.replace('class="c19"', 'class="c19" style="height: auto !important;"')
        html = html.replace('class="c5"', 'class="c5" style="height: auto !important;"')
        html = html.replace('class="c9"', 'class="c9" style="height: auto !important;"')
        
        body_end = html.rfind('</body>')
        if body_end != -1:
            content_before_body = html[:body_end].rstrip()
            content_before_body = re.sub(r'(<p[^>]*><span[^>]*></span></p>\s*)+$', '', content_before_body)
            content_before_body = re.sub(r'(<div[^>]*></div>\s*)+$', '', content_before_body)
            html = content_before_body + '\n</body></html>'

    if template_name != 'garanzia':
        html = html.replace('class="c5"', 'class="c5" style="height: auto !important;"')
        html = html.replace('class="c9"', 'class="c9" style="height: auto !important;"')
    
    # Универсальный анализатор (для очистки рамок и высот)
    if template_name != 'garanzia':
        # Вставка функции analyze_and_fix_problematic_elements (упрощенная)
        # 1. Высоты
        html = re.sub(r'\.([a-zA-Z0-9_-]+)\{[^}]*height:\s*([5-9][0-9]{2}|[0-9]{4,})pt[^}]*\}', r'.\1{height:auto;}', html)
        # 2. Красные рамки
        # (Упрощенно, так как полный парсинг сложен, но попробуем простой реплейс)
        pass 
        
    return html


def main():
    """Функция для тестирования PDF конструктора"""
    import sys
    
    template = sys.argv[1] if len(sys.argv) > 1 else 'contratto'
    
    print(f"🧪 Тестируем PDF конструктор для {template} через API...")
    
    test_data = {
        'name': 'Mario Rossi',
        'amount': 15000.0,
        'tan': 7.86,
        'taeg': 8.30, 
        'duration': 36,
        'payment': monthly_payment(15000.0, 36, 7.86)
    }
    
    try:
        if template == 'contratto':
            buf = generate_contratto_pdf(test_data)
            filename = f'test_contratto.pdf'
        elif template == 'garanzia':
            buf = generate_garanzia_pdf(test_data['name'])
            filename = f'test_garanzia.pdf'
        elif template == 'carta':
            buf = generate_carta_pdf(test_data)
            filename = f'test_carta.pdf'
        elif template == 'approvazione':
            buf = generate_approvazione_pdf(test_data)
            filename = f'test_approvazione.pdf'
        elif template == 'compensazione':
            buf = generate_compensazione_pdf({
                'name': test_data['name'],
                'commission': 260.0,
                'indemnity': 750.0,
            })
            filename = 'test_compensazione.pdf'
        else:
            print(f"❌ Неизвестный тип документа: {template}")
            return
        
        with open(filename, 'wb') as f:
            f.write(buf.read())
            
        print(f"✅ PDF создан через API! Файл сохранен как {filename}")
        print(f"📊 Данные: {test_data}")
        
    except Exception as e:
        print(f"❌ Ошибка тестирования API: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()
