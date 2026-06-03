from django.shortcuts import render
from django.db import connection
from django.http import HttpResponse
import openpyxl
from reportlab.pdfgen import canvas


# -----------------------------
# GET DATA FUNCTION (REUSED)
# -----------------------------
def get_data():

    query = """
         SELECT 

                pol.SEGMENT_CODE,
                pol.ISSUE_DATE,
                pol.SUMINSURED,
                pol.EXRATE,
                pol.GROSS_PREMIUM,
                pol.NET_PREMIUM,
                pol.MODIFIED_BY,
                cust.NAME,
                cls.CLASS_CODE,
                typ.NAME,
                pol.AGENT_TYPE

            FROM igeneral.GPD_POLICIES pol

                JOIN erp.FCS_CUSTOMERS cust
                    ON pol.FCS_CST_ID = cust.ID

                JOIN igeneral.GST_CLASSES cls
                    ON pol.GST_CLS_ID = cls.ID

                JOIN igeneral.GST_ENDORSEMENT_TYPES typ
                    ON pol.GST_ENT_ID = typ.ID

            WHERE pol.IS_POSTED = 1
            and pol.ISSUE_DATE BETWEEN to_date('17/05/2026','dd/mm/yyyy')
            and to_date('19/05/2026','dd/mm/yyyy')
            FETCH FIRST 50 ROWS ONLY
    """

    with connection.cursor() as cursor:
        cursor.execute(query)
        rows = cursor.fetchall()

    result = []

    for row in rows:
        result.append({
            "segment_code": row[0],
            "issue_date": row[1],
            "suminsured": row[2],
            "exrate": row[3],
            "gross_premium": row[4],
            "net_premium": row[5],
            "modified_by": row[6],
            "customer_name": row[7],
            "class_code": row[8],
            "endorsement_type": row[9],
            "agent_type": row[10],
        })

    return result


# -----------------------------
# MAIN PAGE VIEW
# -----------------------------
def customer_page(request):

    data = get_data()

    return render(request, "customers.html", {"data": data})


# -----------------------------
# PDF EXPORT
# -----------------------------
def export_pdf(request):

    data = get_data()

    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="report.pdf"'

    p = canvas.Canvas(response)

    y = 800

    p.setFont("Helvetica-Bold", 14)
    p.drawString(200, y, "Customer Report")
    y -= 40

    p.setFont("Helvetica", 10)

    for item in data:

        line = f"{item['segment_code']} | {item['issue_date']} | {item['customer_name']} | {item['gross_premium']}"

        p.drawString(40, y, line)

        y -= 20

        if y < 50:
            p.showPage()
            y = 800

    p.save()
    return response


# -----------------------------
# EXCEL EXPORT
# -----------------------------
def export_excel(request):

    data = get_data()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"

    headers = [
        "SEGMENT",
        "ISSUE DATE",
        "SUM INSURED",
        "EX RATE",
        "GROSS PREMIUM",
        "NET PREMIUM",
        "MODIFIED BY",
        "CUSTOMER",
        "CLASS CODE",
        "ENDORSEMENT",
        "AGENT TYPE"
    ]

    ws.append(headers)

    for item in data:

        ws.append([
            item["segment_code"],
            item["issue_date"],
            item["suminsured"],
            item["exrate"],
            item["gross_premium"],
            item["net_premium"],
            item["modified_by"],
            item["customer_name"],
            item["class_code"],
            item["endorsement_type"],
            item["agent_type"],
        ])

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    response['Content-Disposition'] = 'attachment; filename="report.xlsx"'

    wb.save(response)
    return response