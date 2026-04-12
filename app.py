import os
import io
import re
from datetime import datetime

import anthropic
import yfinance as yf
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from dotenv import load_dotenv
from flask import Flask, render_template, request, send_file

load_dotenv()

app = Flask(__name__)

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")


def format_number(value):
    """Format a number in millions with Spanish locale style."""
    if value is None:
        return {"formatted": "N/D", "negative": False, "pct": None}
    millions = value / 1_000_000
    negative = millions < 0
    formatted = f"{millions:,.1f} M"
    return {"formatted": formatted, "negative": negative, "pct": None}


def format_pct(value, revenue):
    """Format a number with its percentage over revenue."""
    if value is None or revenue is None or revenue == 0:
        return {"formatted": "N/D", "negative": False, "pct": None}
    millions = value / 1_000_000
    negative = millions < 0
    pct = (value / revenue) * 100
    formatted = f"{millions:,.1f} M"
    pct_str = f"({pct:+.1f}%)"
    return {"formatted": formatted, "negative": negative, "pct": pct_str}


def safe_get(df, key, col):
    """Safely get a value from a DataFrame."""
    try:
        return float(df.loc[key, col])
    except (KeyError, TypeError, ValueError):
        return None


def get_financial_data(ticker_symbol):
    """Download and process financial data from Yahoo Finance."""
    ticker = yf.Ticker(ticker_symbol)
    income_stmt = ticker.income_stmt

    if income_stmt is None or income_stmt.empty:
        return None, None, None, None

    info = ticker.info
    company_name = info.get("longName") or info.get("shortName") or ticker_symbol.upper()
    company_info = {
        "sector": info.get("sector", "N/D"),
        "industry": info.get("industry", "N/D"),
        "currency": info.get("currency", "USD"),
    }

    columns = income_stmt.columns[:4]
    years = [col.strftime("%Y") if hasattr(col, "strftime") else str(col) for col in columns]

    key_map = {
        "Total Revenue": ["Total Revenue", "TotalRevenue"],
        "Cost Of Revenue": ["Cost Of Revenue", "CostOfRevenue"],
        "Gross Profit": ["Gross Profit", "GrossProfit"],
        "Operating Expense": ["Operating Expense", "OperatingExpense", "Total Operating Expenses"],
        "Operating Income": ["Operating Income", "OperatingIncome"],
        "EBITDA": ["EBITDA", "Normalized EBITDA"],
        "Net Income": ["Net Income", "NetIncome", "Net Income Common Stockholders"],
        "EBIT": ["EBIT"],
        "Interest Expense": ["Interest Expense", "InterestExpense"],
        "Tax Provision": ["Tax Provision", "TaxProvision", "Income Tax Expense"],
    }

    def find_key(label):
        candidates = key_map.get(label, [label])
        for candidate in candidates:
            if candidate in income_stmt.index:
                return candidate
        return None

    rows = []
    for col_idx, col in enumerate(columns):
        revenue = safe_get(income_stmt, find_key("Total Revenue"), col)
        if col_idx == 0:
            pass
        rows.append({"col": col, "revenue": revenue})

    def build_row(label, key_label, css_class="", show_pct=True):
        key = find_key(key_label)
        values = []
        for r in rows:
            val = safe_get(income_stmt, key, r["col"]) if key else None
            if show_pct:
                values.append(format_pct(val, r["revenue"]))
            else:
                values.append(format_number(val))
        return {"label": label, "cells": values, "css_class": css_class}

    financial_data = [
        build_row("Ingresos totales", "Total Revenue", css_class="subtotal", show_pct=False),
        build_row("Coste de ventas", "Cost Of Revenue"),
        build_row("Margen bruto", "Gross Profit", css_class="subtotal"),
        build_row("Gastos operativos", "Operating Expense"),
        build_row("Resultado operativo (EBIT)", "Operating Income", css_class="subtotal"),
        build_row("EBITDA", "EBITDA"),
        build_row("Gastos financieros", "Interest Expense"),
        build_row("Impuestos", "Tax Provision"),
        build_row("Beneficio neto", "Net Income", css_class="total"),
    ]

    raw_data = {}
    for year_label, col in zip(years, columns):
        year_data = {}
        for concept, key_label in [
            ("ingresos", "Total Revenue"),
            ("coste_ventas", "Cost Of Revenue"),
            ("margen_bruto", "Gross Profit"),
            ("gastos_operativos", "Operating Expense"),
            ("resultado_operativo", "Operating Income"),
            ("ebitda", "EBITDA"),
            ("gastos_financieros", "Interest Expense"),
            ("impuestos", "Tax Provision"),
            ("beneficio_neto", "Net Income"),
        ]:
            key = find_key(key_label)
            year_data[concept] = safe_get(income_stmt, key, col) if key else None
        raw_data[year_label] = year_data

    return financial_data, years, company_name, company_info, raw_data


def generate_memo(company_name, ticker_symbol, years, raw_data, currency):
    """Generate the financial memo using Anthropic's Claude API."""
    data_text = ""
    for year in years:
        d = raw_data[year]
        data_text += f"\n--- {year} ---\n"
        for concept, value in d.items():
            if value is not None:
                data_text += f"  {concept}: {value / 1_000_000:,.1f} millones {currency}\n"
            else:
                data_text += f"  {concept}: No disponible\n"

    prompt = f"""Eres un analista financiero profesional. Redacta una nota de memoria financiera en español
con lenguaje formal contable sobre la empresa {company_name} (ticker: {ticker_symbol}).

Utiliza los siguientes datos reales de la cuenta de pérdidas y ganancias:

{data_text}

La nota debe incluir:
1. Un encabezamiento formal indicando que es la nota explicativa de la cuenta de pérdidas y ganancias
2. Un análisis de la evolución de los ingresos y el margen bruto
3. Un análisis de los gastos operativos y el resultado operativo
4. Un comentario sobre el EBITDA y su evolución
5. Un análisis del beneficio neto y la rentabilidad
6. Una conclusión con valoración general de la salud financiera

Usa formato profesional con párrafos bien estructurados. Incluye cifras concretas y porcentajes
de variación interanual. Redacta en un tono formal adecuado para una memoria anual corporativa.
No uses formato markdown. Escribe en texto plano con párrafos separados por líneas en blanco."""

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )

    return message.content[0].text


def create_word_document(company_name, ticker_symbol, memo_text, financial_data, years):
    """Create a Word document with the financial memo."""
    doc = Document()

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x1A, 0x23, 0x32)

    title = doc.add_heading(level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(f"Nota de Memoria Financiera")
    run.font.color.rgb = RGBColor(0x0A, 0x24, 0x63)
    run.font.size = Pt(22)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f"{company_name} ({ticker_symbol.upper()})")
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x1E, 0x52, 0x99)
    run.bold = True

    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = date_para.add_run(f"Fecha de elaboracion: {datetime.now().strftime('%d/%m/%Y')}")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    doc.add_paragraph("")

    heading = doc.add_heading("Datos Financieros", level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x0A, 0x24, 0x63)

    num_cols = 1 + len(years)
    table = doc.add_table(rows=1, cols=num_cols)
    table.style = "Light Grid Accent 1"

    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Concepto"
    for i, year in enumerate(years):
        hdr_cells[i + 1].text = year

    for row_data in financial_data:
        row_cells = table.add_row().cells
        row_cells[0].text = row_data["label"]
        for i, val in enumerate(row_data["cells"]):
            text = val["formatted"]
            if val["pct"] is not None:
                text += f" {val['pct']}"
            row_cells[i + 1].text = text

    doc.add_paragraph("")

    heading = doc.add_heading("Nota Explicativa", level=1)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x0A, 0x24, 0x63)

    for paragraph_text in memo_text.split("\n\n"):
        paragraph_text = paragraph_text.strip()
        if paragraph_text:
            doc.add_paragraph(paragraph_text)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


@app.route("/", methods=["GET", "POST"])
def index():
    context = {"current_year": datetime.now().year}

    if request.method == "POST":
        ticker_symbol = request.form.get("ticker", "").strip().upper()
        context["ticker"] = ticker_symbol

        if not ticker_symbol:
            context["error"] = "Por favor, introduce un ticker valido."
            return render_template("index.html", **context)

        if not ANTHROPIC_API_KEY:
            context["error"] = (
                "No se ha configurado la API key de Anthropic. "
                "Crea un archivo .env con ANTHROPIC_API_KEY=tu_clave"
            )
            return render_template("index.html", **context)

        try:
            result = get_financial_data(ticker_symbol)
            if result[0] is None:
                context["error"] = (
                    f"No se encontraron datos financieros para '{ticker_symbol}'. "
                    "Verifica que el ticker sea correcto."
                )
                return render_template("index.html", **context)

            financial_data, years, company_name, company_info, raw_data = result
            currency = company_info.get("currency", "USD")

            memo_raw = generate_memo(company_name, ticker_symbol, years, raw_data, currency)
            memo_html = memo_raw.replace("\n\n", "</p><p>").replace("\n", "<br>")
            memo_html = f"<p>{memo_html}</p>"

            app.config[f"memo_{ticker_symbol}"] = memo_raw
            app.config[f"data_{ticker_symbol}"] = financial_data
            app.config[f"years_{ticker_symbol}"] = years
            app.config[f"name_{ticker_symbol}"] = company_name

            context.update({
                "financial_data": financial_data,
                "years": years,
                "company_name": company_name,
                "company_info": company_info,
                "memo": memo_html,
            })

        except Exception as e:
            context["error"] = f"Error al procesar los datos: {str(e)}"

    return render_template("index.html", **context)


@app.route("/download/<ticker>")
def download(ticker):
    ticker = ticker.upper()
    memo = app.config.get(f"memo_{ticker}")
    financial_data = app.config.get(f"data_{ticker}")
    years = app.config.get(f"years_{ticker}")
    company_name = app.config.get(f"name_{ticker}")

    if not memo:
        return "No hay datos disponibles. Realiza primero el analisis.", 404

    buffer = create_word_document(company_name, ticker, memo, financial_data, years)
    filename = f"Nota_Memoria_{company_name.replace(' ', '_')}_{ticker}.docx"

    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
