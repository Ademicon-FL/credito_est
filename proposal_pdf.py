"""
proposal_pdf.py
Geração de PDF de Proposta de Crédito Estruturado (A4)
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.colors import HexColor, lightgrey
from reportlab.lib.units import cm
from datetime import datetime
import io
COR_PRINCIPAL = HexColor("#1F3864")
COR_SECUNDARIA = HexColor("#2F75B6")

def gerar_pdf(resultado, nome_cliente: str, data_base: datetime, titulo: str = "Proposta de Crédito Estruturado via Consórcio") -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=2*cm, leftMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Titulo", fontSize=18, leading=22, alignment=TA_CENTER, textColor=COR_PRINCIPAL, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="Texto", fontSize=10, leading=14))
    elements = []
    elements.append(Paragraph(titulo, styles["Titulo"]))
    elements.append(Spacer(1, 20))
    dados = [["Cliente", nome_cliente or "—"],["Data base", data_base.strftime('%d/%m/%Y')],["Crédito líquido total", f"R$ {resultado.credito_liquido_total:,.2f}"]]
    t = Table(dados, colWidths=[6*cm,8*cm])
    t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.5,lightgrey)]))
    elements.append(t)
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()
