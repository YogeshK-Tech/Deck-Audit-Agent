from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import pandas as pd
import tempfile
import os
from typing import List, Dict, Any
from datetime import datetime

class ReportGenerator:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            textColor=colors.black,
            alignment=1  # Center alignment
        )
        
    def generate_pdf_report(self, audit_results: List[Dict[str, Any]]) -> str:
        """Generate PDF audit report"""
        try:
            # Create temporary file
            temp_dir = tempfile.mkdtemp()
            pdf_path = os.path.join(temp_dir, f"audit_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            
            # Create PDF document
            doc = SimpleDocTemplate(pdf_path, pagesize=A4)
            story = []
            
            # Title
            title = Paragraph("Deck Audit Report", self.title_style)
            story.append(title)
            story.append(Spacer(1, 20))
            
            # Summary statistics
            total_numbers = len(audit_results)
            matches = len([r for r in audit_results if r["status"] == "Match"])
            mismatches = len([r for r in audit_results if r["status"] == "Mismatch"])
            untraceable = len([r for r in audit_results if r["status"] == "Untraceable"])
            
            summary_text = f"""
            <b>Summary:</b><br/>
            Total Numbers Analyzed: {total_numbers}<br/>
            Matches: {matches} ({matches/total_numbers*100:.1f}%)<br/>
            Mismatches: {mismatches} ({mismatches/total_numbers*100:.1f}%)<br/>
            Untraceable: {untraceable} ({untraceable/total_numbers*100:.1f}%)<br/>
            Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            """
            
            summary = Paragraph(summary_text, self.styles['Normal'])
            story.append(summary)
            story.append(Spacer(1, 30))
            
            # Detailed results table
            if audit_results:
                story.append(Paragraph("Detailed Audit Results", self.styles['Heading2']))
                story.append(Spacer(1, 10))
                
                # Prepare table data
                table_data = [['Slide', 'PPT Text', 'Status', 'PPT Value', 'Excel Value', 'Suggested Fix', 'Excel Source']]
                
                for result in audit_results:
                    row = [
                        str(result.get("slide", "")),
                        result.get("text", "")[:30] + "..." if len(result.get("text", "")) > 30 else result.get("text", ""),
                        result.get("status", ""),
                        str(result.get("ppt_value", "")),
                        str(result.get("excel_value", "")) if result.get("excel_value") is not None else "N/A",
                        result.get("suggested_fix", "") or "N/A",
                        f"{result.get('excel_file', '')}/{result.get('excel_sheet', '')}" if result.get('excel_file') else "N/A"
                    ]
                    table_data.append(row)
                
                # Create table
                table = Table(table_data, colWidths=[0.5*inch, 1.5*inch, 0.8*inch, 0.8*inch, 0.8*inch, 1*inch, 1.5*inch])
                
                # Style the table
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                
                # Color code by status
                for i, result in enumerate(audit_results, 1):
                    if result["status"] == "Match":
                        table.setStyle(TableStyle([('BACKGROUND', (2, i), (2, i), colors.lightgreen)]))
                    elif result["status"] == "Mismatch":
                        table.setStyle(TableStyle([('BACKGROUND', (2, i), (2, i), colors.lightcoral)]))
                    elif result["status"] == "Untraceable":
                        table.setStyle(TableStyle([('BACKGROUND', (2, i), (2, i), colors.lightyellow)]))
                
                story.append(table)
            
            # Build PDF
            doc.build(story)
            print(f"PDF report generated: {pdf_path}")
            return pdf_path
            
        except Exception as e:
            print(f"Error generating PDF report: {str(e)}")
            raise Exception(f"Failed to generate PDF report: {str(e)}")
    
    def generate_csv_report(self, audit_results: List[Dict[str, Any]]) -> str:
        """Generate CSV audit report"""
        try:
            # Create temporary file
            temp_dir = tempfile.mkdtemp()
            csv_path = os.path.join(temp_dir, f"audit_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            
            # Convert to DataFrame
            df_data = []
            for result in audit_results:
                df_data.append({
                    'Slide_Number': result.get("slide", ""),
                    'PPT_Text': result.get("text", ""),
                    'Status': result.get("status", ""),
                    'PPT_Value': result.get("ppt_value", ""),
                    'Excel_Value': result.get("excel_value", ""),
                    'Suggested_Fix': result.get("suggested_fix", ""),
                    'Excel_File': result.get("excel_file", ""),
                    'Excel_Sheet': result.get("excel_sheet", ""),
                    'Excel_Cell': result.get("cell", ""),
                    'Reasoning': result.get("reasoning", ""),
                    'Context': result.get("context", ""),
                    'Confidence': result.get("confidence", "")
                })
            
            df = pd.DataFrame(df_data)
            df.to_csv(csv_path, index=False)
            
            print(f"CSV report generated: {csv_path}")
            return csv_path
            
        except Exception as e:
            print(f"Error generating CSV report: {str(e)}")
            raise Exception(f"Failed to generate CSV report: {str(e)}")