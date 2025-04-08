import sys
import os
import json
import csv
import traceback
import glob
from datetime import datetime
import re

# Check if reportlab is available, otherwise fallback to simpler PDF generation
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    reportlab_available = True
except ImportError:
    print("Warning: ReportLab not found. PDF generation might be basic or skipped.")
    reportlab_available = False
    try:
        # Fallback to fpdf if available
        from fpdf import FPDF
        fpdf_available = True
    except ImportError:
        fpdf_available = False
        print("Warning: FPDF not found either. Will generate CSV output only.")

def generate_bom(stats_dir_or_file_path):
    """
    Generates a Bill of Materials (BOM) in CSV and PDF formats from slicing statistics.

    Args:
        stats_dir_or_file_path (str): Path to a directory containing JSON slicing statistics files
                                       OR path to a single JSON statistics file.

    Returns:
        tuple: (csv_path, pdf_path) or None if generation failed
    """
    try:
        print(f"Attempting to generate BOM from: {stats_dir_or_file_path}")
        stats_files = []
        output_base_dir = None

        # Check if input is a directory
        if os.path.isdir(stats_dir_or_file_path):
            print(f"Input is a directory. Scanning for '*_stats.json' files...")
            stats_files = glob.glob(os.path.join(stats_dir_or_file_path, "*_stats.json"))
            if not stats_files:
                print(f"No '*_stats.json' files found in directory: {stats_dir_or_file_path}")
                return None
            output_base_dir = stats_dir_or_file_path # Place BOM folder inside this dir

        # Check if input is a file
        elif os.path.isfile(stats_dir_or_file_path):
            if stats_dir_or_file_path.lower().endswith(".json"):
                print(f"Input is a single file: {stats_dir_or_file_path}")
                stats_files = [stats_dir_or_file_path]
                output_base_dir = os.path.dirname(stats_dir_or_file_path) # Place BOM folder alongside this file
            else:
                print(f"Error: Input file is not a JSON file: {stats_dir_or_file_path}")
                return None
        else:
            print(f"Error: Input path is not a valid directory or file: {stats_dir_or_file_path}")
            return None

        print(f"Found {len(stats_files)} statistics file(s) to process.")

        # Parse all stats files and collect data
        parts_data = []
        total_cost = 0.0
        total_weight = 0.0

        for stats_file in stats_files:
            try:
                with open(stats_file, 'r') as f:
                    metrics = json.load(f)

                # Basic validation - check for expected keys
                if not all(k in metrics for k in ["part_name", "total_weight_g", "price_egp"]):
                     print(f"Warning: Skipping {stats_file} - missing required keys (part_name, total_weight_g, price_egp).")
                     continue

                parts_data.append(metrics)
                # Ensure values are numeric before adding, default to 0 if None or invalid
                total_cost += float(metrics.get("price_egp") or 0)
                total_weight += float(metrics.get("total_weight_g") or 0)
            except json.JSONDecodeError:
                print(f"Warning: Skipping {stats_file} - Invalid JSON.")
            except (ValueError, TypeError) as e:
                 print(f"Warning: Skipping {stats_file} - Invalid numeric data for weight/price: {e}")
            except Exception as e:
                print(f"Warning: Failed to process {stats_file}: {str(e)}")
                continue

        if not parts_data:
            print("No valid statistics data found after processing.")
            return None

        # Sort parts by name
        parts_data.sort(key=lambda x: x.get("part_name", ""))

        # Create output directory relative to where stats files were found
        if not output_base_dir: # Should not happen if logic above is correct, but as safety
             output_base_dir = os.getcwd()
             print(f"Warning: Could not determine output base directory, using current working directory: {output_base_dir}")

        bom_dir = os.path.join(output_base_dir, "BOM")
        if not os.path.exists(bom_dir):
            try:
                 os.makedirs(bom_dir)
                 print(f"Created BOM output directory: {bom_dir}")
            except OSError as e:
                 print(f"Error: Could not create BOM output directory '{bom_dir}': {e}")
                 return None


        # Generate timestamp for filenames
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # Create a base filename from the input path if it was a single file
        if len(stats_files) == 1 and os.path.isfile(stats_dir_or_file_path):
            base_filename = os.path.splitext(os.path.basename(stats_dir_or_file_path))[0]
            # Remove common suffixes like '_stats' if present
            if base_filename.lower().endswith('_stats'):
                base_filename = base_filename[:-6]
            bom_base_name = f"BOM_{base_filename}_{timestamp}"
        else:
            # Use a generic name if multiple files or directory input
            # Extract folder name if input was a directory
            dir_name = os.path.basename(os.path.normpath(stats_dir_or_file_path)) if os.path.isdir(stats_dir_or_file_path) else "MultiFile"
            bom_base_name = f"BOM_{dir_name}_{timestamp}"


        # Generate CSV BOM
        csv_path = os.path.join(bom_dir, f"{bom_base_name}.csv")
        generate_csv_bom(csv_path, parts_data, total_cost, total_weight)

        # Generate PDF BOM
        pdf_path = os.path.join(bom_dir, f"{bom_base_name}.pdf")
        pdf_generated = False
        if reportlab_available:
            pdf_generated = generate_pdf_bom_reportlab(pdf_path, parts_data, total_cost, total_weight)
        elif fpdf_available:
             # Only try fpdf if reportlab wasn't available OR if reportlab failed
             if not pdf_generated:
                 print("Attempting PDF generation with FPDF as fallback...")
                 pdf_generated = generate_pdf_bom_fpdf(pdf_path, parts_data, total_cost, total_weight)
        else:
            print("Skipping PDF generation due to missing dependencies (ReportLab/FPDF).")

        # Set pdf_path to None if PDF generation ultimately failed or was skipped
        if not pdf_generated:
             pdf_path = None


        return csv_path, pdf_path

    except Exception as e:
        print(f"× ERROR in generate_bom: {str(e)}")
        print("Detailed traceback:")
        traceback.print_exc()
        return None


def generate_csv_bom(csv_path, parts_data, total_cost, total_weight):
    """Generates a CSV BOM file"""
    try:
        with open(csv_path, 'w', newline='') as csv_file:
            fieldnames = [
                "Part Name",
                "Object Weight (g)",
                "Supports Weight (g)",
                "Total Weight (g)",
                "Price (EGP)",
                "Dimensions (mm)",
                "Print Time",
                "Print Settings"
            ]

            writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            writer.writeheader()

            for part in parts_data:
                writer.writerow({
                    # Provide default values directly in .get() for simplicity
                    "Part Name": part.get("part_name", "Unknown"),
                    "Object Weight (g)": f"{float(part.get('object_weight_g') or 0):.4f}",
                    "Supports Weight (g)": f"{float(part.get('supports_weight_g') or 0):.4f}",
                    "Total Weight (g)": f"{float(part.get('total_weight_g') or 0):.4f}",
                    "Price (EGP)": f"{float(part.get('price_egp') or 0):.2f}",
                    "Dimensions (mm)": part.get("dimensions_mm", "N/A"), # Use N/A for missing string values
                    "Print Time": part.get("print_time", "N/A"),
                    "Print Settings": part.get("print_settings", "N/A")
                })

            # Add summary row
            writer.writerow({
                "Part Name": f"TOTAL ({len(parts_data)} parts)",
                "Object Weight (g)": "",
                "Supports Weight (g)": "",
                "Total Weight (g)": f"{total_weight:.4f}",
                "Price (EGP)": f"${total_cost:.2f}",
                "Dimensions (mm)": "",
                "Print Time": "",
                "Print Settings": ""
            })

        print(f"CSV BOM generated: {csv_path}")
        return True

    except Exception as e:
        print(f"Error generating CSV BOM: {str(e)}")
        return False

def generate_pdf_bom_reportlab(pdf_path, parts_data, total_cost, total_weight):
    """Generates a PDF BOM file using ReportLab"""
    if not reportlab_available: return False
    try:
        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=landscape(letter),
            rightMargin=0.5*inch, leftMargin=0.5*inch,
            topMargin=0.5*inch, bottomMargin=0.5*inch
        )
        styles = getSampleStyleSheet()
        title_style = styles["Title"]
        normal_style = styles["Normal"]
        cell_style = ParagraphStyle("CellStyle", parent=normal_style, fontSize=9, leading=12)
        elements = []

        elements.append(Paragraph("3D Printing Bill of Materials", title_style))
        elements.append(Spacer(1, 0.25*inch))
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        elements.append(Paragraph(f"Generated on: {current_date}", normal_style))
        elements.append(Paragraph(f"Total Parts: {len(parts_data)}", normal_style))
        elements.append(Paragraph(f"Total Weight: {total_weight:.4f} g", normal_style))
        elements.append(Paragraph(f"Total Cost: ${total_cost:.2f}", normal_style))
        elements.append(Spacer(1, 0.25*inch))

        table_data = [
            ["Part Name", "Object Wt (g)", "Supports Wt (g)", "Total Wt (g)",
             "Price (EGP)", "Dimensions (mm)", "Print Time", "Print Settings"]
        ]

        for part in parts_data:
            # *** FIX: Ensure None values are converted to strings before passing to Paragraph ***
            part_name = str(part.get("part_name", "Unknown"))
            obj_weight = f"{float(part.get('object_weight_g') or 0):.4f}"
            sup_weight = f"{float(part.get('supports_weight_g') or 0):.4f}"
            total_weight_part = f"{float(part.get('total_weight_g') or 0):.4f}"
            price = f"${float(part.get('price_egp') or 0):.2f}"
            dimensions = str(part.get("dimensions_mm", "N/A")) # Use N/A or ""
            print_time = str(part.get("print_time", "N/A"))     # Use N/A or ""
            settings = str(part.get("print_settings", "N/A")) # Use N/A or ""

            table_data.append([
                Paragraph(part_name, cell_style),
                Paragraph(obj_weight, cell_style),
                Paragraph(sup_weight, cell_style),
                Paragraph(total_weight_part, cell_style),
                Paragraph(price, cell_style),
                Paragraph(dimensions, cell_style), # Pass the ensured string
                Paragraph(print_time, cell_style), # Pass the ensured string
                Paragraph(settings, cell_style)    # Pass the ensured string
            ])

        table_data.append([
            Paragraph(f"TOTAL ({len(parts_data)} parts)", cell_style),
            "", "", Paragraph(f"{total_weight:.4f}", cell_style),
            Paragraph(f"${total_cost:.2f}", cell_style), "", "", ""
        ])

        col_widths = [1.5*inch, 1*inch, 1*inch, 1*inch, 0.8*inch, 1.5*inch, 1*inch, 2*inch]
        table = Table(table_data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACgROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACgROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (1, 1), (4, -1), 'CENTER'),
            ('ALIGN', (5, 1), (-1, -1), 'CENTER'),
        ]))
        elements.append(table)
        doc.build(elements)
        print(f"PDF BOM generated using ReportLab: {pdf_path}")
        return True

    except Exception as e:
        print(f"Error generating PDF BOM with ReportLab: {str(e)}")
        traceback.print_exc()
        return False

def generate_pdf_bom_fpdf(pdf_path, parts_data, total_cost, total_weight):
    """Generates a PDF BOM file using FPDF (fallback option)"""
    if not fpdf_available: return False
    try:
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, "3D Printing Bill of Materials", 0, 1, 'C')
        pdf.ln(5)
        pdf.set_font('Arial', '', 12)
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        pdf.cell(0, 6, f"Generated on: {current_date}", 0, 1)
        pdf.cell(0, 6, f"Total Parts: {len(parts_data)}", 0, 1)
        pdf.cell(0, 6, f"Total Weight: {total_weight:.4f} g", 0, 1)
        pdf.cell(0, 6, f"Total Cost: ${total_cost:.2f}", 0, 1)
        pdf.ln(5)

        col_widths = [45, 30, 30, 30, 25, 40, 30, 47]
        pdf.set_font('Arial', 'B', 10)
        pdf.set_fill_color(200, 220, 255)
        headers = ["Part Name", "Object Wt (g)", "Supports Wt (g)", "Total Wt (g)",
                   "Price (EGP)", "Dimensions (mm)", "Print Time", "Print Settings"]
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 8, header, 1, 0, 'C', 1)
        pdf.ln()

        pdf.set_font('Arial', '', 9)
        for part in parts_data:
            # *** FIX: Ensure None values are converted to strings before passing to cell/multi_cell ***
            part_name = str(part.get("part_name", "Unknown"))
            obj_weight = f"{float(part.get('object_weight_g') or 0):.4f}"
            sup_weight = f"{float(part.get('supports_weight_g') or 0):.4f}"
            total_weight_part = f"{float(part.get('total_weight_g') or 0):.4f}"
            price = f"${float(part.get('price_egp') or 0):.2f}"
            dimensions = str(part.get("dimensions_mm", "N/A")) # Use N/A or ""
            print_time = str(part.get("print_time", "N/A"))     # Use N/A or ""
            settings = str(part.get("print_settings", "N/A")) # Use N/A or ""

            # Use MultiCell for potentially long fields like name and settings to allow wrapping
            y_start = pdf.get_y()
            pdf.multi_cell(col_widths[0], 6, part_name, 1, 'L')
            pdf.set_xy(pdf.get_x() + col_widths[0], y_start) # Reset X position for next cell in row

            pdf.cell(col_widths[1], 6, obj_weight, 1, 0, 'C')
            pdf.cell(col_widths[2], 6, sup_weight, 1, 0, 'C')
            pdf.cell(col_widths[3], 6, total_weight_part, 1, 0, 'C')
            pdf.cell(col_widths[4], 6, price, 1, 0, 'C')
            pdf.cell(col_widths[5], 6, dimensions, 1, 0, 'C') # Dimensions usually fit
            pdf.cell(col_widths[6], 6, print_time, 1, 0, 'C') # Print time usually fits

            # Use MultiCell for settings, which might wrap
            y_before_settings = pdf.get_y()
            x_before_settings = pdf.get_x()
            pdf.multi_cell(col_widths[7], 6, settings, 1, 'L')
            # Adjust Y position for the next row based on max height of the row (primarily settings cell)
            y_after_settings = pdf.get_y()
            # Set Y position for the next row start; max(y_start + cell_height, y_after_settings)
            # Using simple y_after_settings might be sufficient if settings is the only multi_cell per row
            pdf.set_y(y_after_settings if y_after_settings > y_start else y_start + 6)
            pdf.set_x(pdf.l_margin) # Go back to left margin for next row


        pdf.set_font('Arial', 'B', 9)
        pdf.set_fill_color(220, 220, 220)
        pdf.cell(col_widths[0], 6, f"TOTAL ({len(parts_data)} parts)", 1, 0, 'L', 1)
        pdf.cell(col_widths[1], 6, "", 1, 0, 'C', 1)
        pdf.cell(col_widths[2], 6, "", 1, 0, 'C', 1)
        pdf.cell(col_widths[3], 6, f"{total_weight:.4f}", 1, 0, 'C', 1)
        pdf.cell(col_widths[4], 6, f"${total_cost:.2f}", 1, 0, 'C', 1)
        pdf.cell(col_widths[5], 6, "", 1, 0, 'C', 1)
        pdf.cell(col_widths[6], 6, "", 1, 0, 'C', 1)
        pdf.cell(col_widths[7], 6, "", 1, 0, 'C', 1)

        pdf.output(pdf_path)
        print(f"PDF BOM generated using FPDF: {pdf_path}")
        return True

    except Exception as e:
        print(f"Error generating PDF BOM with FPDF: {str(e)}")
        traceback.print_exc()
        return False


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python generate_bom.py <stats_directory | path_to_single_stats.json>")
        sys.exit(1)

    stats_path_arg = os.path.abspath(sys.argv[1])
    if not os.path.exists(stats_path_arg):
        print(f"Error: Input path not found: {stats_path_arg}")
        sys.exit(1)

    result = generate_bom(stats_path_arg)

    if result is None:
        print("--- Failed to generate BOM ---")
        sys.exit(1)

    csv_path, pdf_path = result
    print("--- BOM generation completed ---") # Changed message slightly
    if csv_path:
        print(f"CSV Output: {csv_path}")
    else:
        print("CSV Output: Failed or skipped")

    if pdf_path:
        print(f"PDF Output: {pdf_path}")
    else:
        print("PDF Output: Failed or skipped (missing dependencies or error during PDF generation)")

    sys.exit(0)