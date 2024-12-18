from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.chart import (BarChart, LineChart, PieChart, Reference)
from openpyxl.styles import Font, Alignment, NamedStyle, PatternFill
from openpyxl.utils import get_column_letter
import os
import shutil
from datetime import datetime
from fastapi.middleware.cors import CORSMiddleware
import uuid

# Initialize FastAPI app
app = FastAPI()
app.mount("/static", StaticFiles(directory="."), name="static")
templates = Jinja2Templates(directory=".")

# CORS Middleware to allow frontend communication
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/index")
async def get_index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

def generate_report(input_file: BytesIO, month: str = None) -> str:
    """
    Generate an enhanced Excel report from input file
    
    Args:
        input_file (BytesIO): Excel file in memory
        month (str, optional): Month for the report. Defaults to current month.
    
    Returns:
        str: Path to the generated report file
    """
    # If no month provided, use current month
    if not month:
        month = datetime.now().strftime("%B %Y")

    # Load workbook
    wb = load_workbook(input_file)
    sheet = wb.active

    # Validate data
    if sheet.max_row <= 1:
        raise ValueError("No data found in the Excel file")

    # Add report title and styling
    add_report_title(sheet, month)

    # Create charts
    create_charts(sheet)

    # Apply formatting
    apply_formatting(sheet)

    # Generate unique filename
    report_filename = f"report_{month.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    report_file_path = os.path.join(UPLOAD_DIR, report_filename)

    # Save the report
    wb.save(report_file_path)
    return report_file_path

def add_report_title(sheet, month):
    """Add a professional title to the Excel report"""
    # Insert rows for title
    sheet.insert_rows(1, amount=3)
    
    # Merge and style main title
    title_cell = sheet.cell(row=1, column=1, value='Sales Report')
    sheet.merge_cells(f'A1:{get_column_letter(sheet.max_column)}1')
    title_cell.font = Font(name='Calibri', size=20, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Add month subtitle
    subtitle_cell = sheet.cell(row=2, column=1, value=month)
    sheet.merge_cells(f'A2:{get_column_letter(sheet.max_column)}2')
    subtitle_cell.font = Font(name='Calibri', size=14)
    subtitle_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Light blue background for title rows
    title_fill = PatternFill(start_color='E6F2FF', end_color='E6F2FF', fill_type='solid')
    for row in sheet[1:3]:
        for cell in row:
            cell.fill = title_fill

def create_charts(sheet):
    """Create multiple charts for the report"""
    # Determine data range (start from 4th row after title rows)
    min_row = 4
    max_row = sheet.max_row
    min_column = 1
    max_column = sheet.max_column

    # Prepare references
    data = Reference(sheet, 
                     min_col=min_column + 1, 
                     max_col=max_column, 
                     min_row=min_row, 
                     max_row=max_row)
    category = Reference(sheet, 
                         min_col=min_column, 
                         max_col=min_column, 
                         min_row=min_row + 1, 
                         max_row=max_row)

    # Bar Chart
    barchart = BarChart()
    barchart.add_data(data, titles_from_data=True)
    barchart.set_categories(category)
    barchart.title = "Sales by Product Line"
    barchart.x_axis.title = "Product"
    barchart.y_axis.title = "Sales Amount"
    sheet.add_chart(barchart, 'B10')

    # Line Chart
    linechart = LineChart()
    linechart.add_data(data, titles_from_data=True)
    linechart.set_categories(category)
    linechart.title = "Sales Trend Analysis"
    linechart.x_axis.title = "Product"
    linechart.y_axis.title = "Sales"
    sheet.add_chart(linechart, 'K10')

    # Pie Chart
    piechart = PieChart()
    pie_data = Reference(sheet, 
                         min_col=max_column, 
                         max_col=max_column, 
                         min_row=min_row + 1, 
                         max_row=max_row)
    pie_labels = Reference(sheet, 
                           min_col=min_column, 
                           max_col=min_column, 
                           min_row=min_row + 1, 
                           max_row=max_row)
    piechart.add_data(pie_data, titles_from_data=False)
    piechart.set_categories(pie_labels)
    piechart.title = "Sales Distribution"
    sheet.add_chart(piechart, 'B30')

def apply_formatting(sheet):
    """Apply advanced formatting to the sheet"""
    # Currency formatting for numeric columns
    for col in range(2, sheet.max_column + 1):
        column_letter = get_column_letter(col)
        
        # Add total row
        total_cell = sheet.cell(row=sheet.max_row + 1, column=col, 
                                value=f'=SUM({column_letter}4:{column_letter}{sheet.max_row})')
        
        # Apply number formatting
        for row in range(4, sheet.max_row + 2):
            cell = sheet.cell(row=row, column=col)
            try:
                float(cell.value)
                cell.number_format = '_-₹* #,##0.00_-'
            except (TypeError, ValueError):
                pass

    # Label the total row
    sheet.cell(row=sheet.max_row, column=1, value="Total")
    sheet.cell(row=sheet.max_row, column=1).font = Font(bold=True)


@app.post("/uploadfile/")
async def upload_file(file: UploadFile = File(...)):
    """
    Endpoint to upload Excel files
    
    Args:
        file (UploadFile): Uploaded Excel file
    
    Returns:
        dict: Upload status message
    """
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        
        # Generate unique filename to prevent overwriting
        unique_filename = f"{uuid.uuid4()}_{file.filename}"
        file_location = os.path.join(UPLOAD_DIR, unique_filename)
        
        # Save the file
        with open(file_location, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        return {
            "message": f"File uploaded successfully", 
            "filename": unique_filename
        }
    except Exception as e:
        raise HTTPException(
            status_code=500, 
            detail=f"Error uploading file: {str(e)}"
        )

@app.post("/generate-report/")
async def generate_report_endpoint(file: UploadFile = File(...), month: str = "January"):
    """
    Endpoint to generate enhanced Excel report
    
    Args:
        file (UploadFile): Uploaded Excel file
        month (str, optional): Month for the report. Defaults to "January".
    
    Returns:
        FileResponse: Generated report file
    """
    try:
        print(f"[DEBUG] Endpoint '/generate-report/' hit successfully", flush=True)
        
        # Read file into memory
        input_file = BytesIO(await file.read())
        print(f"[DEBUG] File received: {file.filename}", flush=True)
        
        # Generate report
        report_file_path = generate_report(input_file, month)
        print(f"[DEBUG] Report generated at path: {report_file_path}", flush=True)
        
        # Return file for download
        return FileResponse(
            report_file_path, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(report_file_path)
        )
    except Exception as e:
        print(f"[ERROR] Error in '/generate-report/': {str(e)}", flush=True)
        raise HTTPException(
            status_code=500, 
            detail=f"Error generating report: {str(e)}"
        )

@app.get("/download-report/{report_file}")
async def download_report(report_file: str):
    """
    Endpoint to download previously generated reports
    
    Args:
        report_file (str): Filename of the report to download
    
    Returns:
        FileResponse: Requested report file
    """
    try:
        print(f"[DEBUG] Endpoint '/download-report/' hit successfully", flush=True)
        
        # Construct full file path
        file_path = os.path.join(UPLOAD_DIR, report_file)
        print(f"[DEBUG] Looking for file at: {file_path}", flush=True)
        
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"[DEBUG] File not found: {file_path}", flush=True)
            raise HTTPException(status_code=404, detail="File not found")
        
        print(f"[DEBUG] Returning file: {file_path}", flush=True)
        return FileResponse(
            file_path, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(file_path)
        )
    except Exception as e:
        print(f"[ERROR] Error in '/download-report/': {str(e)}", flush=True)
        raise HTTPException(
            status_code=500, 
            detail=f"Error downloading report: {str(e)}"
        )
