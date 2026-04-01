import csv
from fastapi import FastAPI, HTTPException

app = FastAPI()

def transform_attendance(file_path):
    try:
        with open(file_path, mode='r', encoding='utf-8') as f:
            reader = list(csv.reader(f))
            
            if len(reader) < 3:
                return []

            # Row 0 contains the dates
            header_dates = reader[0]
            # Row 2 onwards contains the actual data
            data_rows = reader[2:] 
            
            employees_json = []

            for row in data_rows:
                # Basic validation: Skip empty rows or rows that aren't employee data
                if not row or len(row) < 4 or not row[1].strip().isdigit():
                    continue
                    
                emp_data = {
                    "employee_id": float(row[1]),
                    "employee_name": row[2],
                    "department": row[3],
                    "attendance": {}
                }
                
                # Attendance data columns start at index 4
                col_idx = 4
                # Iterate through the header dates row
                for i in range(4, len(header_dates)):
                    date_label = header_dates[i].strip()
                    
                    # Only process when we hit a column with a date string
                    if date_label and "-" in date_label:
                        # Extract the triplet: IN, OUT, HRS
                        val_in = row[col_idx] if col_idx < len(row) else "-"
                        val_out = row[col_idx + 1] if (col_idx + 1) < len(row) else "-"
                        val_hrs = row[col_idx + 2] if (col_idx + 2) < len(row) else "-"
                        
                        # Apply your formatting rules
                        if val_in == "Absent":
                            entry = {"in": "0:00", "out": "0:00", "hrs": "0:00"}
                        else:
                            entry = {
                                "in": val_in if val_in != "-" else None,
                                "out": val_out if val_out != "-" else None,
                                "hrs": val_hrs if val_hrs != "-" else None
                            }
                        
                        emp_data["attendance"][date_label] = entry
                        # Move to the next date block (3 columns ahead)
                        col_idx += 3
                
                employees_json.append(emp_data)
                
            return employees_json
    except FileNotFoundError:
        return None

# --- API Endpoints ---

@app.get("/")
def read_root():
    return {"message": "Attendance API is running"}

@app.get("/get-attendance")
def get_attendance():
    filename = 'attendance_2026-01-01_to_2026-01-30.csv'
    result = transform_attendance(filename)
    
    if result is None:
        raise HTTPException(status_code=404, detail=f"File {filename} not found")
    
    if not result:
        return {"message": "No valid data found in CSV"}

    # Return the first employee record directly as a dictionary
    # FastAPI will automatically handle the JSON conversion
    return result[0]

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)