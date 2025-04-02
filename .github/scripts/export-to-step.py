import sys
import os
import traceback
import time

def export_to_step(inventor_file_path):
    try:
        import win32com.client
        import pythoncom
        
        print(f"Starting export of: {inventor_file_path}")
        pythoncom.CoInitialize()
        
        # Connect to Inventor
        try:
            inventor_app = win32com.client.GetActiveObject('Inventor.Application')
            print("Connected to running Inventor instance")
        except:
            print("Starting new Inventor instance...")
            inventor_app = win32com.client.Dispatch('Inventor.Application')
            inventor_app.Visible = True
            print("Started new Inventor instance")
            time.sleep(2)
        
        # Open the document
        print(f"Opening document...")
        document = inventor_app.Documents.Open(inventor_file_path)
        print(f"Successfully opened: {os.path.basename(inventor_file_path)}")
        
        # Generate output path
        output_dir = os.path.dirname(inventor_file_path)
        base_name = os.path.splitext(os.path.basename(inventor_file_path))[0]
        step_dir = os.path.join(output_dir, "STEP_Exports")
        if not os.path.exists(step_dir):
            os.makedirs(step_dir)
        step_file_path = os.path.join(step_dir, f"{base_name}.step")
        
        print(f"Will save to: {step_file_path}")
        
        # Save as STEP file
        print("Starting export...")
        document.SaveAs(step_file_path, True)  # True means Save Copy
        print(f"SUCCESS: Exported to {step_file_path}")
        
        # Close the document
        document.Close(False)  # False = don't save changes
        print("Document closed")
        
        print("Export process complete!")
        
    except Exception as e:
        print(f"× ERROR: {str(e)}")
        print("Detailed traceback:")
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python export-to-step.py <inventor_file_path>")
        sys.exit(1)
    
    inventor_file_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(inventor_file_path):
        print(f"File not found: {inventor_file_path}")
        sys.exit(1)
    
    export_to_step(inventor_file_path)
    print("Script finished execution")