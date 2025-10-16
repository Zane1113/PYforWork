import zipfile
import os
import shutil
import sys
import glob

# Add PyPDF2 for PDF merging
try:
    from PyPDF2 import PdfMerger
    pdf_merger_available = True
except ImportError:
    print("PyPDF2 not found. PDF merging will be skipped.")
    print("To enable PDF merging, install PyPDF2 with: pip install PyPDF2")
    pdf_merger_available = False

def check_permissions(directory):
    """Check if we have read/write permissions for the directory"""
    read_permission = os.access(directory, os.R_OK)
    write_permission = os.access(directory, os.W_OK)
    execute_permission = os.access(directory, os.X_OK)
    
    print(f"Permission check for {directory}:")
    print(f"- Read permission: {'Yes' if read_permission else 'No'}")
    print(f"- Write permission: {'Yes' if write_permission else 'No'}")
    print(f"- Execute permission: {'Yes' if execute_permission else 'No'}")
    
    return read_permission and write_permission and execute_permission

def combine_pdfs(directory):
    """Combine all PDF files in the directory into a single PDF"""
    print("\n--- Combining PDF files ---")
    
    # Find all PDF files
    pdf_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    
    if not pdf_files:
        print("No PDF files found to combine.")
        return
    
    print(f"Found {len(pdf_files)} PDF files to combine")
    
    # Sort PDF files by name for consistent ordering
    pdf_files.sort()
    
    # Create output file path
    output_path = os.path.join(directory, "merged.pdf")
    
    try:
        merger = PdfMerger()
        
        # Add each PDF to the merger
        for pdf in pdf_files:
            try:
                print(f"Adding: {os.path.basename(pdf)}")
                merger.append(pdf)
            except Exception as e:
                print(f"Error adding {pdf}: {str(e)}")
        
        # Write the combined PDF
        print(f"Writing combined PDF to: {output_path}")
        merger.write(output_path)
        merger.close()
        
        print(f"Successfully created combined PDF: {output_path}")
        print(f"Combined PDF contains content from {len(pdf_files)} files")
    except Exception as e:
        print(f"Error combining PDFs: {str(e)}")
        import traceback
        traceback.print_exc()

def batch_unzip(directory):
    print(f"Starting batch extraction in: {directory}")
    
    # Check if directory exists
    if not os.path.exists(directory):
        print(f"Error: Directory '{directory}' does not exist!")
        return
    
    # Check permissions
    if not check_permissions(directory):
        print(f"Error: Insufficient permissions for directory '{directory}'")
        print("Please ensure you have read and write permissions for this directory.")
        return
    
    # Try different methods to find zip files
    print("Searching for ZIP files using different methods:")
    
    # Method 1: os.listdir
    zip_files = [item for item in os.listdir(directory) if os.path.isfile(os.path.join(directory, item)) and item.lower().endswith('.zip')]
    print(f"Method 1 (os.listdir): Found {len(zip_files)} ZIP files")
    
    # Method 2: glob
    glob_pattern = os.path.join(directory, "*.zip")
    glob_files = glob.glob(glob_pattern)
    print(f"Method 2 (glob): Found {len(glob_files)} ZIP files at {glob_pattern}")
    
    # Method 3: case insensitive search
    all_files = os.listdir(directory)
    case_insensitive_zips = [f for f in all_files if f.lower().endswith('.zip')]
    print(f"Method 3 (case insensitive): Found {len(case_insensitive_zips)} ZIP files")
    print(f"All files in directory: {all_files[:10]}")
    
    # Use the method that found files
    if len(zip_files) > 0:
        print(f"Using method 1: {zip_files}")
    elif len(glob_files) > 0:
        zip_files = [os.path.basename(f) for f in glob_files]
        print(f"Using method 2: {zip_files}")
    elif len(case_insensitive_zips) > 0:
        zip_files = case_insensitive_zips
        print(f"Using method 3: {zip_files}")
    else:
        print(f"No ZIP files found in {directory}")
        return
    
    print(f"Found {len(zip_files)} ZIP files to extract: {zip_files}")
    
    # Create a single extraction directory for all files
    extract_path = os.path.join(directory, "combined_extraction")
    if os.path.exists(extract_path):
        print(f"Removing existing combined folder: {extract_path}")
        shutil.rmtree(extract_path)
    
    os.makedirs(extract_path, exist_ok=True)
    print(f"Created combined extraction directory: {extract_path}")
    
    total_files_extracted = 0
    
    # Extract all ZIP files to the same directory
    for item in zip_files:
        file_path = os.path.join(directory, item)
        try:
            print(f"\n--- Processing {file_path} ---")
            print(f"File exists: {os.path.exists(file_path)}")
            print(f"File size: {os.path.getsize(file_path)} bytes")
            
            # Test if file is a valid zip
            try:
                with zipfile.ZipFile(file_path, 'r') as test_zip:
                    is_valid = test_zip.testzip() is None
                    print(f"ZIP file is valid: {is_valid}")
                    if not is_valid:
                        print("ZIP file is corrupted, skipping...")
                        continue
            except zipfile.BadZipFile:
                print(f"Error: {file_path} is not a valid ZIP file, skipping...")
                continue
            
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                # Get list of files in the zip
                file_list = zip_ref.namelist()
                print(f"ZIP contains {len(file_list)} files/folders")
                if file_list:
                    print(f"First few files in ZIP: {file_list[:3]}")
                
                # Extract all files to the combined directory
                print(f"Extracting to combined directory: {extract_path}")
                zip_ref.extractall(extract_path)
                print("Extraction completed")
                
                # Count files extracted from this ZIP
                files_in_zip = len([f for f in file_list if not f.endswith('/')])
                total_files_extracted += files_in_zip
                print(f"Added {files_in_zip} files to the combined directory")
                    
        except Exception as e:
            print(f"Error extracting {item}: {str(e)}")
            print(f"Exception type: {type(e).__name__}")
            import traceback
            traceback.print_exc()
    
    # Verify the combined extraction
    all_extracted_files = []
    for root, dirs, files in os.walk(extract_path):
        for file in files:
            all_extracted_files.append(os.path.join(root, file))
    
    print(f"\nTotal files extracted to combined directory: {len(all_extracted_files)}")
    print(f"Expected files based on ZIP contents: {total_files_extracted}")
    
    if all_extracted_files:
        print(f"Sample of extracted files: {all_extracted_files[:5]}")
    
    # Combine PDFs if PyPDF2 is available
    if pdf_merger_available:
        combine_pdfs(extract_path)
    else:
        print("\nPDF merging skipped. Install PyPDF2 to enable this feature:")
        print("pip install PyPDF2")
    
    print("\nBatch extraction completed")
    print(f"All files have been extracted to: {extract_path}")

# 调用函数并传入目录路径
directory_path = '/Users/cenghongyi/Downloads/WB9'
print(f"Python version: {sys.version}")
batch_unzip(directory_path)
print(f"\nTo verify extraction, check the contents of: {os.path.join(directory_path, 'combined_extraction')}")
if pdf_merger_available:
    print(f"Combined PDF available at: {os.path.join(directory_path, 'combined_extraction', 'merged.pdf')}")
