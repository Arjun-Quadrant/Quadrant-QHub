from Services import tableau_connection_info_extractor
from Services import tableau_visualizationInfo_extractor
from pathlib import Path

def main():
    # Define file paths
    
    base_path = r"C:\Users\arjun\Repos\Quadrant-QHub\DATA-HUB\Tableau Analysis"
    connection_info_json_file_path = base_path + "\connection info.json"
    connection_info_excel_file_path = base_path + "\connection info.xlsx"
    visualization_info_json_file_path = base_path + "\visualization info.json"
    visualization_info_excel_file_path = base_path + "\visualization info.xlsx"
    workbook_file_path = r"C:\Users\arjun\Repos\Quadrant-QHub\DATA-HUB\Workbooks\World Wide Importers Analysis.twb"
    packaged_workbook_file_path = workbook_file_path + "x"
    try:
        # Step 1: Extract data source information 
        data_source_info = tableau_connection_info_extractor.extract_data_source_info(workbook_file_path)
        print(data_source_info.connections)
        tableau_connection_info_extractor.save_connection_info_to_json_and_excel(data_source_info, connection_info_json_file_path, connection_info_excel_file_path)
        # Step 2: Extract visualization metadata 
        visualization_info = tableau_visualizationInfo_extractor.extract_viz_info(workbook_file_path, packaged_workbook_file_path)
        tableau_visualizationInfo_extractor.save_viz_info_to_json_and_excel(visualization_info, visualization_info_json_file_path, visualization_info_excel_file_path)
    except Exception as ex:
        print(f"Error: {ex}")

if __name__ == "__main__":
    main()