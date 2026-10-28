# common_test_functions.py
import win32com.client
import openpyxl
import pandas as pd
from typing import Dict, List, Any, Optional
from pathlib import Path
from config import TestConfig
from logger_config import TestLogger

class CommonTestFunctions:
    """Class to handle common Excel automation operations for testing."""
    
    def __init__(self):
        """Initialize CommonTestFunctions with configuration and logging setup.
        Sets up necessary configurations and logging for Excel operations."""
        self.config = TestConfig()
        self.logger = TestLogger().get_logger()
        self.excel: Optional[win32com.client.CDispatch] = None
        self.workbook: Optional[win32com.client.CDispatch] = None
        self.logger.info("CommonTestFunctions initialized successfully")

    @TestLogger.log_method
    def initialize(self, excel_file: str) -> None:
        """Initialize Excel application and open workbook."""
        self.logger.info("Initializing Excel application...")
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = False
        excel_path = self.config.TEST_DATA_DIR / excel_file
        self.logger.info(f"Opening workbook: {excel_path}")
        self.workbook = self.excel.Workbooks.Open(str(excel_path))

    @TestLogger.log_method
    def run_statistics_main(self) -> Any:
        """Execute Statistics_Main macro and return persons year list.
        Returns:
            Any: Result of the GetGlobalPersonsYearList macro"""
        self.logger.info("Running Statistics_Main macro...")
        self.excel.Application.Run(self.config.STATISTICS_MAIN)
        self.logger.info("Getting global persons year list...")
        persons_year_list = self.excel.Application.Run(self.config.GLOBAL_PERSONS_LIST)
        self.workbook.Save()
        return persons_year_list
    
    @TestLogger.log_method
    def run_show_year_by_division(self, persons_year_list: Any, ShowYearByDivision_graph: str ) -> None:
        """Execute ShowYearByDivision macro.
        Args:
            persons_year_list: List of persons by year"""
        self.logger.info("Executing ShowYearByDivision macro...")
        chart_sheet = self.workbook.Sheets(self.config.SHEET_NAME)
        #print(f"Graph Name: {ShowYearByDivision_graph}")
        self.excel.Application.Run(ShowYearByDivision_graph, chart_sheet, persons_year_list, self.config.CHART_LEFT, self.config.CHART_TOP)

    @TestLogger.log_method
    def run_get_division_data(self, persons_year_list: Any) -> Dict[str, int]:
        """Calculate division counts and return results.
        Args:
            persons_year_list: List of persons by year
        Returns:
            Dict[str, int]: Dictionary containing division counts"""
        self.logger.info("Calculating division counts...")
        division_counts = self.excel.Application.Run(self.config.CALCULATE_DIVISION_COUNTS, persons_year_list)
        results = {
            self.config.TAE: division_counts.Item(str(self.config.YTAE)),
            self.config.DMF: division_counts.Item(str(self.config.YDMF)),
            self.config.EMB: division_counts.Item(str(self.config.YEMB)),
            self.config.TE: division_counts.Item(str(self.config.YTE)),
            self.config.HWE: division_counts.Item(str(self.config.YHWE))
        }
        return results

    @TestLogger.log_method
    def upload_csv_to_excel(self, csv_file: str, excel_file: str, sheet_name: str) -> None:
        """Upload CSV data to Excel worksheet.
        Args:
            csv_file: Name of the CSV file
            excel_file: Name of the Excel file
            sheet_name: Name of the sheet to create/update"""
        csv_path = self.config.TEST_DATA_DIR / csv_file
        excel_path = self.config.TEST_DATA_DIR / excel_file
        self.logger.info(f"Reading CSV file: {csv_path}")
        df = pd.read_csv(csv_path, sep=';')
        df = df.rename(columns={col: '' for col in df.columns if 'Unnamed' in col})
        self.logger.info(f"Writing data to Excel: {excel_path}")
        with pd.ExcelWriter(excel_path, engine=self.config.OPENPYXL, mode=self.config.APPEND_MODE, if_sheet_exists=self.config.REPLACE, engine_kwargs={self.config.KEEP_VBA: True}) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    @TestLogger.log_method
    def get_expected_values_from_excel(self, column_name: str, excel_file: str) -> Dict[Any, int]:
        """Retrieve values from specified Excel column
        Args:
            column_name: Name of the column to read    
        Returns:
            Dict[Any, int]: Dictionary of values and their counts
        Raises:
            ValueError: If column not found"""
        self.logger.info(f"Reading values from column: {column_name}")
        excel_path = self.config.TEST_DATA_DIR / excel_file
        workbook = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = workbook[self.config.SHEET_NAME]
        header_row = next(sheet.rows)
        column_index = next((idx for idx, cell in enumerate(header_row) if cell.value == column_name), None)
       
        if column_index is None:
            raise ValueError(f"Column '{column_name}' not found")
        values = [row[column_index].value for row in list(sheet.rows)[1:] if row[column_index].value is not None]
        value_counts = {}
       
        for value in values:
            cleaned_value = ''.join(char for char in str(value) if not char.isdigit()).strip()
            value_counts[cleaned_value] = value_counts.get(cleaned_value, 0) + 1
       
        return value_counts

    @TestLogger.log_method
    def inject_vba_code(self) -> None:
        """Inject VBA code from external files into the workbook.
        If modules already exist, they will be removed and replaced with new ones."""
        self.logger.info("Starting VBA code injection...")
        vba_project = self.workbook.VBProject
        bas_files = {self.config.STATISTICS_BAS: self.config.STATISTICS, self.config.TESTS_BAS: self.config.TESTS, self.config.CHARTS_BAS: self.config.CHARTS, self.config.UTILS_BAS: self.config.UTILS}
        cls_files = [self.config.INTERVAL_CLS, self.config.PERSON_CLS]
        self.verify_vba_files(bas_files.keys(), cls_files)
        self.remove_existing_modules(vba_project, bas_files.values(), cls_files)
        self.add_standard_modules(vba_project, bas_files)
        self.add_class_modules(vba_project, cls_files)
        self.workbook.Save()

    def verify_vba_files(self, bas_files: List[str], cls_files: List[str]) -> None:
        """Verify that all required VBA files exist.
        Args:
            bas_files: List of .bas files
            cls_files: List of .cls files
        Raises:
            FileNotFoundError: If any required file is missing"""
        # Convert dict_keys to list before concatenation
        all_files = list(bas_files) + cls_files
        missing_files = [str(self.config.VBA_CODE_DIR / file) for file in all_files if not (self.config.VBA_CODE_DIR / file).exists()]
        if missing_files:
            raise FileNotFoundError(f"Missing VBA files:\n" + "\n".join(missing_files))

    def remove_existing_modules(self, vba_project, standard_modules: List[str], class_modules: List[str]) -> None:
        """Remove existing VBA modules from the workbook.
        Args:
            vba_project: Excel VBA project
            standard_modules: List of standard module names to remove
            class_modules: List of class module names to remove"""
        for component in list(vba_project.VBComponents):
            component_name = component.Name
            if component_name in standard_modules or any(component_name == cls_file.replace('.cls', '') for cls_file in class_modules):
                vba_project.VBComponents.Remove(component)

    def add_standard_modules(self, vba_project, bas_files: Dict[str, str]) -> None:
        """Add standard modules to the VBA project.
        Args:
            vba_project: Excel VBA project
            bas_files: Dictionary mapping .bas files to module names
        Raises:
            ValueError: If required file is empty"""
        for bas_file, module_name in bas_files.items():
            module = vba_project.VBComponents.Add(1)  # 1 = standard module
            bas_path = self.config.VBA_CODE_DIR / bas_file
            with open(bas_path, 'r', encoding='utf-8-sig') as file:
                code = file.read()
                if code.strip():
                    module.CodeModule.AddFromString(code)
                else:
                    raise ValueError(f"Empty code file: {bas_file}")
            module.Name = module_name

    def add_class_modules(self, vba_project, cls_files: List[str]) -> None:
        """Add class modules to the VBA project.
        Args:
            vba_project: Excel VBA project
            cls_files: List of .cls files to import
        Raises:
            FileNotFoundError: If any required file is missing"""
        for cls_file in cls_files:
            cls_path = self.config.VBA_CODE_DIR / cls_file
            self.logger.debug(f"Importing class module: {cls_file}")
                    
            if cls_path.exists():
                vba_project.VBComponents.Import(str(cls_path))
                self.logger.debug(f"Successfully imported: {cls_file}")
            else:
                raise FileNotFoundError(f"Class module file not found: {cls_file}")
    
    @TestLogger.log_method
    def close_workbook(self) -> None:
        """Close workbook and quit Excel application.
        Ensures proper cleanup of Excel resources."""
        self.logger.info("Closing workbook and quitting Excel...")
        if self.workbook:
            self.workbook.Close(False)
            self.logger.info("Workbook closed successfully")
        if self.excel:
            self.excel.Quit()
            self.logger.info("Excel application quit successfully")
        self.workbook = None
        self.excel = None