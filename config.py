# config.py
from pathlib import Path
from dataclasses import dataclass


@dataclass
class TestConfig:
    """Configuration class for test parameters."""
    csv_filename: str = None
    excel_filename: str = None
    ShowYearByDivision_graph: str = None

    # Directory paths
    BASE_DIR: Path = Path(__file__).parent.resolve()
    TESTS_DIR: Path = BASE_DIR / 'tests'
    TEST_DATA_DIR: Path = TESTS_DIR / 'TestData'
    VBA_CODE_DIR: Path = TESTS_DIR / 'VBA_code'
    
    # Excel configurations
    SHEET_NAME: str = '2024 wk1 1-1'
    STATISTICS_MAIN: str = 'Statistics.Statistics_Main'
    CALCULATE_DIVISION_COUNTS: str = 'Statistics.CalculateDivisionCounts'
    GLOBAL_PERSONS_LIST: str = 'GetGlobalPersonsYearList'
    CHART_LEFT: int = 100
    CHART_TOP: int = 150
    YTAE: str = 'YTAE'
    YDMF: str = 'YDMF'
    YEMB: str = 'YEMB'
    YTE: str = 'YTE'
    YHWE: str = 'YHWE'
    STATISTICS_BAS: str = 'Statistics.bas'
    TESTS_BAS: str = 'Tests.bas'
    CHARTS_BAS: str = 'charts.bas'
    UTILS_BAS: str =  'utils.bas'
    INTERVAL_CLS: str = 'clsInterval.cls'
    PERSON_CLS: str ='clsPerson.cls'

    # Test constants
    DIVISION_EMB: str = 'EMB'
    DIVISION_UNIT: str = 'Unit'
    STATISTICS: str = 'Statistics'
    TESTS: str = 'Tests'
    CHARTS: str = 'Charts'
    UTILS: str = 'Utils'
    TAE: str = 'TAE'
    DMF: str = 'DMF'
    EMB: str = 'EMB'
    TE: str = 'TE'
    HWE: str = 'HWE'

    # Method names
    INITIALIZE_METHOD: str = 'initialize'
    STATISTICSMAIN_METHOD: str = 'run_statistics_main'
    SHOWYEARBYDIVISION_METHOD: str = 'run_show_year_by_division'
    GETDIVISIONDATA_METHOD: str = 'run_get_division_data'
    UPLOADCSVTOEXCEL_METHOD: str = 'upload_csv_to_excel'
    GETEXPECTEDVALUESFROMEXCEL_METHOD: str = 'get_expected_values_from_excel'
    INJECTVBACODE_METHOD: str = 'inject_vba_code'
    CLOSEWORKBOOK_METHOD: str = 'close_workbook'

    # Other configurations
    OPENPYXL: str = 'openpyxl'
    REPLACE: str = 'replace'
    KEEP_VBA: str = 'keep_vba'
    APPEND_MODE: str = 'a'