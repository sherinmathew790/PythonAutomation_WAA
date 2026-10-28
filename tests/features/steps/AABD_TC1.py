# AABD_TC1.py
from behave import given, when, then
from common_test_functions import CommonTestFunctions
from config import TestConfig
from logger_config import TestLogger
from typing import Dict, Any, Optional
from dataclasses import dataclass

@dataclass
class TestContext:
    """Class to store test context data."""
    common_functions: Optional[CommonTestFunctions] = None
    persons_year_list: Any = None
    emb_value: Optional[int] = None

class AABDTestCase:
    """Test case implementation for AABD scenarios."""

    def __init__(self):
        """Initialize test case with configuration and logging setup."""
        self.config = TestConfig()
        self.logger = TestLogger().get_logger()
        self.context = TestContext()
        self.context.common_functions = CommonTestFunctions()
        self.csv_filename = None
        self.excel_filename = None
        self.ShowYearByDivision_graph = None
        self.logger.info("AABD Test Case initialized")

    @TestLogger.log_method
    def setup_test_environment(self) -> None:
        """Set up the test environment with required data and configurations."""       
        self.context.common_functions.initialize(self.config.excel_filename)
        self.logger.info("Injecting VBA code...")
        self.context.common_functions.inject_vba_code()
        self.logger.info("VBA code injection completed")

    @TestLogger.log_method
    def validate_division_data(self, values: Dict[str, int], expected_values: Dict[str, int]) -> None:
        """Validate division data against expected values.
        Args:
            values: Dictionary containing actual division values
            expected_values: Dictionary containing expected division values"""
        # Validate individual divisions
        for division, expected_count in expected_values.items():
            actual_count = values.get(division)
        assert actual_count == expected_count, (
                    f"Division count mismatch for {division}: "
                    f"expected {expected_count}, got {actual_count}"
                )
        self.logger.info(f"Validation passed for {division}: {expected_count}")
        # Validate total employees
        total_employees = sum(values.values())
        expected_total = sum(expected_values.values())
        assert total_employees == expected_total, (
            f"Total employees mismatch: "
            f"expected {expected_total}, got {total_employees}"
        )
        self.logger.info(f"Total employees validation passed: {expected_total}")

# Behave step implementations
@given('dataset "{csv_filename}"')
def step_impl(context, csv_filename):
    """Initialize the testcase"""
    context.test_case = AABDTestCase()
    context.test_case.config.csv_filename = csv_filename

@when('the dataset is injected into "{excel_filename}"')
def step_impl(context, excel_filename):
    """Upload test data from CSV to Excel."""
    context.test_case.config.excel_filename = excel_filename
    context.test_case.context.common_functions.upload_csv_to_excel(
            csv_file=context.test_case.config.csv_filename,
            excel_file=context.test_case.config.excel_filename,
            sheet_name=context.test_case.config.SHEET_NAME
        )   
    """Initialize the test environment"""
    context.test_case.setup_test_environment()
    
@when('all charts are visualized')
def step_impl(context):
    """Execute statistics main macro."""
    context.test_case.context.common_functions.initialize(context.test_case.config.excel_filename)
    context.test_case.context.persons_year_list = (
        context.test_case.context.common_functions.run_statistics_main()
    )  

@then('"{ShowYearByDivision_graph}" should correctly display the number of employees per division')
def step_impl(context, ShowYearByDivision_graph):
    """Generate annual availability graph."""
    context.test_case.config.ShowYearByDivision_graph = ShowYearByDivision_graph
    context.test_case.context.common_functions.run_show_year_by_division(
        context.test_case.context.persons_year_list,context.test_case.config.ShowYearByDivision_graph
    )

    """Validate division data display."""
    # Get actual values
    values = context.test_case.context.common_functions.run_get_division_data(
        context.test_case.context.persons_year_list
    )
    # Get expected values
    column_values = context.test_case.context.common_functions.get_expected_values_from_excel(
        context.test_case.config.DIVISION_UNIT,context.test_case.config.excel_filename
    )
    # Validate data
    context.test_case.validate_division_data(values, column_values)
    context.test_case.context.common_functions.close_workbook()