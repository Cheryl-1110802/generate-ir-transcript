# Core module for transcript and management report generation

from .data_parser import TranscriptParser, ManagementReportParser
from .report_generator import main as generate_reports
from .history_utils import check_if_history_already_updated, mark_history_as_updated
from .excel_table_to_word import ExcelTable2Word, create_excel_table2word, quick_paste_excel_to_word

__all__ = [
    'TranscriptParser',
    'ManagementReportParser', 
    'generate_reports',
    'check_if_history_already_updated',
    'mark_history_as_updated',
    'ExcelTable2Word',
    'create_excel_table2word',
    'quick_paste_excel_to_word'
]
