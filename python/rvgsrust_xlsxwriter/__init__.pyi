from typing import Optional
from ._core import (
    Chart as Chart,
    ChartDataLabel as ChartDataLabel,
    ChartFont as ChartFont,
    ChartFormat as ChartFormat,
    ChartMarker as ChartMarker,
    ChartSeries as ChartSeries,
    ChartTrendline as ChartTrendline,
    ConditionalFormat2ColorScale as ConditionalFormat2ColorScale,
    ConditionalFormat3ColorScale as ConditionalFormat3ColorScale,
    ConditionalFormatAverage as ConditionalFormatAverage,
    ConditionalFormatBlank as ConditionalFormatBlank,
    ConditionalFormatCell as ConditionalFormatCell,
    ConditionalFormatDataBar as ConditionalFormatDataBar,
    ConditionalFormatDate as ConditionalFormatDate,
    ConditionalFormatDuplicate as ConditionalFormatDuplicate,
    ConditionalFormatError as ConditionalFormatError,
    ConditionalFormatFormula as ConditionalFormatFormula,
    ConditionalFormatText as ConditionalFormatText,
    ConditionalFormatTop as ConditionalFormatTop,
    DataValidation as DataValidation,
    Format as Format,
    Sparkline as Sparkline,
    Table as Table,
    TableColumn as TableColumn,
    Worksheet as Worksheet,
    Workbook as _CoreWorkbook,
)

class Workbook(_CoreWorkbook):
    """Adds context-manager support and a close() that defaults to the
    path given to the constructor, on top of the _core Workbook.
    """
    path: Optional[str]
    def __new__(cls, path: Optional[str] = None) -> "Workbook": ...
    def __enter__(self) -> "Workbook": ...
    def __exit__(self, exc_type: object, exc_val: object, exc_tb: object) -> None: ...
    def close(self, path: Optional[str] = None) -> None: ...

__all__: list[str]

