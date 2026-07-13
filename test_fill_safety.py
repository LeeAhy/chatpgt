from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

from fill_sales import (
    build_code_lookup,
    match_owner_fuzzy_code,
    match_prediction_row_code,
    process_sales_workbooks,
)


def test_owner_fuzzy_matching_is_unique_and_exact_first() -> None:
    lookup, norms = build_code_lookup(["E1A20009A-416"])
    assert match_owner_fuzzy_code("E1A20009A-426", "王永仁", lookup, norms) == "E1A20009A-416"

    lookup, norms = build_code_lookup(["E1A20009A-416", "E1A20009A-417"])
    assert match_owner_fuzzy_code("E1A20009A-426", "王永仁", lookup, norms) is None

    lookup, norms = build_code_lookup(["E1A20009A-416", "E1A20009A-426"])
    assert (
        match_prediction_row_code(
            ["E1A20009A-426"],
            [0],
            "王永仁",
            lookup,
            norms,
        )
        == "E1A20009A-426"
    )

    lookup, norms = build_code_lookup(["C5PS07 Holder+BG半成品", "M9XZ11"])
    assert (
        match_owner_fuzzy_code("C5PSO7 Holder+BG半成品", "周文龙", lookup, norms)
        == "C5PS07 Holder+BG半成品"
    )


def test_unmatched_rows_stay_blank_and_only_forecast_cells_change(tmp_path: Path) -> None:
    sales_path = tmp_path / "sales.xlsx"
    prediction_path = tmp_path / "prediction.xlsx"
    output_path = tmp_path / "output.xlsx"

    sales_wb = Workbook()
    sales_ws = sales_wb.active
    sales_ws.title = "26年7月"
    sales_ws["E3"] = "业务担当"
    sales_ws["N3"] = "客户机种"
    sales_ws["T3"] = "6/29预估（7月）"
    sales_ws["T4"] = "数量"
    sales_ws["U4"] = "金额"
    sales_ws["V3"] = "7/13预估（7月）"
    sales_ws["V4"] = "数量"
    sales_ws["W4"] = "金额"
    sales_ws["E5"] = "王永仁"
    sales_ws["N5"] = "E1A20009A-416"
    sales_ws["T5"] = 10
    sales_ws["U5"] = 20
    sales_ws["E6"] = "王永仁"
    sales_ws["N6"] = "UNMATCHED-MODEL"
    sales_ws["A5"] = "必须保持不变"
    sales_wb.save(sales_path)

    prediction_wb = Workbook()
    prediction_ws = prediction_wb.active
    prediction_ws["A1"] = "子件描述"
    prediction_ws["B1"] = "7月"
    prediction_ws["A2"] = "E1A20009A-426"
    prediction_ws["B2"] = 20000
    prediction_wb.save(prediction_path)

    before_wb = load_workbook(sales_path, data_only=False)
    before = {
        (ws.title, cell.coordinate): (cell.value, cell.number_format)
        for ws in before_wb.worksheets
        for row in ws.iter_rows()
        for cell in row
    }
    before_wb.close()

    summary = process_sales_workbooks(
        pred_path=prediction_path,
        sales_path=sales_path,
        output_path=output_path,
        freeze_formulas=False,
        business_owner="王永仁",
        as_of_date=date(2026, 7, 13),
    )

    result_wb = load_workbook(output_path, data_only=False)
    result_ws = result_wb["26年7月"]
    assert summary["updated_rows"] == 1
    assert summary["zero_filled_rows"] == []
    assert result_ws["V5"].value == 2
    assert result_ws["W5"].value == 4
    assert result_ws["V6"].value is None
    assert result_ws["W6"].value is None

    allowed_changes = {("26年7月", "V5"), ("26年7月", "W5")}
    for ws in result_wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                key = (ws.title, cell.coordinate)
                if key in allowed_changes:
                    continue
                assert (cell.value, cell.number_format) == before.get(key, (None, "General"))
    result_wb.close()
