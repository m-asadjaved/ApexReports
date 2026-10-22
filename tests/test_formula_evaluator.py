from openpyxl import Workbook
from app.engine.formula_evaluator import FormulaEvaluator

def test_counta_evaluation():
    wb = Workbook()
    ws = wb.active

    # Rows 9, 10, 11 contain employee names
    ws["A9"] = "CLARK"
    ws["A10"] = "KING"
    ws["A11"] = "MILLER"

    # Row 12 has COUNTA formula
    ws["A12"] = "=COUNTA(A9:A11)"

    cache = FormulaEvaluator.evaluate_worksheet(ws)
    assert cache[(12, 1)] == 3

def test_count_and_subtotal():
    wb = Workbook()
    ws = wb.active

    ws["B1"] = 10
    ws["B2"] = 20
    ws["B3"] = 30

    ws["B4"] = "=SUM(B1:B3)"
    ws["B5"] = "=AVERAGE(B1:B3)"
    ws["B6"] = "=COUNT(B1:B3)"
    ws["B7"] = "=ROUND(B5, 0)"

    cache = FormulaEvaluator.evaluate_worksheet(ws)
    assert cache[(4, 2)] == 60
    assert cache[(5, 2)] == 20
    assert cache[(6, 2)] == 3
    assert cache[(7, 2)] == 20
