from app.engine.formula_adjuster import FormulaAdjuster

def test_adjust_formula_repeated_row():
    # Template formula on row 5: =B5*C5
    # Should become =B8*C8 on target row 8
    mapping = {5: [6, 7, 8]}
    f = FormulaAdjuster.adjust_formula(
        formula="=B5*C5",
        orig_row=5,
        target_row=8,
        row_mapping=mapping
    )
    assert f == "=B8*C8"

def test_expand_summary_formula():
    # Template detail row 3 expanded to target rows [2, 3, 4]
    # Summary formula on template row 5: =SUM(D3:D3)
    # Target row is 5
    mapping = {
        1: [1],
        3: [2, 3, 4],
        5: [5]
    }
    f = FormulaAdjuster.adjust_formula(
        formula="=SUM(D3:D3)",
        orig_row=5,
        target_row=5,
        row_mapping=mapping
    )
    assert f == "=SUM(D2:D4)"
