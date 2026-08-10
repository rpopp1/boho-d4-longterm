from config_map import CONFIG_FIELDS, CONFIG_MONTH_COLUMNS, config_cell, constant_map_for_month, matrices_only_constant_map_for_month


def test_config_month_columns():
    assert CONFIG_MONTH_COLUMNS["January"] == "B"
    assert CONFIG_MONTH_COLUMNS["December"] == "I"


def test_config_cell():
    assert config_cell("January", 14) == "B14"
    assert config_cell("August", 14) == "F14"


def test_config_fields_matches_original_layout_rows():
    # Same fields/rows/defaults as the single-month CONFIG_LAYOUT, just
    # without a hardcoded single-column cell address.
    by_name = {name: (row, default) for name, row, default in CONFIG_FIELDS}
    assert by_name["ZThreshold_Tier1"] == (14, 2.25)
    assert by_name["ContractLabel"] == (2, "TEMPLATE")
    assert by_name["StopLossUSD_Base"] == (27, -65000)


def test_constant_map_for_month_points_at_that_months_config_column():
    jan_map = constant_map_for_month("January")
    aug_map = constant_map_for_month("August")
    assert jan_map["$DV$6"] == "Config!B17"       # ZExitThreshold, January column B
    assert aug_map["$DV$6"] == "Config!F17"        # ZExitThreshold, August column F
    assert jan_map["$DK$5"] == "Config!B7"         # HOGalPerContract


def test_matrices_only_constant_map_for_month():
    assert matrices_only_constant_map_for_month("January") == {"$B$3": "Config!B11"}
    assert matrices_only_constant_map_for_month("December") == {"$B$3": "Config!I11"}
