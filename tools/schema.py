"""Canonical schema for the CCPI workbook.

Single source of truth for column names, controlled vocabularies, numeric
ranges and natural keys. Both the validator (append tool) and the documentation
reference this. Derived directly from the actual uploaded workbook structure.
"""

TAB1 = "CCPI Data"
TAB2 = "CCPI Subgroup Breakdown"

# --- exact column orders (must match the workbook) -------------------------
TAB1_COLS = [
    "Source File", "Base Year Used", "Year", "Month Number", "Month", "Date",
    "CCPI Index", "CCPI Core Index", "CCPI MoM Change", "CCPI Core MoM Change",
    "CCPI YoY Inflation", "CCPI Core YoY Inflation",
    "CCPI 12M Moving Avg Inflation", "CCPI Core 12M Moving Avg Inflation",
]
TAB2_COLS = [
    "Source File", "Base Year Used", "H/H Size", "Year", "Month", "Month No",
    "Period Type", "Group", "Base Value (Rs. Cts)", "Weight (%)", "Index Number",
    "Month-to-Month Inflation (%)", "Year-on-Year Inflation (%)",
    "Annual Average Inflation (%)",
]

# --- natural keys (uniqueness) ---------------------------------------------
TAB1_KEY = ["Base Year Used", "Year", "Month Number"]
TAB2_KEY = ["Base Year Used", "Year", "Month No", "Period Type", "Group"]

# --- controlled vocabularies (reject anything outside these) ---------------
# NOTE: Tab1 stores inflation as DECIMAL fractions (0.022 = 2.2%);
#       Tab2 stores inflation as PERCENT (10.3 = 10.3%). Never normalise on append.
BASE_YEARS_TAB1 = {"1952=100", "2002=100", "2006/07=100", "2013=100", "2021=100"}
BASE_YEARS_TAB2 = {"2002=100", "2006/07=100", "2013=100", "2021=100"}

GROUPS_TAB2 = {
    "All Items", "Food and Non-Alcoholic Beverages", "Non Food",
    "Alcoholic Beverages and Tobacco", "Clothing and Footwear",
    "Housing, Water, Electricity, Gas and Other Fuels",
    "Furnishing, Household Equipment and Routine Maintenance", "Health",
    "Transport", "Communication", "Recreation and Culture", "Education",
    "Restaurants and Hotels", "Miscellaneous Goods and Services",
}
PERIOD_TYPES = {"Monthly", "Annual"}
MONTHS = {"January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"}

# Groups that are aggregates, not components (exclude from breakdown to avoid
# double counting). "All Items" = total; "Non Food" = sum of non-food divisions.
AGGREGATE_GROUPS = {"All Items", "Non Food"}

# --- numeric sanity ranges (hard reject outside; structural gate) ----------
# Broad enough not to false-reject real history, tight enough to catch errors.
RANGES = {
    "Year": (1950, 2100),
    "Month Number": (1, 12),
    "Month No": (1, 12),
    "H/H Size": (1.0, 12.0),
    "Weight (%)": (0.0, 100.0),
    "CCPI Index": (10.0, 10000.0),
    "CCPI Core Index": (10.0, 10000.0),
    "Index Number": (10.0, 10000.0),
    "Base Value (Rs. Cts)": (0.0, 1_000_000.0),
    # Tab1 inflation = decimal fraction
    "CCPI MoM Change": (-0.9, 5.0),
    "CCPI Core MoM Change": (-0.9, 5.0),
    "CCPI YoY Inflation": (-0.9, 5.0),
    "CCPI Core YoY Inflation": (-0.9, 5.0),
    "CCPI 12M Moving Avg Inflation": (-0.9, 5.0),
    "CCPI Core 12M Moving Avg Inflation": (-0.9, 5.0),
    # Tab2 inflation = percent
    "Month-to-Month Inflation (%)": (-90.0, 500.0),
    "Year-on-Year Inflation (%)": (-90.0, 500.0),
    "Annual Average Inflation (%)": (-90.0, 500.0),
}

# columns that must never be null (the keys + the core measure)
TAB1_REQUIRED = ["Base Year Used", "Year", "Month Number", "Month", "CCPI Index"]
TAB2_REQUIRED = ["Base Year Used", "Year", "Month No", "Period Type", "Group", "Index Number"]

SPEC = {
    TAB1: dict(cols=TAB1_COLS, key=TAB1_KEY, bases=BASE_YEARS_TAB1,
              required=TAB1_REQUIRED),
    TAB2: dict(cols=TAB2_COLS, key=TAB2_KEY, bases=BASE_YEARS_TAB2,
              required=TAB2_REQUIRED),
}

# --- security limits -------------------------------------------------------
MAX_FILE_BYTES = 25 * 1024 * 1024      # 25 MB upload cap
MAX_INPUT_ROWS = 100_000               # row cap (zip-bomb / DoS guard)
ALLOWED_INPUT_EXT = {".xlsx", ".csv"}  # NO macro formats (.xlsm/.xltm/.xls)
