"""Canonical column naming and alias resolution."""

# Canonical column names used throughout the codebase
REQUIRED_COLUMNS = {
    "AcctNo",
    "Total Items",
    "Paid Items",
    "Returned Items",
    "OD Limit",
    "OD Status",
    "Product Code",
    "Business Flag",
    "Account Status",
    "Reg E Flag",
    "Open Date",
    "Avg Bal",
    "Deposit Amount",
    "Deposit Count",
    "Swipes",
    "Spend",
}

# Maps raw column names (from various CSV formats) to canonical names.
# Merges aliases from both the script's RENAME_MAP and the notebook's to_canonical().
COLUMN_ALIASES: dict[str, str] = {
    # Script RENAME_MAP aliases
    "TOTALITEMS": "Total Items",
    "PaidItems": "Paid Items",
    "ReturnedItems": "Returned Items",
    "ODLimit": "OD Limit",
    "ODStatus": "OD Status",
    "ProdCode": "Product Code",
    "BusinessFlag": "Business Flag",
    "AccountStatus": "Account Status",
    "RegEValue": "Reg E Flag",
    "OpenDate": "Open Date",
    "AvgColBal": "Avg Bal",
    "DepositAmount": "Deposit Amount",
    "DepositCount": "Deposit Count",
    "swipes": "Swipes",
    "spend": "Spend",
    # Notebook to_canonical() aliases
    "ACCTNO": "AcctNo",
    "Acctno": "AcctNo",
    "acctno": "AcctNo",
    "TotalItems": "Total Items",
    "Total_Items": "Total Items",
    "Paid_Items": "Paid Items",
    "Returned_Items": "Returned Items",
    "OD_Limit": "OD Limit",
    "OD_Status": "OD Status",
    "Business_Flag": "Business Flag",
    "Account_Status": "Account Status",
    "Reg_E_Flag": "Reg E Flag",
    "RegE": "Reg E Flag",
    "Open_Date": "Open Date",
    "Avg_Bal": "Avg Bal",
    "Deposit_Amount": "Deposit Amount",
    "Deposit_Count": "Deposit Count",
    # Normalize column name variants
    "# of Paid Items": "# of Items Paid",
}


def resolve_columns(df):
    """Rename DataFrame columns to canonical names using the alias map.

    Returns a new DataFrame with renamed columns (does not mutate input).
    """
    rename_map = {}
    for col in df.columns:
        if col in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[col]
    return df.rename(columns=rename_map)
