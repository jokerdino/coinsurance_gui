import pandas as pd

def generate_pivot_table(excel_filename, company_name, path_string_wip):
    print(f"{company_name=}")
    try:
        df_claims = pd.read_excel(excel_filename, "CR", converters={"Follower Office Code":str, "Origin Office": str})
        df_pivot_claims = pd.pivot_table(
            df_claims,
            index=["Origin Office", "Follower Office Code"],
            values="Total claim amount",
            aggfunc="sum",
        )
    except ValueError:
        df_pivot_claims = pd.DataFrame()
        print("No claims sheet for this company")

    try:
        df_premium = pd.read_excel(excel_filename, "PP", converters={"Follower Office Code":str, "Origin Office": str})
        df_pivot_premium = pd.pivot_table(
            df_premium,
            index=["Origin Office", "Follower Office Code"],
            values="Net Premium payable",
            aggfunc="sum",
        )
    except ValueError:
        df_pivot_premium = pd.DataFrame()
        print("No premium sheet for this company")

    pd.set_option("display.float_format", "{:,.2f}".format)
    if not df_pivot_premium.empty:
        if not df_pivot_claims.empty:
            df_pivot_merged = pd.merge(df_pivot_premium, df_pivot_claims, left_index=True, right_index=True, how="outer")
            df_pivot_merged.reset_index(inplace=True)
            df_pivot_merged.fillna(0, inplace=True)
            net_amount = df_pivot_merged["Total claim amount"].sum() - df_pivot_merged["Net Premium payable"].sum()
            if net_amount > 0:
                df_pivot_merged["Net Receivable by UIIC"] = (
                    df_pivot_merged["Total claim amount"] - df_pivot_merged["Net Premium payable"]
                )
                df_pivot_merged.style.format({"Net Receivable by UIIC": "{0:,.2f}"})
            else:
                df_pivot_merged["Net Payable by UIIC"] = (
                    df_pivot_merged["Net Premium payable"] - df_pivot_merged["Total claim amount"]
                )
                df_pivot_merged.style.format({"Net Payable by UIIC": "{0:,.2f}"})
            df_pivot_merged.rename(
                {"Total claim amount": "Claims receivable"}, axis=1, inplace=True
            )
        else:
            df_pivot_merged = df_pivot_premium
            df_pivot_merged.reset_index(inplace=True)
    else:
        df_pivot_merged = df_pivot_claims
        df_pivot_merged.reset_index(inplace=True)

    df_pivot_merged.loc["Total"] = df_pivot_merged.sum(numeric_only=True, axis=0)

    print(df_pivot_merged)
    if path_string_wip:
        file_path = f"{path_string_wip}/{company_name}/{company_name}_summary.xlsx"
    else:
        file_path = f"{company_name}_summary.xlsx"
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        df_pivot_merged.to_excel(writer, sheet_name="Summary", index=False)
        format_workbook = writer.book

        format_currency = format_workbook.add_format({"num_format": "##,##,#0"})
        format_bold = format_workbook.add_format({"num_format": "##,##,#", "bold": True})
        format_header = format_workbook.add_format(
            {"bold": True, "text_wrap": True, "valign": "top"}
        )
        format_worksheet = writer.sheets["Summary"]

        format_worksheet.set_column("E:E", 12, format_bold)
        format_worksheet.set_column("C:D", 12, format_currency)
        format_worksheet.set_row(-1, 12, format_bold)
        format_worksheet.set_row(0, None, format_header)
