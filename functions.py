import pandas as pd
from datetime import timedelta


def build_report(self, path_to_order_status_report, path_to_supply_line_template):

    sheet_name = "Direct FG Work Orders"
    columns = ['Customer Part', 'Assembly ID', 'Order-Rel', 'Start Date', 'Est Finish Date', 'Quantity', 'Type']

    old_row_count =  pd.read_excel(path_to_supply_line_template, sheet_name=1, skiprows=16).shape[0]

    # Read Web & customer part numbers (using customer report template for lookup)
    part_numbers = pd.read_excel(path_to_supply_line_template, sheet_name=0, skiprows=13)
    part_numbers = part_numbers.drop(
        columns=[
            "Quantity_On_Hand",
            "Lead_Time",
        ]
    ).set_index("Supplier_Item_Code")

    # Read order Statuses, drop uneccessary columns, and rename others to customer convention
    order_status = pd.read_excel(
        path_to_order_status_report,
        dtype={"Order No": str, "Rel No": int, "Quantity": int},
    )

    # Format order dates
    order_status["Est Finish Date"] = pd.to_datetime(
        order_status["Est Finish Date"]
    ).dt.date

    order_status["Start Date"] = order_status["Est Finish Date"] - timedelta(
        days=7
    )

    # Format work order as: production order - release
    order_status["Rel No"] = order_status["Rel No"].map("{:03.0f}".format)
    order_status["Order-Rel"] = order_status[["Order No", "Rel No"]].agg(
        "-".join, axis=1
    )

    # Lookup customer part number using Web part number
    order_status["Customer Part"] = order_status["Assembly ID"].map(
        part_numbers["FG_Item_Code"]
    )

    # Order status / type
    order_status["Type"] = order_status["Status"].map(
        {"Firm Planned": "PLANNED", "In Process": "WIP", "Completed": "WIP"}
    )

    # Remove unnececssary columns
    order_status = order_status.drop(
        columns=[
            "Order No",
            "Rel No",
            "Customer ID",
            "Revision Number",
            "Sales Order",
            "Customer PO Number",
            "Location ID",
            "Planner",
            "Status",
            "Est Start Date",
        ]
    )

    order_status = order_status[columns] # Reorder columns

    # Append blank rows to end of new data if old data is longer than new data
    new_row_count = order_status.shape[0]

    if new_row_count < old_row_count:
        print(f'New data is {old_row_count - new_row_count} smaller than old data, appending blank rows')
        blank_row = ["" for x in [i for i in range(0,7)]]
        for i in range(0, old_row_count - new_row_count):
            order_status.loc[order_status.shape[0]] = blank_row

    ## Save new report into the provided template
    with pd.ExcelWriter(
        path_to_supply_line_template,
        mode="a",
        if_sheet_exists="overlay",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        order_status.to_excel(
            writer,
            sheet_name="Direct FG Work Orders",
            startrow=17,
            header=False,
            index=False,
        )
