import os
import pandas as pd
import xlsxwriter

pd.set_option("display.expand_frame_repr", False)


def read_assets(file_path):
    """
    Reads file from given path into dataframe
    :return: Dataframe containing assets with columns "Model Number", "Asset Description", "Segment Description",
    "Site Code", "Shop"
    """

    df = pd.read_excel(file_path, usecols=["Model Number",
                                           "Asset Description",
                                           "Segment Description",
                                           "Site Code",
                                           "Shop"])
    return df


def add_count_col(df):
    """
    Groups given df by all fields and creates a new "Count" column indicating the number of assets in each group
    :param df: Dataframe containing asset details
    :return: Dataframe with identical asset/model/site combination grouped in single row and new "Count" column
             indicating the number of such assets
    """

    return df.\
        groupby(["Model Number", "Asset Description", "Site Code", "Shop", "Segment Description"]).\
        size().\
        to_frame("Count").\
        reset_index()


def create_dict(df):
    """
    Converts given dataframe into a dictionary; see :return: for dictionary details
    :param df: Dataframe with asset details and "Count" column
    :return: Dictionary with key: "Model Number"
                           value: [["Model Number", "Asset Description", ...], ["Model Number", ...], ...]

             Note: Value is a 2D list that contains unique entries for a given model number, if they exist
                   For example, model number "VC150" may have been purchased for two different sites "MSJ" and "SPH"
                   In this case, there will be two entries in the 2D list stored at key "VC150" in the dictionary
    """

    assets_dict = {}

    for index, row in df.iterrows():
        if row["Model Number"] in assets_dict:
            assets_dict.get(row["Model Number"]).append(row.tolist())
        else:
            assets_dict[row["Model Number"]] = []
            assets_dict.get(row["Model Number"]).append(row.tolist())
    return assets_dict


def update_count(new_assets_dict, retired_assets_dict):
    """
    Iterates through each asset in new_assets_dict and decreases its "Count" if there is a corresponding asset in the
    retired_assets_dict. After "Count" is updated, the corresponding asset is removed from retired_assets_dict.
    :param: new_assets_dict: Dictionary containing newly accepted assets
    :param: retired_assets_dict: Dictionary containing retired assets
    :return: None
    """

    # Iterate through each entry in the new assets dictionary
    for key, val in new_assets_dict.items():
        # Check to see if the newly purchased asset was also retired (key is model number)
        if key in retired_assets_dict:
            # Find the exact asset match by iterating through the dictionary value (2D list of asset details)
            for retired_asset in retired_assets_dict.get(key):
                for new_asset in new_assets_dict.get(key):
                    # Index 1 gives asset description, index 2 gives site, and index 3 gives shop code
                    if new_asset[1] == retired_asset[1] and \
                            new_asset[2] == retired_asset[2] and \
                            new_asset[3] == retired_asset[3]:
                        # Decrease new asset "Count" by retired asset "Count" (add since retired asset "Count" is neg)
                        new_asset[5] += retired_asset[5]
                        # Remove retired asset from retired_assets_dict after count in new_assets_dict has been updated
                        retired_assets_dict.get(key).remove(retired_asset)
                # Delete key from retired_assets_dict if there are no more assets in its 2D list
                if len(retired_assets_dict.get(key)) == 0:
                    del retired_assets_dict[key]


def merge_dict(new_assets_dict, retired_assets_dict):
    """
    Merges retired_assets_dict into new_assets_dict
    :param new_assets_dict: Dictionary containing new assets with updated "Count"
    :param retired_assets_dict: Dictionary containing retired assets that did not have a corresponding asset in
           new_assets_dict
    :return: None
    """

    # If retired_assets_dict isn't empty, merge it with new_assets_dict
    if len(retired_assets_dict) > 0:
        for key, val in retired_assets_dict.items():
            if key not in new_assets_dict:
                new_assets_dict[key] = val
            else:
                for asset_list in val:
                    new_assets_dict.get(key).append(asset_list)


def write_to_excel(new_assets_dict):
    """
    Writes given dictionary to Excel in format that can be accepted by biomed-service-delivery-cost-model
    :param new_assets_dict: Dict containing information about new and retired assets
    :return: None
    """

    # Initialization
    net_new_file_path = r"{dir_path}\output\net_new_assets.xlsx".format(dir_path=os.getcwd())
    workbook = xlsxwriter.Workbook(net_new_file_path)
    worksheet = workbook.add_worksheet("Net New Assets")

    # Formatting
    heading = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#244062"})

    # Write headers
    headers = ["model_num", "asset_description", "quantity", "health_auth", "site_code", "shop_code"]
    worksheet.write_row(0, 0, headers, heading)

    # Write asset details
    row = 1
    col = 0

    for key, val in new_assets_dict.items():

        for asset_list in val:

            row_data = [asset_list[0],   # model_num
                        asset_list[1],   # asset_description
                        asset_list[5],   # quantity
                        asset_list[4],   # health_auth
                        asset_list[2],   # site_code
                        asset_list[3]]   # shop_code

            worksheet.write_row(row, col, row_data)
            row += 1

    # Set column width
    worksheet.set_column(0, 0, 15)   # A: model_num
    worksheet.set_column(1, 1, 70)   # B: asset_description
    worksheet.set_column(2, 2, 10)   # C: quantity
    worksheet.set_column(3, 3, 13)   # D: health_auth
    worksheet.set_column(4, 5, 10)   # E, F: site_code, shop_code

    workbook.close()


def main():

    print("Generating list of net new assets...")

    # Get name of TMS exports of new and retired assets
    new_assets_path = "new_assets/{file}".format(file=os.listdir("new_assets")[0])
    retired_assets_path = "retired_assets/{file}".format(file=os.listdir("retired_assets")[0])

    # Read TMS exports for new and retired assets into separate dataframes
    new_assets_df = read_assets(new_assets_path)
    retired_assets_df = read_assets(retired_assets_path)

    # Group df by rows so that we have one row for each unique asset/model/site combination, add count column
    new_assets_df = add_count_col(new_assets_df)
    retired_assets_df = add_count_col(retired_assets_df)

    # Convert all "Count" values in the retired df to a negative value
    retired_assets_df["Count"] *= -1

    # Convert df to dictionary with key: "Model Number" and value: [["Asset 1 details"], ["Asset 2 details"], ...]
    # Note: Assets with same model but different sites are grouped separately, hence the 2D list
    new_assets_dict = create_dict(new_assets_df)
    retired_assets_dict = create_dict(retired_assets_df)

    # If there is an exact asset match between the two dicts, decrease new assets "Count" by retired assets "Count"
    # After "Count" has been updated, retired asset is removed from its dictionary
    # This will give us the count of net new assets at each site in new_assets_dict
    update_count(new_assets_dict, retired_assets_dict)

    # Merge retired_assets_dict (only retired assets without corresponding new asset remain) into new_assets_dict
    merge_dict(new_assets_dict, retired_assets_dict)

    # Write new_assets_dict to Excel
    write_to_excel(new_assets_dict)

    input("Net new assets list successfully generated. Press 'Enter' to close this window.")


if __name__ == "__main__":
    main()