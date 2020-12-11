import os
import pandas as pd

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
    :return: Dataframe with identical entries grouped and a "Count" field indicating the number of entries
    """

    return df.\
        groupby(["Model Number", "Asset Description", "Site Code", "Shop", "Segment Description"]).\
        size().\
        to_frame("Count").\
        reset_index()


def create_dict(df):
    """
    Converts given dataframe into a dictionary; see :return: for dictionary details
    :param df: Dataframe with asset details grouped into unique entries and a count of the # in each group
    :return: Dictionary with key: "Model Number"
                           value: [["Model Number", "Asset Description", ...], ["Model Number", ...], ...]

             Note: Value is a 2D list that contains unique entries for a given model number, if they exist
                   For example, model number "VC150" may have been purchased for two different sites "MSJ" and "SPH"
                   In this case, there will be two entries in the 2D list stored at "VC150" in the dictionary
    """

    assets_dict = {}
    for index, row in df.iterrows():
        if row["Model Number"] in assets_dict:
            assets_dict.get(row["Model Number"]).append(row.tolist())
        else:
            assets_dict[row["Model Number"]] = []
            assets_dict.get(row["Model Number"]).append(row.tolist())
    return assets_dict


def main():

    # Get name of TMS exports of new and retired assets
    new_assets_path = "new_assets/{file}".format(file=os.listdir("new_assets")[0])
    retired_assets_path = "retired_assets/{file}".format(file=os.listdir("retired_assets")[0])

    # Read TMS exports for new and retired assets into separate dataframes
    new_assets_df = read_assets(new_assets_path)
    retired_assets_df = read_assets(retired_assets_path)

    # Group df by rows with perfect match, add count column indicating number of matching asset entries
    new_assets_df = add_count_col(new_assets_df)
    retired_assets_df = add_count_col(retired_assets_df)

    # Converting all "Count" values in the retired df to a negative value
    retired_assets_df["Count"] *= -1

    # Convert df to dictionary with key: "Model Number" and value: [["Entry 1 details"], ["Entry 2 details"], ...]
    new_assets_dict = create_dict(new_assets_df)
    retired_assets_dict = create_dict(retired_assets_df)

    print(new_assets_dict.get("692"))
    print(retired_assets_dict.get("692"))

    # Iterate through each entry in the new assets dictionary
    for key, val in new_assets_dict.items():
        # Check to see if the new model purchased was also retired
        if key in retired_assets_dict:
            for retired_asset in retired_assets_dict.get(key):
                for new_asset in new_assets_dict.get(key):
                    # Index 1 gives asset description, index 2 gives site, and index 3 gives shop code
                    if new_asset[1] == retired_asset[1] and \
                            new_asset[2] == retired_asset[2] and \
                            new_asset[3] == retired_asset[3]:
                        # Add because retired_assets_dict contains negative "Count" values
                        new_asset[5] += retired_asset[5]
                        # Remove retired asset from retired_assets_dict after count in new_assets_dict has been updated
                        retired_assets_dict.get(key).remove(retired_asset)
                # Delete key from retired_assets_dict if there are no more assets in its 2D list
                if len(retired_assets_dict.get(key)) == 0:
                    del retired_assets_dict[key]

    # If retired_assets_dict isn't empty, merge it with new_assets_dict
    if len(retired_assets_dict) > 0:
        for key, val in retired_assets_dict.items():
            if key not in new_assets_dict:
                new_assets_dict[key] = val
            else:
                for li in val:
                    new_assets_dict.get(key).append(li)

    print("======================Post Processing======================")
    print(new_assets_dict.get("692"))
    print(retired_assets_dict.get("692"))


if __name__ == "__main__":
    main()