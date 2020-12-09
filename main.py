import os
import pandas as pd

pd.set_option("display.expand_frame_repr", False)


def read_new_assets():

    new_assets_file = os.listdir("new_assets")[0]
    new_assets_path = "new_assets/{file}".format(file=new_assets_file)
    new_assets_df = pd.read_excel(new_assets_path,
                                  usecols=["Model Number",
                                           "Asset Description",
                                           "Segment Description",
                                           "Site Code",
                                           "Shop"])
    return new_assets_df


def read_retired_assets():

    retired_assets_file = os.listdir("retired_assets")[0]
    retired_assets_path = "retired_assets/{file}".format(file=retired_assets_file)
    retired_assets_df = pd.read_excel(retired_assets_path,
                                      usecols=["Model Number",
                                               "Asset Description",
                                               "Segment Description",
                                               "Site Code",
                                               "Shop"])

    return retired_assets_df


def add_count_col(df):
    return df.\
        groupby(["Model Number", "Asset Description", "Site Code", "Shop", "Segment Description"]).\
        size().\
        to_frame("Count").\
        reset_index()


def create_dict(df):
    assets_dict = {}
    for index, row in df.iterrows():
        if row["Model Number"] in assets_dict:
            assets_dict.get(row["Model Number"]).append(row.tolist())
        else:
            assets_dict[row["Model Number"]] = []
            assets_dict.get(row["Model Number"]).append(row.tolist())
    return assets_dict


def main():

    new_assets_df = read_new_assets()
    retired_assets_df = read_retired_assets()

    new_assets_df = add_count_col(new_assets_df)
    retired_assets_df = add_count_col(retired_assets_df)

    new_assets_dict = create_dict(new_assets_df)
    retired_assets_dict = create_dict(retired_assets_df)

    print(new_assets_dict)
    print(retired_assets_dict)

    return new_assets_dict


if __name__ == "__main__":
    main()