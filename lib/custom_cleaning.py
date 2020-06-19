#!/usr/bin/env python

def estimate_gender(df, first_name_col, file_names_list="Names_list.csv"):
    '''
    Estimates customer's gender based on first name using a names' database.

    -----------------------------------------------------------------------
    Parameters:
    - df = name of the dataframe
    - first_name_col = name of the column containing the first name (ex: "First
    name")
    - file_names_list = name of the csv file containing the names list
    '''
    import pandas as pd
    import numpy as np

    dr_names = pd.read_csv(file_names_list)
    df["First first name"] = df[first_name_col].str.split(pat=" ").str[0]

    cond_male = df["First first name"].isin(list(dr_names["Male names"]))
    cond_female = df["First first name"].isin(list(dr_names["Female names"]))

    male_col = np.where(cond_male, True,
                        np.where(cond_female, False, np.NaN)).astype("bool")
    return male_col


def gen_month_code(df, date_col):
    '''
    Creates a month code in the form NNYYYY based on a date column.

    -----------------------------------------------------------------------
    Parameters:
    - df = name of the dataframe
    - date_col = name of the column containing the date (ex: "date")
    '''
    import pandas as pd
    import numpy as np

    month = pd.DatetimeIndex(df[date_col]).month
    month_code_col = (pd.DatetimeIndex(df[date_col]).year.astype(str)
                      + "M"
                      + np.where(month.astype(str).str.len() == 1,
                                 "0" + month.astype(str), month.astype(str))
                      )
    return month_code_col


def unpack_multiple_orders(df, items_col="Items",
                           items_count_col="Item count",
                           total_price_col="Item total",
                           total_paid_col="Total price",
                           code_col="Number",
                           discount_col="Total discount",
                           shipping_col="Total shipping"):
    '''
    Separate orders with different items into multiple rows.

    -----------------------------------------------------------------------
    Parameters:
    - df = name of the dataframe
    - col parameters = name of the columns (ex: "column 1")
    '''
    import numpy as np

    df["Different item"] = df[items_col].str.count("product_name")

    while df["Different item"].max() >= 2:
        filt = (df["Different item"] >= 2)
        df_temp = df.loc[filt].copy()
        df_temp[items_col] = df_temp[items_col].str.split(";", n=1).str[1]
        df.loc[filt, items_col] = df.loc[filt, items_col].str.split(";", n=1).str[0].copy()

        df_temp[items_count_col] = df_temp[items_col].str.extract(pat="quantity:(\d+)").astype(int).copy()
        df_temp[total_price_col] = df_temp[items_col].str.extract(pat="total:(\d+.+)").astype(float).copy()

        df[items_count_col] = np.where(filt, df[items_col].str.extract(pat="quantity:(\d+)").astype(int),
                                       df[items_count_col])
        df[total_price_col] = np.where(filt, df[items_col].str.extract(pat="total:(\d+.+)").astype(float),
                                       df[total_price_col])

        df_temp[total_paid_col] = df_temp[total_price_col]
        df_temp[discount_col] = 0
        df_temp[shipping_col] = 0
        code_letters = df_temp[code_col].str.split("-").str[0]
        code_new_number = (df_temp[code_col].str.split("-").str[1].astype(int) + 1).astype(str)
        df_temp[code_col] = (code_letters + "-" + code_new_number)
        df = df.append(df_temp, ignore_index=True)
        df["Different item"] = df[items_col].str.count("product_name")

    df.drop("Different item", axis=1, inplace=True)

    return df


def match_zip_to_city(df, zip_column="Shipping zip",
                      zip_column_newname="Zip", file_zip="Lista_comuni.xlsx",
                      duplicate_column="Number"):
    '''
    Derive city, province and region from the zip code using an external source.

    The issue with the zip code is that in the source file sone are left
    generic. Like Milan is associated with all the zip codes of the form 201??.
    The program is structured to fill up the information in two rounds, first
    working on the extact matches and then on the partial ones. (It could
    probably be optimized).
    -----------------------------------------------------------------------
    Parameters:
    - df = name of the dataframe
    - file_zip = name of the file containing the information on zip codes.
    - col parameters = name of the columns (ex: "column 1")
    '''
    import pandas as pd
    import numpy as np

    # Drop columns containing info which will be infered from the zip code
    to_drop = ["Shipping city", "Shipping state"]
    df.drop(to_drop, axis=1, inplace=True)
    # Load the file containing the info on the zip codes
    dr_zips = pd.read_excel(file_zip)
    # Align types and column names
    df.rename(columns={zip_column : zip_column_newname}, inplace=True)
    df[zip_column_newname] = df[zip_column_newname].astype("str")
    dr_zips.rename(columns={"CAP" : zip_column_newname}, inplace=True)

    # Fill the first round of zip codes, the complete ones (ex: NNNNNN)
    df = pd.merge(df, dr_zips, on=zip_column_newname, how="left", )

    # Create the new zip codes which are used for the generic ones (ex: NNNN??)
    new_zip = df[zip_column_newname].str[0:3]
    # New fill just for the ones which were not matched
    df["Zip 1"] = np.where(df["Comune"].isnull(), new_zip, np.nan)
    dr_zips["Zip 1"] = dr_zips[zip_column_newname].str[0:3]
    dr_zips.drop(zip_column_newname, axis=1, inplace=True)

    # Fill the second round of zip codes
    df = pd.merge(df, dr_zips, on="Zip 1", how="left")

    # Remove all the redundant information and correct names
    df = df.drop_duplicates(subset=[duplicate_column, zip_column_newname])

    to_keep = ["Comune_x", "Provincia_x", "Regione_x"]
    to_discard = ["Comune_y", "Provincia_y", "Regione_y"]
    for final, temporary in zip(to_keep, to_discard):
        df[final] = df[final].fillna(df[temporary])

    to_discard.append("Zip 1")
    df.drop(to_discard, axis=1, inplace=True)

    new_names = {"Comune_x" : "City", "Provincia_x" : "Province",
                 "Regione_x" : "Region"}
    df.rename(columns=new_names, inplace=True)

    return df


def aggregate_status(df, list_of_columns,
                     accepted_list=["completed", "shipped"], drop=False):
    '''
    Aggregate multiple status column into a single one with AND operator.

    -----------------------------------------------------------------------
    Parameters:
    - df = name of the dataframe
    - list_of_columns = columns in form of list which should be checked
    '''
    new_status = df[list_of_columns[0]].isin(accepted_list)

    for col in list_of_columns[1:]:
        new_status = df[col].isin(accepted_list) & new_status

    if drop:
        df.drop(columns=list_of_columns, inplace=True)

    return new_status
