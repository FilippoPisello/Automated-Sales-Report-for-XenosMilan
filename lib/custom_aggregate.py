def aggregate_by_date(original_dataframe,
                      time_serie_to_groupby,
                      serie_name="Period"):
    """Perform a groupby on a time serie with premade operations

    Parameters:
    - orginal_dataframe = name of the dataframe on which the aggregation is
    performed (ex: "df")
    - time_serie_to_groupby = name of the serie/column which is used as key
    for the aggregation, is intended to be a time serie (ex: "Month code")
    - serie_name = name to be assigned to the column containing the key of the
    aggregation after the groupby is performed (ex: "column 1")
    """

    aggregate = {"Items count" : ["sum", "mean"], "Male" : "mean",
                 "Raw price" : ["sum", "mean"], "Net earnings" : "sum"}

    df1 = original_dataframe.groupby([time_serie_to_groupby],
                                     as_index=False).agg(aggregate)

    df1.columns = df1.columns.droplevel(0)

    df1.columns = [serie_name, "Tot items ordered", "Avg items per order",
                   "Males percentage", "Tot raw price", "Avg raw price",
                   "Tot net earnings"]

    df1["Share of net earnings"] = ((df1["Tot net earnings"]
                                     / df1["Tot net earnings"].sum())).round(2)

    df1["Cumulative net earnings"] = df1["Tot net earnings"].cumsum().round(1)

    df1["Cumulative items sold"] = df1["Tot items ordered"].cumsum()

    df1 = df1.round({"Males percentage" : 2,
                     "Tot raw price" : 1,
                     "Avg raw price" : 1,
                     "Tot net earnings" : 1})
    return df1


def aggregate_by_category(original_dataframe,
                          serie_to_groupby,
                          serie_name="Values"):
    """Perform a groupby with premade operations

    Parameters:
    - orginal_dataframe = name of the dataframe on which the aggregation is
    performed (ex: "df")
    - serie_to_groupby = name of the serie/column which is used as key for the
    aggregation (ex: "column 1")
    - serie_name = name to be assigned to the column containing the key of the
    aggregation after the groupby is performed (ex: "column 1")
    """

    aggregate = {"Items count" : "sum", "Male" : "mean",
                 "Raw price" : ["sum", "mean"], "Net earnings" : "sum"}

    df1 = original_dataframe.groupby([serie_to_groupby],
                                     as_index=False).agg(aggregate)

    df1.columns = df1.columns.droplevel(0)
    df1.columns = [serie_name, "Tot items count", "Males percentage",
                   "Tot raw price", "Avg raw price", "Tot net earnings"]

    df1["Share net earnings"] = ((df1["Tot net earnings"]
                                  / df1["Tot net earnings"].sum()) ).round(2)

    df1 = df1.round({"Males percentage" : 2,
                     "Tot raw price" : 1,
                     "Avg raw price" : 1,
                     "Tot net earnings" : 1})
    return df1
