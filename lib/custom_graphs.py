def plot_vert_time_serie(dataframe, x1_serie, y1_serie, y2_serie, gender_serie,
                         title, x1_title, y1_title, y2_title,
                         label_x1="Month",
                         label_y1_males="Total number of items sold to males",
                         label_y1_females="Total number of items sold to females",
                         label_y2="Total net earnings",
                         width=0.3,
                         color_y1_males=(0, 0, 0),
                         color_y1_females=(0, 0, 0, 0.65),
                         color_y2="#666699",
                         save=False,
                         file_name="Image 1.jpg",
                         show=True):
    """Create a vertical plot of two series same x axis.

    The x axis is meant to represent a time serie. The serie y1 is designed to
    be a stacked bar based on a third serie "gender_serie" providing the gender
    proportion.
    -----------------------------------------------------------------------
    Parameters:
    - dataframe = name of the dataframe (ex: "df")
    - x1_serie, y1_serie, y2_serie = name of the chosen columns (ex: "Type")
    - gender_serie = name of the serie containing the proportion of males as
    a decimal
    - file_name = works only if save is True, name of the image saved in the
    work directory
    """
    import numpy as np
    from matplotlib import pyplot as plt
    import matplotlib.ticker as ticker

    # Create image
    fig = plt.figure(figsize=(8 * 1.5, 4.5 * 1.5), dpi=80)

    # Create the subplots and rename the axis
    plot1 = fig.add_subplot()
    plot1.set_title(title, fontsize=20)
    plot1.set_xlabel(x1_title, fontsize=15)
    plot1.set_ylabel(y1_title, fontsize=15)

    plot1b = plot1.twinx()
    plot1b.set_ylabel(y2_title, fontsize=15)

    # Get the series ready to be plotted
    x1_serie = dataframe[x1_serie]
    y1_serie = dataframe[y1_serie]
    y2_serie = dataframe[y2_serie]
    gender_serie = dataframe[gender_serie]

    x_serie = np.arange(len(x1_serie))

    items_males = y1_serie * gender_serie
    items_females = y1_serie * (1 - gender_serie)

    # Plot the series
    plot1.bar(x_serie - width / 2 - 0.01, items_males,
              width=width,
              color=color_y1_males,
              label=label_y1_males)
    plot1.bar(x_serie - width / 2 - 0.01, items_males + items_females,
              width=width,
              color=color_y1_females,
              label=label_y1_females)

    plot1b.bar(x_serie + width / 2 + 0.01, y2_serie,
               width=width,
               color=color_y2,
               label=label_y2)

    # Set axis properties
    number_x_ticks = max(1, len(x1_serie) // 8)
    plot1.set_xticks(np.arange(0, len(x1_serie), number_x_ticks))
    plot1.set_xticklabels(x1_serie)

    plot1.yaxis.set_major_locator(ticker.MaxNLocator(10))
    plot1b.yaxis.set_major_locator(ticker.MaxNLocator(10))

    formatter = ticker.FormatStrFormatter('€%1.0f')
    plot1b.yaxis.set_major_formatter(formatter)

    # Add further details
    fig.legend(loc="lower center", ncol=3, fontsize=9)
    plot1.grid(True, axis="both", ls="--")

    # Save
    if save:
        fig.savefig(file_name)

    # Show
    if not show:
        plt.close()


def plot_cumulative_time_serie(dataframe, x1_serie, y1_serie, y2_serie,
                               title, x1_title, y1_title, y2_title,
                               label_x1="Month",
                               label_y1="Cumulative number of items sold",
                               label_y2="Cumulative net earnings",
                               width=0.3,
                               color_y1=(0, 0, 0),
                               color_y2="#666699",
                               save=False,
                               file_name="Image 1.jpg",
                               show=True):
    """Create a line plot of two series with the same x axis.

    -----------------------------------------------------------------------
    Parameters:
    - dataframe = name of the dataframe (ex: "df")
    - x1_serie, y1_serie, y2_serie = name of the chosen columns (ex: "Type")
    - file_name = works only if save is True, name of the image saved in the
    work directory
    """

    import numpy as np
    from matplotlib import pyplot as plt
    import matplotlib.ticker as ticker

    # Create image
    fig = plt.figure(figsize=(8 * 1.5, 4.5 * 1.5), dpi=80)

    # Create the subplots and rename the axis
    plot1 = fig.add_subplot()
    plot1.set_title(title, fontsize=20)
    plot1.set_xlabel(x1_title, fontsize=15)
    plot1.set_ylabel(y1_title, fontsize=15)

    plot1b = plot1.twinx()
    plot1b.set_ylabel(y2_title, fontsize=15)

    # Get the series ready to be plotted
    x1_serie = dataframe[x1_serie]
    y1_serie = dataframe[y1_serie]
    y2_serie = dataframe[y2_serie]

    x_serie = np.arange(len(x1_serie))

    # Plot the series
    plot1.plot(x_serie, y1_serie,
               "--", marker="o",
               color=color_y1,
               label=label_y1)

    plot1b.plot(x_serie, y2_serie,
                "--", marker="o",
                color=color_y2,
                label=label_y2)

    # Set axis properties
    number_x_ticks = max(1, len(x1_serie) // 8)
    plot1.set_xticks(np.arange(0, len(x1_serie), number_x_ticks))
    plot1.set_xticklabels(x1_serie)

    plot1.yaxis.set_major_locator(ticker.MaxNLocator(8))
    plot1b.yaxis.set_major_locator(ticker.MaxNLocator(8))

    formatter = ticker.FormatStrFormatter('€%1.0f')
    plot1b.yaxis.set_major_formatter(formatter)

    # Add further details
    fig.legend(loc="lower center", ncol=3, fontsize=9)
    plot1.grid(True, axis="both", ls="--")

    # Save
    if save:
        fig.savefig(file_name)

    # Show
    if not show:
        plt.close()


def plot_horizontal_bar(dataframe, y1_serie, x1_serie, x2_serie, gender_serie,
                        title, y1_title, x1_title, x2_title,
                        label_x1_males="Total number of items sold to males",
                        label_x1_females="Total number of items sold to females",
                        label_x2="Total net earnings",
                        custom_width=0.4,
                        color_x1_males=(0, 0, 0),
                        color_x1_females=(0, 0, 0, 0.65),
                        color_x2="#666699",
                        save=False,
                        file_name="Image 1.jpg",
                        show=True):
    """Create an horizontal plot of two series same y axis.

    The y axis is meant to be categorical. The serie x1 is designed to be a
    stacked bar based on a third serie "gender_serie" providing the gender
    proportion.
    -----------------------------------------------------------------------
    Parameters:
    - dataframe = name of the dataframe (ex: "df")
    - y1_serie, x1_serie, x2_serie = name of the chosen columns (ex: "Type")
    - file_name = works only if save is True, name of the image saved in the
    work directory
    """

    import numpy as np
    from matplotlib import pyplot as plt
    import matplotlib.ticker as ticker

    # Create image
    fig = plt.figure(figsize=(8 * 1.5, 4.5 * 1.5), dpi=80)

    # Create the subplots and rename the axis
    plot1 = fig.add_subplot()
    plot1.set_title(title, fontsize=20)
    plot1.set_ylabel(y1_title, fontsize=15)
    plot1.set_xlabel(x1_title, fontsize=15)

    plot1b = plot1.twiny()
    plot1b.set_xlabel(x2_title, fontsize=15)

    # Get the series ready to be plotted
    y1_serie = dataframe[y1_serie]
    x1_serie = dataframe[x1_serie]
    x2_serie = dataframe[x2_serie]
    gender_serie = dataframe[gender_serie]

    width = custom_width
    y_serie = np.arange(len(y1_serie))
    items_males = x1_serie * gender_serie
    items_females = x1_serie * (1 - gender_serie)

    # Plot the series
    plot1.barh(y_serie - width / 2 - 0.01, items_males,
               height=width,
               color=color_x1_males,
               label=label_x1_males)
    plot1.barh(y_serie - width / 2 - 0.01, items_males + items_females,
               height=width,
               color=color_x1_females,
               label=label_x1_females)
    plot1b.barh(y_serie + width / 2 + 0.01, x2_serie,
                height=width,
                color=color_x2,
                label=label_x2)

    # Set axis properties
    plot1.set_yticks(np.arange(0, len(y1_serie), 1))
    plot1.set_yticklabels(y1_serie)
    plot1.set_xticks(np.arange(0, x1_serie.max() + 1, 1))

    formatter = ticker.FormatStrFormatter('€%1.0f')
    plot1b.xaxis.set_major_formatter(formatter)

    # Add further details
    fig.legend(loc="lower center", ncol=3, fontsize=9)

    # Save
    if save:
        fig.savefig(file_name)
    # Show
    if not show:
        plt.close()
