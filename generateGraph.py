import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import json
import matplotlib.patches as mpatches
import matplotlib
import textwrap
from pathlib import Path
import os
import gc

from CsvParser import ExcelParser
file_extension = '.xlsx'
directory = Path('')
for file_path in directory.glob(f'*{file_extension}'):
    with open("config.json", "r") as file:
        config = json.load(file)

    parser = ExcelParser(file_path)

    # Load the Excel file
    parser.load_excel()

    # Load a specific sheet
    data = parser.load_sheet(0)

    matplotlib.rc('font', family=config["font"])
    colors = [config["color1"], config["color2"], config["color3"], config["color4"]]

    fig, axes = plt.subplots(ncols = 1, nrows = data.index.size,constrained_layout = True, figsize = (config["figsize_x"],config["figsize_multy"]*len(data.index)),sharex=True)


    for i, ax in enumerate(axes):
        sns.barplot(data.loc[data.index[i]], orient='h', edgecolor="white", width=1, ax = axes[i])
        for j, bar in enumerate(ax.patches):
            bar.set_facecolor(colors[j])
        for spine in ax.spines.values():
            spine.set_visible(False)

        for p in ax.patches:
            ax.annotate(f'{int (p.get_width()*100):}%',  # Display value
                        (p.get_width(), p.get_y() + p.get_height() / 2),  # Position to the right of the bar
                        ha='left', va='center',  # Aligning text
                        fontsize=config["fontsize bar percents"], color='black',  # Font size and color
                        xytext=(5, 0),  # Offset the text to the right
                        textcoords='offset points')
        ax.tick_params(axis='x', which='both', length=0)
        ax.set_yticks([])
        ax.grid(True,linestyle='-', linewidth=1, color='grey', alpha=0.1)
        ax.set_xlabel('')
        ax.set_title("")
        ax.set_axisbelow(True)
        ax.set_ylabel(textwrap.fill(data.index[i],config["charlimit"]), rotation = 0,labelpad = 5, ha = 'right',fontsize = config["fontsize items"], va = "center_baseline", weight = "bold")


    plt.tick_params(labelcolor="black", bottom=False, left=False)
    plt.xticks([0,0.25,0.50,0.75,1], fontsize = 12)
    plt.gca().xaxis.set_major_formatter(mticker.PercentFormatter(xmax=1, symbol=" %"))  # xmax=1 means values are in the range [0, 1]
    plt.grid(True, linestyle='-', linewidth=1, color='grey', alpha=0.1)
    plt.yticks([])
    custom_patch_list = []
    for i,column in enumerate(data.columns):
        custom_patch = mpatches.Patch(color=colors[i], label=column)
        custom_patch_list.append(custom_patch)
    plt.legend(handles=custom_patch_list, loc = "lower center",ncol = 4,  bbox_to_anchor=(-0.01, -1), fontsize = 14)
    plt.savefig("abbildungen/" + str(file_path)+".jpg",dpi = config["dpi"])
    plt.close()  # Close the current figure
    plt.close('all')
    file.close()



