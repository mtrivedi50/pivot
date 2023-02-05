# Dynamic Excel Pivot Tables for "Big Data"

## The problem
Pivot tables are one of the most powerful tools one can use to explore and analyze data. However it's difficult to pivot "big data" (>1M rows) without learning new tooling (e.g., Looker) or writing custom SQL. Business users want to be able to pivot their data in the tools with which they are comfortable (e.g., Excel). In addition, they want an interface that is intuitive, interactive, and dynamic — one that has the sophistication of modern software, but the power and simplicity of Excel.


## The solution
A dynamic pivot table application built using Excel, VBA, Power Query, and Power Pivot.


## How it works
Users provide the following via a secure platform:
- The data as a CSV
- Which columns to use as filters
- Specific data aggregations

We take these inputs, transform your raw data to a form that is compatible with our tooling, and then generate your Excel dashboard. This dashboard is of file type `.xlsm` and will be available via download from the platform.


## How to use the product
There are three sections to the dashboard: the filter section, the button section, and the actual pivot table. The videos below show how to use each. This example workbook uses airline delay data from [here](https://www.kaggle.com/datasets/heemalichaudhari/airlines-delay).

<details>
<summary>The buttons on the left-hand-side in blue (called "slicers" in Excel terminology) act as filters. They can be used to filter the data to a specific combination of inputs.</summary>
<video src="./assets/pivot_filters.mov" controls="controls" style="max-width:1000px"></video>
</details>

<details>
<summary>The buttons on the top in blue control columns. If multiple values are selected for a specific filter, you have the option of placing that filter out wide into columns.</summary>
<video src="./assets/pivot_filter_cols.mov" controls="controls" style="max-width:1000px"></video>
</details>

<details>
<summary>The buttons on the top in green control data aggregations. These can be toggled on and off.</summary>
<video src="./assets/pivot_values.mov" controls="controls" style="max-width:1000px"></video>
</details>

<details>
<summary>The buttons in the middle in yellow control the rows. Only one of these buttons can be specified at a time (though, we're working on allowing users to specify multiple).</summary>
<video src="./assets/pivot_rows.mov" controls="controls" style="max-width:1000px"></video>
</details>


## Sample workbook
Download a sample workbook here.


## Limitations
This tool is extraordinarily powerful, and it allow business users to explore, analyze, and summarize data without dedicated tooling or data analyst/scientist resources. However, it has the following limitations:
- We find that this dashboard works best on data that is <10M rows.
- The more filters requested, the slower the dashboard will be.
- The dashboard itself is of filetype `.xlsm`. Make sure that you enable macros when opening the workbook, and confirm that this does not cause security concerns for users prior to distributing.
- The underlying code will be locked to protect IP.


## Waitlist
Join the waitlist here!