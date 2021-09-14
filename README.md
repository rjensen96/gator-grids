# Gator Grids [Coming Soon!]

## Meet your friendly aggre-gator!

Gator Grids helps you sum up your data exactly the way you want - without writing a single formula.

<img src="./GG_sample.gif" />
![Gator Grids Demo](./GG_sample.gif)

Connect to a table in your Excel workbook and assign filters to rows and columns. In cells where those filters intersect, you get data!

Install Gator Grids for free from Microsoft AppSource. Click <strong>Insert > Office Add-ins</strong> in Excel to get started.

## Getting Started

<h3>Install the add-in</h3>

Gator Grids is <b>coming soon</b> to Microsoft AppSource! Search for it in Excel by clicking <b>Insert > Add-ins</b> and searching for Gator Grids.

<h3>Make a table</h3>

To create a Grid, your data must be in a <a href=https://support.microsoft.com/en-us/office/overview-of-excel-tables-7ab0bb7d-3a9e-4b56-a3c9-6c94334e492c target="blank">Table<NewTabIcon /></a>. One quick way to create a table is to select your data and click <strong>Home > Format as Table</strong>.

<h3>Select columns and rows</h3>
<p>
    On a different worksheet, Select a full column where you want data. You can use any worksheet, but the first
    cell must be blank. For example, if you select Column C, cell C1 must be clear.
</p>

<h3>Pick data source & filters</h3>

<p>
    In this taskpane, click <strong>Menu > Edit</strong>. Select a table and choose which column to sum. Click a
    field name to expand a list of all unique values in that field, and select one or more to use as filters.
</p>
<p>
    Click <strong>Save</strong>. A random, unique five-digit code (a <i>GatorKey</i>) will appear in the first
    cell. This code is a key that Gator Grids uses to save your filters and settings for each row and column.
    Repeat this process for a few rows and columns.
</p>

<h3>Refresh</h3>

<p>
    In the Gator Grids ribbon, click <strong>Refresh Grids</strong>. In cells where the columns and rows with
    GatorKeys intersect, Gator Grids will merge your row and column filter settings and aggregate your data,
    returning the result in each cell.
</p>

<h3>Bonus: View mode</h3>
<p>
    After setting up a few GatorKeys, click the glasses icon to view settings for the selected rows or columns. This view updates as you select rows and columns around your Grid.
</p>
<p>
    Selecting multiple rows or column will display only the filters and settings that the selected GatorKeys have
    in common. View mode also turns green to indicate that you have selected multiple GatorKeys.
</p>

## About

Gator Grids helps people create and maintain Excel reports faster and better.

### How it works

Select full rows or columns, select some options in the taskpane, and click `Save`. The random and unique 5-digit key in the first row or column tells Gator Grids which settings apply to which ranges. Rows and columns with these keys make up a "Grid."

Click `Refresh Grids` to calculate cells in the Grid. At each intersection of a row and column with keys, the filter settings for the row and column are merged to calculate the final cell value.

Filter settings are saved as a custom XML part in the workbook. Gator Grids does not use an external server to store data because everything is embedded directly into the workbook.

When compared to formulas and Pivot tables, Gator Grids is a clear winner. Grids are easier to maintain and create. Creating a Grid requires no extra technical knowledge. The only requirement is that a user knows what they want to see.

### Powered by Office.js

Gator Grids a modern Office Add-in built with Office.js and the React framework. You can use it on Windows and Mac, both online and in the desktop.

### Developer

My name is Ryan Jensen. I have Master's degree in accounting and am working on a Master's degree in Software Development. While working as an accountant and financial analyst, I saw a great need for a tool like this in my own work. I hope that this tool can help you in your Excel work as well.

## Contributing

For now, I have decided not to open-source this project, although the add-in is available for free. If you want to contribute, please reach out to me privately and express your interest.

## Issues

Please use this repo to log issues or bugs. I am currently a team of one and a full-time student, so please understand that it may take some time for me to push updates.
