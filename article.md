---
author: "Kyle Jones"
date_published: "August 23, 2025"
date_exported_from_medium: "November 10, 2025"
canonical_link: "https://medium.com/@kyle-t-jones/stop-wrestling-with-excel-use-python-to-automate-it-6d169e5346d6"
---

# Stop Wrestling with Excel: Use Python to Automate It Most analysts live in Excel. It's familiar, visual, and flexible. But
anyone who has pushed Excel past a few hundred thousand rows, chained...

### *Stop Wrestling with Excel: Use Python to Automate It* 

Most analysts live in Excel. It's familiar, visual, and flexible. But anyone who has pushed Excel past a few hundred thousand rows, chained together too many `VLOOKUP`s, or debugged broken VBA knows its limits.

This is where Python shines. With libraries like Pandas and Openpyxl, you can treat Excel files as inputs and outputs in a larger, repeatable workflow. Instead of clicking and dragging formulas by hand, you can automate them. Instead of relying on VBA macros that break easily, you can write clean Python code that works across datasets.

In this article, we'll walk through how Python can read, transform, and write Excel files. You'll see why many workflows are faster, more reliable, and more transparent when moved from Excel to Python.

### Excel vs. Python: When to Use Each
Excel is great for quick exploration. You can scan a sheet, filter, sort, and build a quick pivot table in minutes. For small data tasks, it's still the fastest way to "see" what's happening.

Python begins to outperform Excel when the work gets heavier. If you need to:

- Process more than one million rows
- Apply nested lookups, matches, or array functions
- Repeat the same transformations across multiple datasets
- Automate cleaning tasks without manual clicks
- Replace fragile VBA scripts with maintainable code

Then Python wins. Pandas can handle large datasets in memory, clean them with chainable transformations, and save them back into Excel for colleagues who still prefer spreadsheets.

### Tutorial: Using Python with Excel
Let's walk through an example that shows both Pandas and Openpyxl in action.

First, install the required libraries:

```python
! pip install -r requirements.txt
import pandas as pd
import openpyxl
```

### Paths to input and output files
``` 
path_to_data = 'data/excel_input.xlsx'
path_to_output = 'data/excel_output_new_file.xlsx'
```

### Read data from Excel
``` 
retail = pd.read_excel(path_to_data, sheet_name='retail')
mfn = pd.read_excel(path_to_data, sheet_name='mfn')

print(retail.head())
print(mfn.head())
```

### Combine the data
``` 
combined = pd.concat([retail, mfn])
combined.to_excel(path_to_output, sheet_name='combined', index=False)
```

Now you've created a new Excel file with a combined worksheet.

### Writing to the Same File
Sometimes you don't want a new file --- you want to add a new sheet to an existing workbook. Openpyxl makes that possible:

``` 
workbook = openpyxl.load_workbook(path_to_data)
writer = pd.ExcelWriter(path_to_data, engine='openpyxl')
writer.book = workbook

combined.to_excel(writer, sheet_name='combined', index=False)
writer.save()
```

⚠️ Always back up your original file before writing to it.

### A Worked Example: Cleaning Excel Data
Here's a more advanced case where we clean a messy dataset directly from Excel.

```python
import pandas as pd
import openpyxl 

filename = 'data/XL-02-PC-05-Cleaning-Up-Data-Before.xlsx'
wb = openpyxl.load_workbook(filename)
ws = wb['Data']
df = pd.DataFrame(ws.values)
# Make first row the header
df.columns = df.iloc[0]
df = df.drop(df.index[0])
# Split full name
df[['First Name', 'Last Name']] = df['Customer Name'].str.split(" ", n=1, expand=True)
# Split address field
df[['Address', 'City', 'State', 'ZIP']] = df['Address, City, State, and ZIP'].str.split(",", n=3, expand=True)
# Fix case
df['Address'] = df['Address'].str.title()
df['City'] = df['City'].str.title()
df['State'] = df['State'].str.upper()
from openpyxl.utils.dataframe import dataframe_to_rows
# Save cleaned data back into Excel
wb.create_sheet(index=0, title='Cleaned Data') 
ws = wb['Cleaned Data']
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)
wb.save(filename)
wb.close()
```

This turns a raw file into a clean, structured dataset, all with reproducible code.
### Extensions 

If you want to take this further, try exploring:

- When Openpyxl is better than Excel (e.g., automating repetitive reporting tasks)
- How to replace fragile VBA scripts with Python functions
- How to schedule Python scripts to run nightly so that your Excel workbooks are always up to date

### Other useful resrouces
- [[*Python for Data Analysis*](http://shop.oreilly.com/product/0636920023784.do), by Wes McKinney]
- [[*Automate the Boring Stuff with Python*](https://automatetheboringstuff.com/), by Al Sweigart (see Chapter 13 on Excel)]
- [[Openpyxl Documentation](http://openpyxl.readthedocs.org/en/default)]
- [[Working with Excel sheets in Python using openpyxl](https://medium.com/aubergine-solutions/working-with-excel-sheets-in-python-using-openpyxl-4f9fd32de87f)]

### Summary
Excel remains a great tool for quick inspection and small analyses. But Python takes over when you need scale, speed, and reliability. With Pandas and Openpyxl, you don't have to abandon Excel --- you can extend it.

The best workflow often blends the two: clean and process data in Python, then hand the results back in Excel format for easy sharing.
