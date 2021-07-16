# Michigan Department of State Limited English Proficiency Data Manipulation

## Summary
This was a project I implemented for the Michigan Department of State in a 48 hour timeframe.

The goal of the project was to consolidate information contained in two different excel documents, merging relevant information together.
One spreadsheet contained the addresses of all Secretary of State branches in Michigan. The other spreadsheet contained the primary and secondary
foreign languages spoken in each county in Michigan. I used the openpyxl package in combination with the mapquest API to associate SOS branch zipcodes
with their respective counties. I then linked the primary and secondary foreign languages spoken in each Michigan county to their SOS branch(es) in one
unified excel doc.

The Secretary of State was able to use this document to provide all branches across the state with documents written in those counties' primary foreign languages
so that visitors with limited english proficiency could more easily find help.

### Things I learned from this project:
1. You likely won't know everything you need to in order to complete a new project.
2. Knowing what you don't know and learning on the fly is key to success.
3. Code can always be refactored and improved—this project is no exception.
4. Python is super useful and fun.

### Why Python?
1. I had just completed a graduate-level intermediate course in Python and I wanted to apply my skills to a real world client project.
2. This data cleaning and manipulating *could* have been done by hand, but what if the dataset was ten thousand rows instead of around one hundred? What if other states needed help with a similar issue? What if the SOS wanted to scale this project? The code is more adaptable.
3. Python is extremely popular—another developer could take on this project once my work with MDOS was over.

## Packages required to run this program
```python
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import requests
import json
```

Feel free to reach out to me if you see anything that could use improvement!
