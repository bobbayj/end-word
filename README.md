# end-word

Can we automate report generation from copy-pasting .docx and .xls files with some Python? **Yes we can!**

Using a bit of perserverance and digging, we can get everything we need using Microsoft Office's OpenXML structuring and create templated reports using FOSS Python libraries.

**Progress:** Still prototyping. Play around with the jupyter notebook if you would like.

## Third-party Dependencies

- `python-docx`; Python wrapper for Word's OpenXML
  - Currently using `bayoo-docx`. See this [StackOverflow](https://stackoverflow.com/questions/30292039/pip-install-forked-github-repo) for details on how to install it from the [python-docx fork](https://github.com/BayooG/bayoo-docx)
- `docxtpl`; Enables the use of pre-set Word templates. Built over `python-docx`
- `openpyxl`; Python wrapper for Excel's OpenXML. *Change this to xlwings for simpler interface with Excel*

## Goal

1. Have a title page template that we can populate and then sequentially fill/append with data
2. Copy over Excel table data with cell formatting
3. Copy over Word data. Preserve all formatting and images (+ image formats)
4. Have boilerplate code so the project is DRY

### Regarding footnotes

- Template must contain all the styles required (including footnote reference and text styles!)
- Currently using Bayoo-docx in local dir to support footnotes read/write.
  - Should work fine installing as mentioned above. If not, I uncommented a function related to footnotes.

---

## To-do List

- In style_tbl > font, include font size change
- Re-factor code into a proper python package
  - Draft a project structure
  - Use Jupyter notebook for control
- Convert openpyxl code to xlwings code
- Chart_xlsx() - Excel chart creation | [Documentation](https://openpyxl.readthedocs.io/en/stable/charts/introduction.html)
  - [Reader drawings module](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.reader.drawings.html)
- Check_styles() - Formatting parser when publishing docx
  - Ensure:
    - There is not too much blank space on a page
    - Tables and titles are not left orphaned on a page
    - Consistent formatting throughout

## Longer-term issues

- Front-end | Parameter entry:
  - Manual entry
  - Auto-load in
- Front-end <--link--> Back-end
  - Where will files be stored? Privacy issues?
- Flexibly defined style formats?
- Creation of some standard templates
  Equity f/s
  - Fixed income f/s
    - Fund f/s
- Scheduler (.bat?)
- Delivery method
  - Email, file saved
    - Cloud-based vs Local?
    - Security concerns?

## Testing Feedback

- Needs to look flawless...the perfect output
- Needs a very simple user interface
  - Jupyter notebook; simple to create, but must be careful about who our target users are
  - Traditional Windows DOS program; time-consuming to create, anyone can use
