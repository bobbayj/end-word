# end-word
Can we automate report generation from copy-pasting .docx and .xls files with some Python? **Yes we can!**

Using a bit of perserverance and digging, we can get everything we need using Microsoft Office's OpenXML structuring and create templated reports using FOSS Python libraries.

**Progress:** Still prototyping. Play around with the jupyter notebook if you would like.

## Third-party Dependencies
- `python-docx`; Python wrapper for Word's OpenXML
    - Currently using `bayoo-docx`. See this [StackOverflow](https://stackoverflow.com/questions/30292039/pip-install-forked-github-repo) for details on how to install it from the [python-docx fork](https://github.com/BayooG/bayoo-docx)
- `docxtpl`; Enables the use of pre-set Word templates. Built over `python-docx`
- `openpyxl`; Python wrapper for Excel's OpenXML

## Goal
1. Have a title page template that we can populate and then sequentially fill/append with data
2. Copy over Excel table data with cell formatting
    - No number formatting yet
3. Copy over Word data. Preserve all formatting and images (+ image formats)
4. Have boilerplate code so the project is DRY