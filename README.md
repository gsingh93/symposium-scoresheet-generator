# Sikh Symposium Scoresheet Generator

Generates an Excel scoresheet for use in the [Sikh Youth Symposium](http://www.sikhyouthalliance.org/youth-symposium/). Requires Python 2, the [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) library, and the [xlrd](https://pypi.python.org/pypi/xlrd) library.

## Usage

Generate a scoresheet by running `python gen.py`. This command looks for a spreadsheet named `scoresheet_info_<year>.xlsx`, where <year> is the current year. For example, if the current year is 2015, it will look for a spreadsheet named `scoresheet_info_2015.xlsx`. An example spreadsheet is provided. Each worksheet in this workbook contains the info for one group. One spreadsheet will be generated for each group.
