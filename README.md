<!-- README.md is generated from README.Rmd. Please edit that file -->
seamless
--------

[![Linux/OSX](https://travis-ci.org/itsdalmo/seamless.svg?branch=master)](https://travis-ci.org/itsdalmo/seamless) [![Windows](https://ci.appveyor.com/api/projects/status/github/itsdalmo/seamless?branch=master&svg=true)](https://ci.appveyor.com/project/itsdalmo/seamless) [![Coverage](http://codecov.io/github/itsdalmo/seamless/coverage.svg?branch=master)](http://codecov.io/github/itsdalmo/seamless?branch=master)

This package is meant to provide a more "seamless" experience when using R in a workflow with MS Office and other programs in general. In particular, it is a provides a consistent syntax for reading and writing data to several formats, including the windows/osx clipboards, and sharepoint (read-only) over HTTP.

Note: This is a work in progress.

Installation
------------

#### Install seamless

Development version:

``` r
devtools::install_github("itsdalmo/seamless")
```

CRAN:

``` r
# Not on CRAN yet.
```

Optional: Powerpoint support
----------------------------

-   First install the latest JRE from [Java](http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html).
-   Next, install [ReporteRs](https://github.com/davidgohel/ReporteRs) from CRAN:

``` r
install.packages("ReporteRs")
```

Usage
-----

#### Read/write data

-   `read_data()`: Reads Excel (.xls, .xlsx), SPSS (.sav) and R (.Rdata, .Rda) files based on their extension.
-   `write_data()`: Write data to a specific format based on extension. Including workbooks (see below).
-   `from_clipboard()`: Read contents (both strings and tabular data) from OS X/Windows clipboards.
-   `to_clipboard()`: Write strings or data to the clipboard.

#### Excel

-   `excel_workbook()`: Create a Excel workbook in R.
-   `to_excel()`: Send supported data formats from R to the workbook object.
-   Use `write_data()` to save the Excel workbook to a `.xlsx` file.

#### Powerpoint

-   `ppt_workbook()`: Create a Powerpoint workbook (R6 wrapper for ReporteRs' doc object).
-   `to_ppt()`: Send supported R objects (data.frame, plots, markdown etc.) to the workbook object.
-   Use `write_data()` to save the Powerpoint workbook to a `.pptx` file.

#### Sharepoint

-   `sharepoint_link()`: Let's you create a sharepoint link and read files (using `httr`) on sharepoint with `read_data()`.
-   `sharepoint_mount()`: Converts a link to a Windows path for the same destination, if sharepoint is mounted.

#### Piping with `%>%`

Both Excel and Powerpoint workbooks are mutable objects, which means that `to_excel` and `to_ppt` can be used without assigning results. And since both functions also take data as the first argument, they work well with dplyr/tidyr and `%>%` in general.
