<!-- README.md is generated from README.Rmd. Please edit that file -->
officeR
-------

[![Build Status](https://travis-ci.org/itsdalmo/officeR.svg?branch=master)](https://travis-ci.org/itsdalmo/officeR) [![codecov.io](http://codecov.io/github/itsdalmo/officeR/coverage.svg?branch=master)](http://codecov.io/github/itsdalmo/officeR?branch=master)

officeR is meant to provide a more seamless experience when working with MS Office and R. The toolbox mostly consists of wrapper-functions that provide a consistent syntax for a few select packages.

Note: This is a work in progress.

Installation
------------

#### Dependencies ahead of CRAN

-   [Haven](https://github.com/hadley/haven) 0.2.0.9000: Fixes crashes when writing strings longer than 256 characters.

``` r
if (!require(devtools)) {
    install.packages("devtools")
}
devtools::install_github("hadley/haven")
```

#### Install officeR

Development version:

``` r
devtools::install_github("itsdalmo/officeR")
```

CRAN:

``` r
# Not on CRAN yet.
```

#### Optional: Powerpoint support

-   The latest JRE from [Java](http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html).
-   [ReporteRs](https://github.com/davidgohel/ReporteRs) from CRAN:

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
