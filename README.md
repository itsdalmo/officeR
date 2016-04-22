<!-- README.md is generated from README.Rmd. Please edit that file -->
officeR
-------

[![Build Status](https://travis-ci.org/itsdalmo/officeR.svg?branch=master)](https://travis-ci.org/itsdalmo/officeR) [![codecov.io](http://codecov.io/github/itsdalmo/officeR/coverage.svg?branch=master)](http://codecov.io/github/itsdalmo/officeR?branch=master)

Tiny toolbox (work in progress) for working with office and R.

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
