#' Read from Windows/OSX clipboard
#'
#' Thin wrapper for reading from windows/OSX clipboards with the most-used defaults.
#' The function first reads in the lines, checks if the delimiter is present in the lines
#' and then converts it to a data.frame.
#'
#' @param sep The delimiter for columns.
#' @param header Default assumes the data contains headers. Set to \code{FALSE} if not.
#' @author Kristian D. Olsen
#' @note This function only works on Windows or OSX, and the data-size cannot
#' exceed 128kb in Windows.
#' @export
#' @examples
#' \dontrun{
#' x <- from_clipboard()
#' }

from_clipboard <- function(sep = "\t", header = TRUE) {

  if (on_windows()) {
    file <- "clipboard-128"
  } else if (on_osx()) {
    file <- pipe("pbpaste", "rb")
    on.exit(close(file), add = TRUE)
  } else {
    stop("Writing to clipboard is supported only in Windows or OSX")
  }

  # Read lines
  lines <- suppressWarnings(readLines(file))

  # Workaround for OS X
  if (length(lines) != 1L) {
    lines <- stri_c(lines, collapse = "\n")
  }

  # Check if any of the lines contain the sep
  if (any(stri_detect(lines, regex = stri_c("[", sep, "]")))) {
    lines <- readr::read_delim(lines, delim = sep)
  }

  return(lines)

}

#' Read common data formats
#'
#' A simple wrapper for reading data which currently supports Rdata, sav, txt,
#' csv, csv2 and xlsx. Under the hood, it uses \code{readxl}, \code{readr} and
#' \code{haven}.
#'
#' @param file Path to a Rdata, sav (SPSS), txt, csv, csv2 or xlsx file.
#' @param ... Additional arguments passed to \code{readxl} and \code{readr}. For
#' instance you can use \code{sheet} to specify a xlsx sheet when reading.
#' @param encoding The encoding to use for txt and csv-files.
#' @param decimal The decimal marker in the file. Typically "," in scandinavia,
#' and "." in the U.S.
#' @author Kristian D. Olsen
#' @return A data.frame. If more than one sheet is read from a xlsx file
#' (or you are reading a Rdata file) a list is returned instead.
#' @export
#' @examples
#' x <- read_data(system.file("extdata", "sample.sav", package = "reporttoolDT"))

read_data <- function(file, ..., encoding = "UTF-8", decimal = ".") {

  dots <- list(...)
  file <- clean_path(file)

  if (!file.exists(file)) {
    stop("Path does not exist:\n", file, call. = FALSE)
  }

  # Locale and dots
  loc <- readr::locale(encoding = encoding, decimal_mark = decimal)

  # Pick input-function based on extension
  switch(stri_trans_tolower(tools::file_ext(file)),
         sav = read_spss(file),
         txt = read_flat(file, sep = "\t", loc, dots),
         tsv = read_flat(file, sep = "\t", loc, dots),
         csv = read_flat(file, sep = ",", loc, dots),
         xlsx = read_xlsx(file, dots),
         xls = read_xlsx(file, dots),
         rdata = read_rdata(file),
         stop("Unrecognized input format in:\n", file, call. = FALSE))
}


# Input wrappers ---------------------------------------------------------------

read_spss <- function(file) {

  x <- haven::read_sav(file)

  # BUG in ReadStat (long strings, > 256 characters)
  name <- filename_no_ext(file)
  strings <- file.path(dirname(file), stri_c(name, " (long strings).Rdata"))

  if (file.exists(strings)) {
    strings <- as.data.frame(read_data(strings))
    rows <- match(x$stringID, strings$stringID)
    vars <- intersect(names(strings), names(x))

    x[vars] <- Map(function(d, a) { attr(d, "label") <- attr(a, "label"); d }, strings[rows, vars], x[vars])
    x$stringID <- NULL # Remove string ID when reading
    warning("Found Rdata with long strings in same directory. Joined with data.", call. = FALSE)
  }

  # Return
  x

}

read_rdata <- function(file) {

  # Create an empty environment to load the rdata
  data <- new.env(parent = emptyenv())
  load(file, envir = data)

  # Convert the environment to a list and lowercase names
  data <- as.list(data)

  # Return first element if only one exists
  if (length(data) == 1L) data <- data[[1]]

  data

}

read_flat <- function(file, sep, loc, dots) {

  if (sep == ";") loc$decimal_mark <- ","

  # Update standard args
  args <- list(file = file, delim = sep, locale = loc)
  args <- append(dots, args[!names(args) %in% names(dots)])

  # Read the data
  do.call(readr::read_delim, args)

}

read_xlsx <- function(file, dots) {

  # Get the sheetnames to be read
  wb <- readxl::excel_sheets(file)

  if (!is.null(dots) && "sheet" %in% names(dots)) {
    sheet <- dots$sheet
    if (is.character(sheet)) {
      sheet <- wb[stri_trans_tolower(wb) %in% stri_trans_tolower(sheet)]
    } else if (is.numeric(sheet) || is.integer(sheet)) {
      sheet <- wb[sheet]
    }
    dots <- dots[!names(dots) %in% "sheet"]
  } else {
    sheet <- wb
  }

  # Read data to list
  data <- lapply(sheet, function(x) {
    a <- list(path = file, sheet = x); a <- append(a, dots)
    x <- try(do.call(readxl::read_excel, a), silent = TRUE)
    if (inherits(x, "try-error")) data.frame() else x })

  # Set names
  names(data) <- sheet

  # If only one sheet was read, return a data.frame instead
  if (length(data) == 1L) data <- data[[1]]

  # Return
  data

}
