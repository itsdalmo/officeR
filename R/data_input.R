#' Read from Windows/OSX clipboard
#'
#' Thin wrapper for reading from windows/OSX clipboards with the most-used defaults.
#' The function first reads in the lines, checks if the delimiter is present in the lines
#' and then converts it to a data.frame.
#'
#' @param delim The delimiter for columns.
#' @param ... Further arguments passed to \code{\link[readr]{read_delim}}.
#' @author Kristian D. Olsen
#' @note This function only works on Windows or OSX, and the data-size cannot
#' exceed 128kb in Windows.
#' @export
#' @examples
#' \dontrun{
#' # Only works on Windows and OSX
#' df <- data.frame("String" = c("A", "B"), "Int" = c(1:2L), "Percent" = c(0.5, 0.75))
#' to_clipboard(df)
#' x <- from_clipboard()
#'
#' # All equal - except attributes
#' # (readr attaches attr "problems")
#' all.equal(x, df, check.attributes = FALSE)
#' }

read_clipboard <- function(delim = "\t", ...) {
  dots <- list(...)
  if (on_windows()) {
    file <- "clipboard-128"
  } else if (on_osx()) {
    file <- pipe("pbpaste", "rb")
    on.exit(close(file), add = TRUE)
  } else {
    stop("Writing to clipboard is supported only in Windows or OSX")
  }

  # Read lines and convert input to UTF-8.
  lines <- suppressWarnings(readLines(file))
  lines <- iconv(lines, from = "", to = "UTF-8")

  # Can return multiple lines.
  if (length(lines) != 1L) {
    lines <- paste0(lines, collapse = "\n")
  }

  # Check if any of the lines contain the sep
  if (any(grepl(paste0("[", delim, "]"), lines))) {
    lines <- readr::read_delim(lines, delim = delim, ...)
  }

  return(lines)

}

#' @rdname read_clipboard
#' @export
from_clipboard <- read_clipboard

#' Read common data formats
#'
#' A simple wrapper for reading data which currently supports Rdata, sav, txt,
#' csv, csv2 and xlsx. Under the hood, it uses \pkg{readxl}, \pkg{readr} and
#' \pkg{haven}.
#'
#' @param file Path to a file with a supported extension.
#' @param ... Additional arguments passed to \pkg{readxl} and \pkg{readr}. For
#' instance you can use \code{sheet} to specify a xlsx sheet when reading.
#' @param delim The deliminator used for flat files.
#' @param encoding The current encoding for the file. Defaults to \code{UTF-8}.
#' @author Kristian D. Olsen
#' @return A data.frame. If more than one sheet is read from a xlsx file
#' (or you are reading a Rdata file) a list is returned instead.
#' @export
#' @examples
#' \dontrun{
#' df <- data.frame("String" = c("A", "B"), "Int" = c(1:2L), "Percent" = c(0.5, 0.75))
#' write_data(df, file = "example data.xlsx")
#' read_data(df, file = "example data.xlsx")
#' }

read_data <- function(file, ...) UseMethod("read_data")

#' @rdname read_data
#' @export
read_data.default <- function(file, ..., delim = NULL, encoding = "UTF-8") {
  dots <- list(...)
  file <- clean_path(file)
  if (!file.exists(file)) stop("Path does not exist:\n", file)
  # Pick input-function based on extension
  switch(tolower(tools::file_ext(file)),
         sav = read_spss(file),
         txt = read_flat(file, delim = delim %||% "\t", encoding, dots),
         tsv = read_flat(file, delim = delim %||% "\t", encoding, dots),
         csv = read_flat(file, delim = delim %||% ",", encoding, dots),
         xlsx = read_xlsx(file, dots),
         xls = read_xlsx(file, dots),
         rdata = read_rdata(file),
         stop("Unrecognized input format in:\n", file, call. = FALSE))
}

# Input wrappers ---------------------------------------------------------------

read_spss <- function(file) {
  x <- haven::read_sav(file)

  # WORKAROUND: See explanation for this under write_spss.
  name <- basename_sans_ext(file)
  sfile <- file.path(dirname(file), paste0(name, " (long strings).Rdata"))

  if (file.exists(sfile)) {
    strings <- as.data.frame(read_data(sfile))
    rows <- match(x$string_id, strings$string_id)
    vars <- intersect(names(strings), names(x))

    x[vars] <- Map(function(d, a) { attr(d, "label") <- attr(a, "label"); d }, strings[rows, vars], x[vars])
    x$string_id <- NULL # Remove string ID when reading
    warning("Found Rdata with long strings in same directory. Joined with data.")
  }

  # Return
  x

}

read_rdata <- function(file) {

  # Create an empty environment to load the rdata
  data <- new.env(parent = emptyenv())
  load(file, envir = data)
  data <- as.list(data)

  # Return first element if only one exists
  if (length(data) == 1L) data <- data[[1]]

  data

}

read_flat <- function(file, delim, encoding, dots) {
  # readr expects a locale
  dec <- if (delim != ";") "." else ","
  loc <- readr::locale(encoding = encoding, decimal_mark = dec)

  # Update standard args
  args <- list(file = file, delim = delim, locale = loc)
  if (!is.null(dots)) {
    name <- setdiff(names(dots), names(args))
    args <- append(dots[name], args)
  }

  # Read the data
  do.call(readr::read_delim, args)

}

read_xlsx <- function(file, dots) {

  # Get the sheetnames to be read
  wb <- readxl::excel_sheets(file)

  if (!is.null(dots) && "sheet" %in% names(dots)) {
    sheet <- dots$sheet
    if (is.character(sheet)) {
      sheet <- wb[tolower(wb) %in% tolower(sheet)]
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
