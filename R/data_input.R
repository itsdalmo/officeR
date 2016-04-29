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
#' @param ... Additional arguments passed to underlying functions.
#' (Currently, \code{\link[readr]{read_delim}} and \code{\link[readxl]{read_excel}}.)
#' @param sheet Specify one or more sheets when reading excel files.
#' @param delim The deliminator used for flat files.
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
read_data.default <- function(file, ..., sheet = NULL, delim = NULL) {
  file <- clean_path(file)
  if (!file.exists(file))
    stop("Path does not exist:\n", file)

  switch(tolower(tools::file_ext(file)),
         sav = read_spss(file, ...),
         txt = read_flat(file, delim = delim %||% "\t", ...),
         tsv = read_flat(file, delim = delim %||% "\t", ...),
         csv = read_flat(file, delim = delim %||% ",", ...),
         xlsx = read_xlsx(file, sheet, ...),
         xls = read_xlsx(file, sheet, ...),
         rda = read_rdata(file, ...),
         rdata = read_rdata(file, ...),
         stop("Unrecognized input format in:\n", file, call. = FALSE))
}

# Input wrappers ---------------------------------------------------------------

read_spss <- function(file, ...) {
  x <- haven::read_sav(file)

  # WORKAROUND: See explanation for this under write_spss.
  name <- basename_sans_ext(file)
  sfile <- file.path(dirname(file), paste0(name, " (long strings).Rdata"))

  if (file.exists(sfile)) {
    warn <- "Found Rdata with long strings."
    strings <- as.data.frame(read_data(sfile), stringsAsFactors = FALSE)

    # Find the "string_id" variable. Used to be "stringID".
    str_id <- names(strings)[names(strings) %in% c("string_id", "stringID")]
    if (!length(str_id)) {
      warn <- paste(warn, "Could not join with data. Rdata does not contain 'string_id' or 'stringID'.")
    } else if (length(str_id) > 1L) {
      warn <- paste(warn, "Could not join with data. Rdata contains multiple string identifiers:\n",
                    paste0("'", str_id, "'", collapse = ","), collapse = " ")
    } else if (!str_id %in% names(x)) {
      warn <- paste(warn, "Could not join with data. The SPSS file does not contain:\n",
                    paste0("'", str_id, "'"), collapse = " ")
    } else {
      warn <- paste(warn, "Joined with data.")
      rows <- match(x[[str_id]], strings[[str_id]])
      vars <- intersect(names(strings), names(x))
      x[vars] <- Map(function(s, d) { attr(s, "label") <- attr(d, "label"); s }, strings[rows, vars], x[vars])
      x[[str_id]] <- NULL # Remove string ID when reading
    }
    warning(warn)
  }

  # Return
  x

}

read_rdata <- function(file, ...) {
  # Use an empty environment when loading Rda/Rdata.
  # (Don't want objects attached in current env.)
  env <- new.env(parent = emptyenv())
  load(file, envir = env)

  # write_data only allows 1 object per rdata file,
  # save has no such restrictions.
  data <- as.list(env)
  if (length(data) == 1L) {
    data <- data[[1L]]
  } else {
    warning("Multiple objects stored in file. Returning list.")
  }

  data

}

read_flat <- function(file, delim, encoding = "UTF-8", decimal = ".", ...) {
  loc <- readr::locale(encoding = encoding, decimal_mark = if (delim == ",") "." else decimal)
  readr::read_delim(file, delim, locale = loc, ...)
}

read_xlsx <- function(file, sheet, ...) {
  # Read requested sheets (or all if none are specified.)
  wb <- readxl::excel_sheets(file)
  if (!is.null(sheet)) {
    if (is.character(sheet)) {
      sheets <- wb[tolower(wb) %in% tolower(sheet)]
    } else if (is.numeric(sheet)) {
      sheets <- wb[sheet]
    }
  } else {
    sheets <- wb
  }

  # Read using try. Print errors.
  data <- lapply(sheets, function(s) {
    x <- try(readxl::read_excel(file, sheet = s, col_names = TRUE, ...), silent = TRUE)
    if (inherits(x, "try-error")) {
      warning("Failed to read '", s, "'. Returning empty data.frame.")
      data.frame()
    } else {
      x
    }
  })

  names(data) <- sheets

  # If only one sheet was read, return a data.frame instead
  if (length(data) == 1L) data <- data[[1]]

  # Return
  data

}
