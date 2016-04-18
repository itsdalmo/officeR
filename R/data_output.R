#' Write to Windows/OSX clipboard
#'
#' Wrapper for writing to windows/OSX clipboards with the most-used defaults for a
#' scandinavian locale.
#'
#' @param x The data or text to write.
#' @param ... Further arguments passed to \code{\link{write.table}}.
#' @author Kristian D. Olsen
#' @note This function only works on Windows or OSX, and the data-size cannot
#' exceed 128kb in Windows.
#' @export
#' @examples
#' \dontrun{
#' # Only works on Windows and OSX
#' df <- data.frame("String" = c("A", "B"), "Int" = c(1:2L), "Percent" = c(0.5, 0.75))
#' to_clipboard(df)
#' }

write_clipboard <- function(x, ...) {
  dots <- list(...)
  if (on_windows()) {
    file <- "clipboard-128"
    if (object.size(x) > 120000) {
      stop("The data is too large to write to windows clipboard", call. = FALSE)
    }
  } else if (on_osx()) {
    file <- pipe("pbcopy", "w")
    on.exit(close(file), add = TRUE)
  } else {
    stop("Writing to clipboard is supported only in Windows and OSX")
  }

  if (is.character(x)) {
    writeLines(x, file)
  } else {
    args <- list(
      x = x, file = file, sep = "\t", na = "", dec = ",",
      row.names = FALSE, fileEncoding = "", col.names = is.data.frame(x))

    # User arguments
    if (!is.null(dots)) {
      args <- c(args, dots[setdiff(names(dots), c("x", "file"))])
      args <- args[!duplicated(args, fromLast = TRUE)]
    }

    do.call(utils::write.table, args)
  }

}

#' @rdname write_clipboard
#' @export
to_clipboard <- write_clipboard

#' Write common file formats
#'
#' A simple wrapper for writing common data formats. The format is determined
#' by the extension given in \code{file}. Flat files are written with \pkg{readr},
#' and the encoding is always \code{UTF-8}. For xlsx, the function uses
#' \code{\link{to_excel}} (which in turn uses \pkg{openxlsx}).
#'
#' @param x The data to be written. (\code{data.frame} or \code{list}.
#' \code{matrix} and \code{table} will be coerced to a \code{data.frame}.
#' @param file Path and filename of output.
#' @param ... Further arguments passed to \pkg{readr}, \pkg{openxlsx} or \pkg{haven}.
#' @param delim The delimiter to use when writing flat files.
#' @author Kristian D. Olsen
#' @note Use \code{lapply} to write a list of data to flat files (csv, txt etc).
#' @export
#' @examples
#' \dontrun{
#' df <- data.frame("String" = c("A", "B"), "Int" = c(1:2L), "Percent" = c(0.5, 0.75))
#' write_data(df, file = "example data.csv")
#' }

write_data <- function(x, file, ..., delim = NULL) {
  file <- clean_path(file)
  UseMethod("write_data")
}

#' @rdname write_data
#' @export
write_data.data.frame <- function(x, file, ..., delim = NULL) {
  ext <- tolower(tools::file_ext(file))
  switch(ext,
         rdata = write_rdata(x, file, ...),
         rda = write_rdata(x, file, ...),
         xlsx = write_xlsx(x, file, ...),
         sav = write_spss(x, file, ...),
         txt = write_flat(x, file, delim = delim %||% "\t", ...),
         tsv = write_flat(x, file, delim = delim %||% "\t", ...),
         csv = write_flat(x, file, delim = delim %||% ",", ...),
         stop("Unrecognized output format (for data.frame): ", ext))

  # Suppress printing
  invisible()
}

#' @rdname write_data
#' @export
write_data.list <- function(x, file, ...) {
  ext <- tolower(tools::file_ext(file))
  switch(ext,
         rdata = write_rdata(x, file, ...),
         rda = write_rdata(x, file, ...),
         xlsx = write_xlsx(x, file, ...),
         stop("Unrecognized output format (for list): ", ext))

  # Suppress printing
  invisible()
}

#' @export
write_data.matrix <- function(x, file, ...) {
  write_data(as.data.frame(x, stringsAsFactors = FALSE), file, ...)
}

#' @export
write_data.table <- write_data.matrix

#' @rdname write_data
#' @export
write_data.ggplot <- function(x, file, ...) {
  if (!requireNamespace("ggplot2")) {
    stop("'ggplot2' required to write plots.")
  }
  ggplot2::ggsave(filename = file, plot = x, ...)
}

# Output wrappers --------------------------------------------------------------

write_spss <- function(data, file, ...) {
  if (!any_labelled(data)) warning("No labelled variables found in data.")

  # WORKAROUND: Haven/ReadStat cannot write strings that exceed 256 characters.
  # read_/write_data works around this by writing columns to a separate .Rdata file,
  # and truncating the strings itself before attempting to write - to avoid crashes.
  is_character <- vapply(data, is.character, logical(1))
  if (any(is_character)) {
    is_long <- vapply(data[is_character], function(x) {
      max(nchar(x, keepNA = TRUE), na.rm = TRUE) > 250  } , logical(1))

    if (any(is_long)) {
      name <- basename_sans_ext(file)
      spath <- file.path(dirname(file), paste0(name, " (long strings).Rdata"))

      # We need an ID to match against when reading in again.
      columns <- names(data)[is_long]
      data$string_id <- 1:nrow(data)

      # Write the full-length strings separately and truncate in original data
      write_rdata(data[c(columns, "string_id")], spath)
      data[columns] <- lapply(data[columns], function(x) {
        oa = attributes(x); x <- substr(x, 1L, 250L); attributes(x) <- oa; x
        })
      warning("Detected long strings (> 250) in data. Stored as standalone:\n", spath, call. = FALSE)
    }

  }

  haven::write_sav(data, path = file, ...)

}

write_rdata <- function(data, file, ...) {
  save(data, file = file, ...)
}

write_flat <- function(data, file, delim, ...) {
  readr::write_delim(data, file, delim = delim, ...)
}

write_xlsx <- function(data, file, ...) {
  # to_excel expects a named list.
  if (!is_list(data)) {
    data <- setNames(list(data), basename_sans_ext(file))
  } else {
    if (!is_named(data))
      stop("All elements must be named when writing a list to excel.")
    if (any(duplicated(names(data))))
      stop("Lists cannot contain duplicated names when writing to excel.")
  }

  # Load existing workbook, or create new.
  if (file.exists(file)) {
    wb <- openxlsx::loadWorkbook(file)
  } else {
    wb <- openxlsx::createWorkbook()
  }

  # Add data to wb and write.
  for (sheet in names(data)) {
    to_excel(data[[sheet]], wb, sheet = sheet, title = NULL, format = FALSE, append = FALSE)
  }
  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)

}