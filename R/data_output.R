#' Write to Windows/OSX clipboard
#'
#' Wrapper for writing to windows/OSX clipboards with the most-used defaults for a
#' scandinavian locale.
#'
#' @param x The data or text to write.
#' @param ... Further arguments passed to \code{write.table}.
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
    stop("Writing to clipboard is supported only in Windows or OSX")
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
#' by the extension given in \code{file}. Flat files are written with \code{readr},
#' and the encoding is always \code{UTF-8}. For xlsx, the function uses \code{to_excel}
#' (which in turn uses \code{openxlsx}).
#'
#' @param x The data to be written. (\code{data.frame}, \code{list} or \code{survey}).
#' \code{matrix} and \code{table} will be coerced to a \code{data.frame}.
#' @param file Path and filename of output.
#' @param ... Further arguments passed to \code{readr}, \code{openxlsx} or \code{haven}.
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
  dots <- list(...)
  ext <- tolower(tools::file_ext(file))

  # Excel/Rdata writer expects a named list as input
  if ("sheet" %in% names(dots)) {
    name <- dots$sheet
  } else {
    name <- basename_sans_ext(file)
  }

  switch(ext,
         rdata = write_rdata(setNames(list(x), name), file),
         rda = write_rdata(setNames(list(x), name), file),
         xlsx = write_xlsx(setNames(list(x), name), file, dots),
         sav = write_spss(x, file),
         txt = write_flat(x, file, delim = delim %||% "\t", dots),
         tsv = write_flat(x, file, delim = delim %||% "\t", dots),
         csv = write_flat(x, file, delim = delim %||% ",", dots),
         stop("Unrecognized output format: ", ext))

  # Suppress printing
  invisible()
}

#' @rdname write_data
#' @export
write_data.list <- function(x, file, ...) {
  dots <- list(...)
  ext <- tolower(tools::file_ext(file))
  if (!ext %in% c("xlsx", "rdata")) {
    stop("Lists can only be written to .xlsx or .Rdata files.")
  }

  # All list elements must be named
  unnamed <- is.null(names(x)) || any(names(x) == "")
  dupes <- duplicated(names(x))
  if (unnamed) {
    stop("All list elements must be named.")
  } else if (any(dupes)) {
    stop("List names should not contain duplicates.")
  }

  # The list should only contain data.frame/matrix
  is_df <- vapply(x, is.data.frame, logical(1))
  if (!all(is_df)) {
    is_matrix <- vapply(x, is.matrix, logical(1))
    if (!all(is_df | is_matrix)) {
      stop("All list items must be a data.frame or matrix.")
    }
  }

  switch(ext, rdata = write_rdata(x, file), xlsx = write_xlsx(x, file, dots))

  # Suppress printing
  invisible()
}

#' @export
write_data.matrix <- function(x, file, ...) {
  write_data(as.data.frame(x, stringsAsFactors = FALSE), file, ...)
}

#' @export
write_data.table <- write_data.matrix

#' @export
write_data.character <- function(x, file, ...) {
  # Use readr by default? Encoding?
  stop("Work in progress. Not ready yet.")
}

#' @rdname write_data
#' @export
write_data.ggplot <- function(x, file, ...) {
  if (!requireNamespace("ggplot2")) {
    stop("'ggplot2' required to write plots.")
  }
  ggplot2::ggsave(filename = file, plot = x, ...)
}

# Output wrappers --------------------------------------------------------------

write_spss <- function(data, file) {
  if (!any_labelled(data)) warning("No labelled variables found in data.")

  # WORKAROUND: Haven/ReadStat cannot write strings that exceed 256 characters.
  # read_/write_data works around this by writing columns to a separate .Rdata file,
  # and truncating the strings itself before attempting to write - to avoid crashes.
  is_character <- vapply(data, is.character, logical(1))
  if (any(is_character)) {
    is_long <- vapply(data[is_character], function(x) {
      max(nchar(x, keepNA = TRUE), na.rm = TRUE) > 250  } , logical(1))

    if (any(is_long)) {
      columns <- names(data)[is_long]
      name <- basename_sans_ext(file)
      spath <- file.path(dirname(file), paste0(name, " (long strings).Rdata"))

      # We need an ID to match against when reading in again.
      data$string_id <- 1:nrow(data)

      # Write the full-length strings separately and truncate in original data
      write_rdata(list("x" = data[c(columns, "string_id")]), spath)
      data[columns] <- lapply(data[columns], function(x) {
        oa = attributes(x); x <- substr(x, 1L, 250L); attributes(x) <- oa; x
        })
      warning("Detected long strings (> 250) in data. Stored as standalone:\n", spath, call. = FALSE)
    }

  }

  haven::write_sav(data, path = file)

}

write_rdata <- function(data, file) {
  save(list = names(data), file = file, envir = list2env(data, parent = emptyenv()))
}

write_flat <- function(data, file, delim, dots) {

  # Update default args.
  args <- list(x = data, path = file, delim = delim)
  if (!is.null(dots)) {
    name <- setdiff(names(dots), names(args))
    args <- append(dots[name], args)
  }

  do.call(readr::write_delim, args)

}

write_xlsx <- function(data, file, dots) {
  # Load workbook if it already exists. Create if not.
  if (file.exists(file)) {
    wb <- openxlsx::loadWorkbook(file)
  } else {
    wb <- openxlsx::createWorkbook()
  }

  # Update default args. User should use to_excel if they want to
  # change these defaults.
  args <- list(title = NULL, format = FALSE, row = 1L, append = FALSE)
  if (!is.null(dots)) {
    name <- setdiff(names(dots), c(names(args), "sheet"))
    args <- append(dots[name], args)
  }

  # openxlsx Workbooks are mutable, so we don't have to assign results.
  lapply(names(data), function(name) {
    file_arg <- list(df = data[[name]], wb = wb, sheet = name)
    all_args <- append(file_arg, args)
    do.call(to_excel, all_args)
    })

  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)

}