# #' Write to Windows/OSX clipboard
# #'
# #' Wrapper for writing to windows/OSX clipboards with the most-used defaults for a
# #' scandinavian locale.
# #'
# #' @param x The data or text to write.
# #' @param encoding The encoding to use when writing.
# #' @author Kristian D. Olsen
# #' @note This function only works on Windows or OSX, and the data-size cannot
# #' exceed 128kb in Windows.
# #' @export
# #' @examples
# #' x %>% to_clipboard()
#
# to_clipboard <- function(x, encoding = "") {
#
#   if ((Sys.info()["sysname"] == "Windows")) {
#     file <- "clipboard-128"
#     if (object.size(x) > 120000) {
#       stop("The data is too large to write to windows clipboard", call. = FALSE)
#     }
#   } else if (Sys.info()["sysname"] == "Darwin") {
#     file <- pipe("pbcopy", "w")
#     on.exit(close(file), add = TRUE)
#   } else {
#     stop("Writing to clipboard is supported only in Windows or OSX")
#   }
#
#   if (is.character(x)) {
#     writeLines(x, file)
#   } else {
#
#     cols <- if (is.data.frame(x)) TRUE else FALSE
#     utils::write.table(x = x, file = file, sep = "\t", na = "", dec = ",",
#                        row.names = FALSE, col.names = cols, fileEncoding = encoding)
#
#   }
#
# }
#
# #' Write data to sheet (in an openxlsx workbook)
# #'
# #' This function (a very thin wrapper for openxlsx functions) that writes/appends
# #' a data.frame to the specified sheet in a loaded openxlsx workbook.
# #'
# #' @param df The data to write.
# #' @param wb A loaded openxlsx workbook (use openxlsx::loadWorkbook)
# #' @param title The title to give to the table (only used if style = TRUE).
# #' @param sheet Name of the sheet to write to, will be created if it does not exist.
# #' @param row Optional: Also specify the startingrow for writing data.
# #' @param format_style Set to FALSE if you do not want styling for the data.
# #' @param format_values Set to FALSE and no formatting will be applied based on
# #' variable type. With TRUE, character columns will be left justified, numeric
# #' will have 1 decimal place, integer 0, and columns with values between 1 and 0
# #' as percentages.
# #' @param append Whether or not the function should append or clean the
# #' sheet of existing data before writing.
# #' @author Kristian D. Olsen
# #' @return A list containing data.frames matching the sheets in the .xlsx file.
# #' If only one sheet is read, the function returns a data.frame instead.
# #' @note This function requires openxlsx.
# #' @export
# #' @examples
# #' wb <- openxlsx::loadWorkbook("test.xlsx")
# #' x %>% to_sheet(wb, sheet = "test", append = FALSE)
# #' openxlsx::saveWorkbook(wb, "test.xlsx", overwrite = TRUE)
#
# to_sheet <- function(df, wb, title = "Table", sheet = "tables", row = 1L,
#                      format_style = TRUE, format_values = TRUE, append = TRUE) {
#
#   if (!inherits(wb, "Workbook")) {
#     stop("wb argument must be a (loaded) openxlsx workbook")
#   }
#
#   # Check input
#   if (!is.character(sheet) || length(sheet) != 1L) {
#     stop("The sheet has to be a string of length 1 (not an index).", call. = FALSE)
#   }
#
#   # See if sheet exists
#   sheet_exists <- sheet %in% openxlsx::sheets(wb)
#
#   # Get last row if sheet exists, or create if it does not.
#   if (sheet_exists && isTRUE(append)) {
#     row <- 2L + nrow(openxlsx::read.xlsx(wb, sheet = sheet, colNames = FALSE, skipEmptyRows = FALSE))
#   } else if (sheet_exists) {
#     openxlsx::removeWorksheet(wb, sheet)
#     openxlsx::addWorksheet(wb, sheetName = sheet)
#   } else {
#     openxlsx::addWorksheet(wb, sheetName = sheet)
#   }
#
#   # Set table_row to be the last found row
#   table_row <- row
#
#   # Add data to the workbook
#   if (is.null(names(df)) || identical(names(df), character(0))) {
#     warning(sheet, ": No columnames in data. An empty sheet was created", call. = FALSE)
#   } else {
#
#     # When styling the title must be written first (and convert df names to titles)
#     if (isTRUE(format_style)) {
#       openxlsx::writeData(wb, sheet, title, startRow = row)
#       names(df) <- stri_trans_totitle(names(df), type = "sentence")
#       table_row <- row + 1
#     }
#
#     # Write the data.frame
#     openxlsx::writeData(wb, sheet, df, startRow = table_row)
#   }
#
#   # Apply additional formatting if desired
#   if (isTRUE(format_style) || isTRUE(format_values)) {
#
#     format_xlsx(df, wb, sheet, table_row, style = format_style, values = format_values)
#
#   }
#
#   # Check the dimensions of what was written
#   n_row <- dim(df)
#   n_row <- if(is.null(n_row)) 0 else n_row[1]
#
#   # Invisibly return the rows that have been written to
#   invisible(setNames(c(table_row, table_row + n_row), c("first", "last")))
#
# }
#
# #' Write common file formats
# #'
# #' A simple wrapper for writing common data formats. The format is determined
# #' by the extension given in \code{file}. Flat files are written with \code{readr},
# #' and the encoding is always \code{UTF-8}. For xlsx, the function uses \code{to_sheet}
# #' (which in turn uses \code{openxlsx}).
# #'
# #' @param x The data to be written. (\code{data.frame}, \code{list} or \code{survey}).
# #' @param file Path.
# #' @author Kristian D. Olsen
# #' @note Use \code{lapply} to write a list of data to flat files (csv, txt etc).
# #' @export
# #' @examples
# #' write_data(x, file = "test.xlsx")
#
# write_data <- function(x, file, ...) {
#
#   # Gather dots
#   dots <- list(...)
#
#   # Get file information
#   file <- clean_path(file)
#   ext <- stri_trans_tolower(tools::file_ext(file))
#   name <- filename_no_ext(file)
#
#   # Convert matrix to data.frame
#   if (is.matrix(x)) x <- as_data_frame(x)
#
#   # Check if it is a survey and convert depending on output format
#   if (is.survey(x) && ext == "sav") {
#       x <- to_labelled(x)$df
#   } else if (is.survey(x) && ext == "xlsx") {
#     is_date <- vapply(x$df, inherits, what = "Date", logical(1))
#     x$df[is_date] <- lapply(x$df[is_date], as.character)
#     names(x) <- ordered_replace(names(x), default$structure$survey, default$structure$sheet)
#   } else if (is.data.frame(x)) {
#     if (ext %in% c("xlsx", "rdata")) {
#       x <- if ("sheet" %in% names(dots)) setNames(list(x), dots[["sheet"]]) else setNames(list(x), name)
#     }
#   } else if (!is.list(x)) {
#     stop("This function expects a matrix, data.frame, list or survey.", call. = FALSE)
#   }
#
#   # Only xlsx and rdata supports a list of output
#   supports_list <- ext %in% c("xlsx", "rdata")
#   if (is.list2(x) && !supports_list) {
#       stop("Use lapply to write lists that are not survey objects when output is not xlsx.", call. = FALSE)
#   }
#
#   # Use extension to write correct format
#   switch(ext,
#          sav = write_spss(x, file),
#          rdata = write_rdata(x, file),
#          xlsx = write_xlsx(x, file, dots),
#          txt = write_flat(x, file, delim = "\t", dots),
#          tsv = write_flat(x, file, delim = "\t", dots),
#          csv = write_flat(x, file, delim = ",", dots),
#          stop("Unrecognized output format: ", ext))
#
#   invisible()
# }
#
# # Output wrappers --------------------------------------------------------------
#
# write_spss <- function(data, file) {
#
#   if (!is.spss(data)) {
#     warning("No labelled variables found.", call. = FALSE)
#   }
#
#   # BUG in ReadStat (long strings, > 256 characters)
#   is_character <- vapply(data, is.character, logical(1))
#
#   if (any(is_character)) {
#     strings <- vapply(data[is_character], function(x) max(stri_length(x), na.rm = TRUE) > 250, logical(1))
#     strings <- names(strings[strings])
#
#     if (length(strings)) {
#       name <- filename_no_ext(file)
#       spath <- file.path(dirname(file), stri_c(name, " (long strings).Rdata"))
#
#       # Add stringID to data
#       data$stringID <- 1:nrow(data)
#
#       # Write strings separately and shorten in original data
#       write_rdata(list("x" = data[c(strings, "stringID")]), spath)
#       data[strings] <- lapply(data[strings], function(x) {
#         oa = attributes(x); x <- stri_sub(x, to = 250); attributes(x) <- oa; x
#         })
#       warning("Found long strings (> 250) in data. Writing as separate Rdata.", call. = FALSE)
#     }
#
#   }
#
#   haven::write_sav(data, path = file)
#
# }
#
# write_rdata <- function(data, file) {
#
#   save(list = names(data), file = file, envir = list2env(data, parent = emptyenv()))
#
# }
#
# write_flat <- function(data, file, delim, dots) {
#
#   # Update standard args
#   args <- list(x = data, path = file, delim = delim)
#   args <- append(dots, args[!names(args) %in% names(dots)])
#
#   # Read the data
#   do.call(readr::write_delim, args)
#
# }
#
# write_xlsx <- function(data, file, dots) {
#
#   # If the file exists, load and write to it
#   if (file.exists(file)) {
#     wb <- openxlsx::loadWorkbook(file)
#   } else {
#     wb <- openxlsx::createWorkbook()
#   }
#
#   # Update standard args
#   args <- list(row = 1L, format_style = FALSE, format_values = FALSE, append = FALSE)
#   args <- append(dots[!names(dots) %in% "sheet"], args[!names(args) %in% names(dots)])
#
#   lapply(names(data), function(nm, x, wb) {
#     a <- list(df = x[[nm]], wb = wb, sheet = nm); a <- append(a, args)
#     do.call(to_sheet, a)}, data, wb)
#
#   openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
#
# }