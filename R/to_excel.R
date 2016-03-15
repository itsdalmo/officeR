#' Pass data to Excel workbooks
#'
#' \code{to_excel} allows you to pass R objects (primarily a \code{data.frame}) to
#' an open \code{Workbook}, and write it later with \code{write_data}. The workbook
#' can be created by calling \code{excel_workbook}, which is itself a wrapper for
#' \code{openxlsx::createWorkbook()}.
#'
#' @param df A \code{data.frame}, \code{table} or \code{matrix}.
#' @param wb A \code{Workbook}.
#' @param ... Arguments passed to \code{to_excel.data.frame} (See below).
#' @param title The title to give to the table. Must be \code{NULL} if you do not
#' want the table to be formatted.
#' @param sheet The sheet you want to write the data to.
#' @param format Format values and apply the default template to the table output.
#' @param append Whether or not the function should append to or replace the
#' sheet before writing.
#' @param row Specify the startingrow when writing data to a new sheet.
#' @param col Start column. Same as for row.
#' @author Kristian D. Olsen
#' @note This function requires \code{openxlsx}.
#' @export
#' @examples
#' if (require(openxlsx)) {
#'  wb <- excel_workbook()
#'  df <- data.frame("String" = c("A", "B"), "Int" = c(1:2L), "Percent" = c(0.5, 0.75))
#'
#'  # The workbook is mutable, so we don't have to assign result.
#'  to_excel(df, wb, title = "Example data", sheet = "Example", append = FALSE)
#'
#'  # Data is first argument, so we can use it with dplyr.
#'  # df %>% to_excel(wb, title = "Example dplyr", sheet = "Example", append = TRUE)
#'
#'  # Save the data
#'  write_data(wb, "Example table.xlsx", overwrite = TRUE)
#' }

to_excel <- function(df, wb, ...) {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("This function requires 'openxlsx'.")
  } else if (!inherits(wb, "Workbook")) {
    stop ("'wb' should be a Workbook. See help(to_excel).")
  } else if (!identical(attr(class(wb), "package"), "openxlsx")) {
    stop("Unknown type of 'workbook'.")
  }
  UseMethod("to_excel")
}

#' @rdname to_excel
#' @export
excel_workbook <- function() {
  if (!requireNamespace("openxlsx")) {
    stop("'openxlsx' required to create a Workbook.")
  }
  openxlsx::createWorkbook()
}

#' @rdname to_excel
#' @export
to_excel.data.frame <- function(df, wb, title = " ", sheet = "tables",
                                format = !is.null(title), append = TRUE, row = 1L, col = 1L) {

  # Check input
  if (!is.string(sheet)) {
    stop("The sheet has to be a string (character(1)).")
  } else if (!is.null(title) && !is.string(title)) {
    stop("Title has to be either NULL or a string.")
  } else if (!is.integer(row) || !is.integer(col)) {
    stop("'row' and 'col' must be an integer.")
  }

  # Map out which cells need to be written to and their type.
  index <- list(
    # Don't need to subract 1L from rows because of the header (colnames).
    columns = c(start = col, end = ncol(df) + col - 1L),
    rows    = c(start = row, end = nrow(df) + row),
    format  = excel_formats(df)
  )

  # Get last row if sheet exists, or create if it does not.
  if (sheet %in% openxlsx::sheets(wb)) {
    if (append) {
      row <- nrow(openxlsx::read.xlsx(wb, sheet = sheet, colNames = FALSE, skipEmptyRows = FALSE))
      index$rows <- index$rows + 2L # Space between previous table.
    } else {
      openxlsx::removeWorksheet(wb, sheet)
      openxlsx::addWorksheet(wb, sheetName = sheet)
    }
  } else {
    openxlsx::addWorksheet(wb, sheetName = sheet)
  }

  # Return early if there are no colummnames in the data.
  if (is.null(names(df)) || identical(names(df), character(0))) {
    warning("No columnames in data. An empty sheet was created: ", paste0("'", sheet, "'"))
    return()
  }

  if (format) {
    # Write title on the first row
    openxlsx::writeData(wb, sheet, title %||% " ", startRow = index$rows[1], startCol = index$columns[1])
    index$rows[2] <- index$rows[2] + 1L # +1 for last row written to.

    # Titlecase and write data
    substr(names(df), 1L, 1L) <- toupper(substr(names(df), 1L, 1L))
    openxlsx::writeData(wb, sheet, df, startRow = index$rows[1] + 1L, startCol = index$columns[1])

    # Apply template and formatting
    format_excel_table(wb, sheet, index)
    format_excel_columns(wb, sheet, index)
  } else {
    openxlsx::writeData(wb, sheet, df, startRow = row)
  }

  invisible(index)

}

to_excel.matrix <- function(df, wb, ...) {
  warning("Coercing ", class(df), " to data.frame.")
  to_excel(as.data.frame(df, stringsAsFactors = FALSE), wb, ...)
}

to_excel.table <- to_excel.matrix

# Set column formats in excel --------------------------------------------------
format_excel_columns <- function(wb, sheet, cell) {

  # Format all rows except title and header (always present when formatting).
  # Columns derived from a character vector. Need to add first column from cell
  # and subtract 1 to get correct index.
  rows <- (cell$rows[1] + 2L):cell$rows[2]

  if (any(cell$format == "numeric")) {
    cols <- which(cell$format == "numeric") + cell$columns[1] - 1L
    openxlsx::addStyle(wb, sheet, excel_numeric, rows = rows, cols = cols, gridExpand = TRUE, stack = TRUE)
  }

  if (any(cell$format == "integer")) {
    cols <- which(cell$format == "integer") + cell$columns[1] - 1L
    openxlsx::addStyle(wb, sheet, excel_integer, rows = rows, cols = cols, gridExpand = TRUE, stack = TRUE)
  }

  if (any(cell$format == "percent")) {
    cols <- which(cell$format == "percent") + cell$columns[1] - 1L
    openxlsx::addStyle(wb, sheet, excel_percent, rows = rows, cols = cols, gridExpand = TRUE, stack = TRUE)
  }

  if (any(cell$format %in% c("character", "factor"))) {
    cols <- which(cell$format %in% c("character", "factor")) + cell$columns[1] - 1L
    openxlsx::addStyle(wb, sheet, excel_character, rows = rows, cols = cols, gridExpand = TRUE, stack = TRUE)
  }
}

excel_numeric <- openxlsx::createStyle(numFmt = "0.0")
excel_integer <- openxlsx::createStyle(numFmt = "0")
excel_percent <- openxlsx::createStyle(numFmt = "0%")
excel_character <- openxlsx::createStyle(halign = "left")

# Simplified "class" for columns -----------------------------------------------
excel_formats <- function(df) {
  type <- vapply(df, function(x) class(x)[1], character(1))

  type[vapply(df, is.factor, logical(1))] <- "factor"
  type[vapply(df, is_percent, logical(1))] <- "percent"
  type[vapply(df, is_date, logical(1))] <- "date"

  # Return character vector
  type

}

# Apply excel themes -----------------------------------------------------------
format_excel_table <- function(wb, sheet, cell) {

  # Style is applied to all columns.
  cols <- cell$columns[1]:cell$columns[2]
  body <- (cell$rows[1]:cell$rows[2]) + 1L

  # Style the title and merge columns
  rows <- cell$rows[1]
  openxlsx::addStyle(wb, sheet, excel_title, rows = rows, cols = cols, gridExpand = TRUE)
  openxlsx::mergeCells(wb, sheet, rows = rep(rows, length(cols)), cols = cols)

  # Apply baseline style to the data
  rows <- (cell$rows[1] + 1L):cell$rows[2]
  openxlsx::addStyle(wb, sheet, excel_body, rows = rows, cols = cols, gridExpand = TRUE)

  # Change styling for first (header)
  rows <- cell$rows[1] + 1L
  openxlsx::addStyle(wb, sheet, excel_header, rows = rows, cols = cols, gridExpand = TRUE)

  # Leftmost header should have left-aligned text
  openxlsx::addStyle(wb, sheet, excel_character, rows = rows, cols = cell$columns[1], gridExpand = TRUE, stack = TRUE)

  # ..  and last row
  rows <- cell$rows[2]
  openxlsx::addStyle(wb, sheet, excel_footer, rows = rows, cols = cols, gridExpand = TRUE)

}

excel_title <- openxlsx::createStyle(
  fontName = "Trebuchet MS",
  fontSize = 10,
  fontColour = "#000000",
  border = "Top",
  borderColour = "#0094A5",
  borderStyle = "medium",
  fgFill = "#7DC6CC",
  halign = "left",
  valign = "center",
  textDecoration = "Bold",
  wrapText = TRUE
)

excel_body <- openxlsx::createStyle(
  fontName = "Trebuchet MS",
  fontSize = 8,
  fontColour = "#000000",
  border = "Bottom",
  borderColour = "#BFBFBF",
  borderStyle = "thin",
  fgFill = "#FFFFFF",
  halign = "center",
  valign = "center"
)

excel_header <- openxlsx::createStyle(
  fontName = "Trebuchet MS",
  fontSize = 8,
  fontColour = "#000000",
  border = "Bottom",
  borderColour = "#0094A5",
  borderStyle = "thin",
  fgFill = "#CCE9EB",
  halign = "center",
  valign = "center",
  wrapText = TRUE
)

excel_footer <- openxlsx::createStyle(
  fontName = "Trebuchet MS",
  fontSize = 8,
  fontColour = "#000000",
  border = "Bottom",
  borderColour = "#0094A5",
  borderStyle = "medium",
  fgFill = "#FFFFFF",
  halign = "center",
  valign = "center"
)
