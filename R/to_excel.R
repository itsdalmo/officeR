#' Pass data to a excel workbook
#'
#' This function passes a \code{data.frame} to a excel workbook, and is stored
#' for later writing with \code{openxlsx}.
#'
#' @param df A \code{data.frame}.
#' @param wb A \code{Workbook} object where the data will be appended.
#' @param title The title to give to the table (only used if style = TRUE).
#' @param sheet Name of the sheet to write to, will be created if it does not exist.
#' @param row Optional: Also specify the startingrow for writing data.
#' @param format_style Set to FALSE if you do not want styling for the data.
#' @param format_values Set to FALSE and no formatting will be applied based on
#' variable type. With TRUE, character columns will be left justified, numeric
#' will have 1 decimal place, integer 0, and columns with values between 1 and 0
#' as percentages.
#' @param append Whether or not the function should append or clean the
#' sheet of existing data before writing.
#' @author Kristian D. Olsen
#' @note This function requires \code{openxlsx}.
#' @export
#' @examples
#' wb <- openxlsx::loadWorkbook("test.xlsx")
#' x %>% to_sheet(wb, sheet = "test", append = FALSE)
#' openxlsx::saveWorkbook(wb, "test.xlsx", overwrite = TRUE)

to_excel <- function(df, wb = NULL, title = "Table", sheet = "tables", row = 1L,
                     format_style = TRUE, format_values = TRUE, append = TRUE) {

  if (is.null(wb)) {
    wb <- Workbook$new()
  } else if (!inherits(wb, "Workbook")) {
    stop("wb argument must be a (loaded) openxlsx workbook")
  }

  if (!is.string(sheet)) {
    stop("The sheet has to be a string (character(1)).", call. = FALSE)
  }

  sheet_exists <- sheet %in% openxlsx::sheets(wb)

  # Get last row if sheet exists, or create if it does not.
  if (sheet_exists && isTRUE(append)) {

    row <- 2L + nrow(openxlsx::read.xlsx(wb, sheet = sheet, colNames = FALSE, skipEmptyRows = FALSE))
  } else if (sheet_exists) {
    openxlsx::removeWorksheet(wb, sheet)
    openxlsx::addWorksheet(wb, sheetName = sheet)
  } else {
    openxlsx::addWorksheet(wb, sheetName = sheet)
  }

  # Set table_row to be the last found row
  table_row <- row

  # Add data to the workbook
  if (is.null(names(df)) || identical(names(df), character(0))) {
    warning(sheet, ": No columnames in data. An empty sheet was created", call. = FALSE)
  } else {

    # When styling the title must be written first (and convert df names to titles)
    if (isTRUE(format_style)) {
      openxlsx::writeData(wb, sheet, title, startRow = row)
      names(df) <- stri_trans_totitle(names(df), type = "sentence")
      table_row <- row + 1
    }

    # Write the data.frame
    openxlsx::writeData(wb, sheet, df, startRow = table_row)
  }

  # Apply additional formatting if desired
  if (isTRUE(format_style) || isTRUE(format_values)) {

    format_xlsx(df, wb, sheet, table_row, style = format_style, values = format_values)

  }

  # Check the dimensions of what was written
  n_row <- dim(df)
  n_row <- if(is.null(n_row)) 0 else n_row[1]

  # Invisibly return the rows that have been written to
  invisible(setNames(c(table_row, table_row + n_row), c("first", "last")))

}