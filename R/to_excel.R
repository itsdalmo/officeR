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

to_excel <- function(df, wb, title = " ", format = !is.null(title),
                     sheet = "tables", row = 1L, append = TRUE) {

  if (!is.string(sheet)) {
    stop("The sheet has to be a string (character(1)).")
  } else if (!is.null(title) && !is.string(title)) {
    stop("Title has to be either NULL or a string.")
  } else if (!inherits(wb, "Workbook")) {
    stop("wb argument must be a (loaded) openxlsx workbook")
  }

  # Get last row if sheet exists, or create if it does not.
  if (sheet %in% openxlsx::sheets(wb)) {
    if (append) {
      row <- nrow(openxlsx::read.xlsx(wb, sheet = sheet, colNames = FALSE, skipEmptyRows = FALSE))
      row <- row + 2L # Add spacing between tables.
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

  # Write title and titlecase colnames when formatting.
  if (format) {
    openxlsx::writeData(wb, sheet, title, startRow = row)
    substr(names(df), 1L, 1L) <- toupper(substr(names(df), 1L, 1L))
    row <- row + 1L # Title has been written to the first row.
  }

  # Write the rest of the data before formatting.
  openxlsx::writeData(wb, sheet, df, startRow = row)

  # Apply additional formatting if desired
  if (format) {
    coltypes <- vapply(df, function(x) class(x)[1], character(1))
    format_xlsx(wb, sheet, coltypes, row, dim(df))
  }

  # Check the dimensions of what was written
  last_row <- nrow(df) %||% 0L

  # TODO: Why did I want this again?
  invisible(setNames(c(row, row + last_row), c("first", "last")))

}