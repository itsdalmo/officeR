#' Pass data to a excel workbook
#'
#' This function passes a \code{data.frame} to a openxlsx workbook.
#'
#' @param df A \code{data.frame}.
#' @param wb A \code{Workbook}.
#' @param title The title to give to the table. Must be \code{NULL} if you do not
#' want the table to be styled.
#' @param sheet Name of the sheet to use when writing the data.
#' @param format Format values and apply the default template to the table output.
#' @param append Whether or not the function should append or clean the
#' sheet of existing data before writing.
#' @param row Optional: Specify the startingrow when writing data to a new sheet.
#' @author Kristian D. Olsen
#' @note This function requires \code{openxlsx}.
#' @export
#' @examples
#' if (require(openxlsx)) {
#'  wb <- openxlsx::createWorkbook()
#'  df <- data.frame("String" = c("A", "B"), "Int" = c(1:2L), "Percent" = c(0.5, 0.75))
#'
#'  # The workbook is mutable, so we don't have to assign result.
#'  to_sheet(df, wb, title = "Example data", sheet = "Example", append = FALSE)
#'
#'  # Data is first argument, so we can use it with dplyr.
#'  # df %>% to_sheet(wb, title = "Example dplyr", sheet = "Example", append = TRUE)
#'
#'  # Save the data
#'  write_data(wb, "Example tables.xlsx", overwrite = TRUE)
#' }

to_excel <- function(df, wb, title = " ", sheet = "tables", format = !is.null(title), append = TRUE, row = 1L) {

  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Package 'openxlsx' required to write .xlsx files.")
  } else if (!is.string(sheet)) {
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
    openxlsx::writeData(wb, sheet, title %||% " ", startRow = row)
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

