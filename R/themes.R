# Function used to format style and values in xlsx wrappers --------------------

format_xlsx <- function(wb, sheet, types, row, dims) {

  # Style the title and merge columns
  rows <- row - 1L; cols <- 1:dims[1]
  openxlsx::addStyle(wb, sheet, xlsx_style_title, rows = rows, cols = cols, gridExpand = TRUE)
  openxlsx::mergeCells(wb, sheet, rows = rep(rows, length(cols)), cols = cols)

  # Apply baseline style to the data
  rows <- row:(row + dims[2]); cols <- 1:dims[1]
  openxlsx::addStyle(wb, sheet, xlsx_style_table, rows = rows, cols = cols, gridExpand = TRUE)

  # Change styling for first (header)
  rows <- row; cols <- 1:dims[1]
  openxlsx::addStyle(wb, sheet, xlsx_style_firstrow, rows = rows, cols = cols, gridExpand = TRUE)

  # ..  and last row
  rows <- row + dims[2]; cols <- 1:dims[1]
  openxlsx::addStyle(wb, sheet, xlsx_style_lastrow, rows = rows, cols = cols, gridExpand = TRUE)


  # Get column indicies
  fct_columns <- which(vapply(df, is.factor, logical(1)))
  chr_columns <- which(vapply(df, is.character, logical(1)))
  int_columns <- which(vapply(df, is.integer, logical(1)))
  pct_columns <- which(vapply(df, function(x) all(is.numeric(x) && x <= 1 && x >= 0), logical(1)))
  num_columns <- which(vapply(df, is.numeric, logical(1)))

  num_columns <- setdiff(num_columns, c(int_columns, pct_columns))

  # Format numeric to 1 decimal place
  if (length(num_columns)) {
    openxlsx::addStyle(wb, sheet, xlsx_numeric, rows = all_rows, cols = num_columns,
                       gridExpand = TRUE, stack = TRUE)
  }

  # Format integer to 0 decimal places
  if (length(int_columns)) {
    openxlsx::addStyle(wb, sheet, xlsx_integer, rows = all_rows, cols = int_columns,
                       gridExpand = TRUE, stack = TRUE)
  }

  # Format percentages
  if (length(pct_columns)) {
    openxlsx::addStyle(wb, sheet, xlsx_percent, rows = all_rows, cols = pct_columns,
                       gridExpand = TRUE, stack = TRUE)
  }

  # Change the text orientation of character columns
  if (length(chr_columns) || length(fct_columns)) {
    openxlsx::addStyle(wb, sheet, xlsx_character, rows = all_rows, cols = c(chr_columns, fct_columns),
                       gridExpand = TRUE, stack = TRUE)
  }


}


# openxlsx styles --------------------------------------------------------------
xlsx_numeric <- openxlsx::createStyle(numFmt = "0.0")
xlsx_integer <- openxlsx::createStyle(numFmt = "0")
xlsx_percent <- openxlsx::createStyle(numFmt = "0.0%")
xlsx_character <- openxlsx::createStyle(halign = "left")

xlsx_style_title <- openxlsx::createStyle(
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

xlsx_style_table <- openxlsx::createStyle(
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

xlsx_style_firstrow <- openxlsx::createStyle(
  fontName = "Trebuchet MS",
  fontSize = 8,
  fontColour = "#000000",
  border = "Bottom",
  borderColour = "#0094A5",
  borderStyle = "thin",
  fgFill = "#CCE9EB",
  halign = "center",
  valign = "center"
)

xlsx_style_lastrow <- openxlsx::createStyle(
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
