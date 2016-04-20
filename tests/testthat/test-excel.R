context("Excel workbooks")
test_that("Create excel workbook" , {

  ex <- excel_workbook()
  expect_is(ex, "Workbook")
  expect_identical(attr(class(ex), "package"), "openxlsx")

})

test_that("Add data to and write excel workbook" , {

  ex <- excel_workbook()
  df <- data.frame(A = c("a", "b"), B = c(1, 2), stringsAsFactors = FALSE)
  to_excel(df, ex)

  # Write
  fileName <- file.path(tempdir(), "xlsx.xlsx")
  write_data(ex, fileName, overwrite = TRUE)

  # Read
  re <- read_data(fileName, sheet = "tables", skip = 1L) # Skip title
  class(re) <- "data.frame"
  expect_identical(df, re)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Add data to an existing excel workbook" , {

  ex <- excel_workbook()
  df <- data.frame(A = c("a", "b"), B = c(1, 2), stringsAsFactors = FALSE)
  to_excel(df, ex, title = NULL)

  # Write
  fileName <- file.path(tempdir(), "xlsx.xlsx")
  write_data(ex, fileName, overwrite = TRUE)

  # Load
  ex2 <- expect_error(excel_workbook("test.xls"), "'template' should be NULL or a path")
  ex2 <- excel_workbook(fileName)
  to_excel(df, ex2, title = NULL, sheet = "tables 2")
  write_data(ex2, fileName, overwrite = TRUE)

  # Read
  re <- read_data(fileName)
  expect_true(is.list(re))
  expect_identical(names(re), c("tables", "tables 2"))
  expect_identical(re[["tables"]], re[["tables 2"]])

  unlink(fileName, recursive = TRUE, force = TRUE)

})