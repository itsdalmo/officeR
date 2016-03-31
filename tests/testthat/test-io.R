context("i/o")
test_that("Read/write single-sheet-xlsx" , {

  # Read
  xlsx <- read_data("xlsx.xlsx")
  expect_is(xlsx, "data.frame")
  expect_identical(xlsx$mainentity, paste("Test", 1:3))

  # Write
  fileName <- file.path(tempdir(), "xlsx.xlsx")
  write_data(xlsx, fileName)

  # Read data again
  w_xlsx <- read_data(fileName)
  expect_identical(w_xlsx, xlsx)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Writing tables to xlsx with to_excel" , {

  # Read
  xlsx <- read_data("xlsx.xlsx")

  # Create workbook
  fileName <- file.path(tempdir(), "xlsx.xlsx")
  wb <- openxlsx::createWorkbook()

  # Add table and save
  to_excel(xlsx, wb, title = "Table", sheet = "analysis")
  openxlsx::saveWorkbook(wb, fileName)

  # Read data again
  w_xlsx <- read_data(fileName)

  # Mutate to match
  names(w_xlsx) <- names(xlsx)
  w_xlsx <- w_xlsx[2:4, ] # Drop title row
  row.names(w_xlsx) <- 1:3

  expect_equal(w_xlsx, xlsx)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

# test_that("Tables with write_clipboard and read_clipboard for Windows/OSX" , {
#
#   if (on_windows() || on_osx()) {
#
#     # Read file and write it to clipboard
#     xlsx <- read_data("xlsx.xlsx")
#     write_clipboard(xlsx)
#
#     # Read clipboard and compare
#     cp_xlsx <- read_clipboard()
#     cp_xlsx$missing[1:2] <- NA # write_clipboard recodes NA to ""
#
#     # Compare
#     expect_equal(cp_xlsx, xlsx)
#
#   }
#
# })
#
# test_that("Text with write_clipboard and read_clipboard for Windows/OSX" , {
#
#   if (on_windows() || on_osx()) {
#
#     # Read file and write it to clipboard
#     txt <- "This is \n a test"
#     write_clipboard(txt)
#
#     # Read clipboard and compare
#     cp_txt <- read_clipboard()
#
#     # Compare
#     expect_identical(cp_txt, txt)
#
#   }
#
# })

test_that("Read and write_data with list", {

  # Read example data
  sheet1 <- read_data("xlsx.xlsx")
  sheet2 <- read_data("csv2.csv", encoding = "UTF-8", delim = ";")

  lst <- list("csv" = sheet1, "csv2" = sheet2)

  # Write csv
  dirName <- file.path(tempdir())
  expect_error(write_data(lst, file.path(dirName, "test.csv")))

  # Write xlsx
  fileName <- file.path(tempdir(), "xlsx.xlsx")
  write_data(lst, fileName)

  expect_true(file.exists(fileName))

  # Read data again and convert missing to numeric (openxlsx prob.)
  w_xlsx <- read_data(fileName)
  expect_equal(lst, w_xlsx)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Read and write_data for .csv files", {

  # Read
  xlsx <- read_data("xlsx.xlsx")
  csv <- read_data("csv.csv", delim = ",")
  csv2 <- read_data("csv2.csv", delim = ";")

  expect_equal(csv, csv2)
  expect_equal(csv2, xlsx)
  expect_identical(names(csv2), names(csv2))

  # Write
  fileName <- file.path(tempdir(), "csv2.csv")
  write_data(csv2, fileName)
  w_csv2 <- read_data(fileName)
  expect_identical(w_csv2, csv2)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Read and write_data for .txt files", {

  # Write
  fileName <- file.path(tempdir(), "txt.txt")

  csv2 <- read_data("csv2.csv", delim = ";")
  write_data(csv2, fileName)

  txt <- read_data(fileName, delim = "\t", encoding = "UTF-8")
  expect_identical(txt, csv2)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Read and write_data for .Rdata files", {

  # Write
  fileName <- file.path(tempdir(), "rdata.Rdata")

  csv2 <- read_data("csv2.csv", delim = ";")
  write_data(csv2, fileName)

  rdata <- read_data("rdata.Rdata")
  w_rdata <- read_data(fileName)

  expect_identical(w_rdata, csv2)
  expect_equal(w_rdata, rdata, check.attributes = FALSE) # readr "problems" from csv2.

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Read and write_data for .sav files", {

  sav <- haven::read_sav("sav.sav")
  xlsx <- read_data("xlsx.xlsx")

  sav$missing[1:2] <- NA # sav recodes NA to ""
  expect_identical(sav, xlsx)

  fileName <- file.path(tempdir(), "sav.sav")
  write_data(sav, fileName)

  # Check written data
  w_sav <- read_data(fileName)
  w_sav$missing[1:2] <- NA # Same as above
  expect_identical(sav, w_sav)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("i/o error handeling", {

  expect_error(read_data("invalid.xlsx"))
  expect_error(read_data("invalid.inv"))

})
