context("Utilities")

lnk <- "http://www.such-test.com/directory/file.txt/"

test_that("Convert to network drive" , {
  if (on_windows()) {
    expect_error(expect_warning(as_network_drive("test")))
    out <- as_network_drive(lnk)
  } else {
    expect_error(as_network_drive("test"))
    expect_warning(out <- as_network_drive(lnk), "on Windows.")
  }
  out <- clean_path(out)
  expect_identical(out, "\\\\www.such-test.com@SSL/DavWWWRoot/directory/file.txt")
})

test_that("Check OS's" , {
  on_windows()
  on_osx()
})

test_that("Clean path on non-string fails" , {
  expect_error(clean_path(c("A", "B")))
})