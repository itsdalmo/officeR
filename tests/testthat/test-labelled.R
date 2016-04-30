context("Labelled data")

org <- haven::read_sav("sav.sav")

test_that("Convert to labelled, i/o, and convert back" , {
  sav <- org
  fileName <- file.path(tempdir(), "sav.sav")

  # Create "mock" labelled variables and labels
  sav$mainentity <- factor(sav$mainentity)
  attr(sav, "labels") <- setNames(c("A" ,"B", "C", "D"), names(sav))

  # Convert to labelled and write
  out <- to_labelled(sav)
  write_data(out, fileName)

  # Read and check
  inp <- read_data(fileName)
  expect_equal(inp, out)

  # Convert from labelled
  res <- from_labelled(inp)
  expect_equal(res, sav)

  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Convert to_labelled without calling from_labelled first", {
  res <- to_labelled(org)
  expect_identical(as.character(lapply(res, attr, "label")), rep(NA_character_, 4))
})

test_that("data.table works with from/to_labelled", {
  skip_if_not_installed("data.table")
  res <- from_labelled(data.table::as.data.table(org))
  res <- to_labelled(res)
})
