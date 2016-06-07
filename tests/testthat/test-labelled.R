context("Labelled data")

org <- haven::read_sav("sav.sav")

test_that("Unlabelled values are retained when converting from labelled", {
  vals <- c("Agree", "Neutral", "Disagree", "Don't know")
  labelled <- haven::labelled(c(1, 5), labels = setNames(1:4, vals))

  expect_identical(
    levels(from_labelled(labelled)),
    c(vals, "5")
  )

})

test_that("Converting labelled character to factor conserves label order", {
  vals <- c("Agree", "Neutral", "Disagree", "Don't know")
  labelled <- haven::labelled(vals[c(2, 1, 3)], labels = setNames(vals, vals))

  # Output should be identical to haven
  expect_identical(
    as.character(from_labelled(labelled)),
    as.character(haven::as_factor(labelled))
  )

  # Levels should be identical to values
  expect_identical(
    levels(from_labelled(labelled)),
    vals
  )

  # Haven is not
  expect_false(
    identical(levels(haven::as_factor(labelled)), vals)
  )
})

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
  expect_identical(unlist(lapply(res, attr, "label")), NULL)
})

test_that("data.table works with from/to_labelled", {
  skip_if_not_installed("data.table")
  res <- from_labelled(data.table::as.data.table(org))
  expect_s3_class(res, "data.table")
  res <- to_labelled(res)
  expect_s3_class(res, "data.table")
})
