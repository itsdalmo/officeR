context("Labelled data")

test_that("Convert to labelled, i/o, and convert back" , {
  sav <- haven::read_sav("sav.sav")
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

# TODO