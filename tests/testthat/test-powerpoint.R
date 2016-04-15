context("Powerpoint workbooks")

test_that("Create powerpoint workbook" , {
  skip_if_not_installed("ReporteRs")
  ppt <- ppt_workbook()
  expect_s3_class(ppt, c("pptWorkbook", "R6"))
})

test_that("warning 'deprecated' when using addMarkdown" , {
  skip_if_not_installed("ReporteRs")
  ppt <- ppt_workbook()
  ppt$obj <- ReporteRs::addSlide(ppt$obj, slide.layout = 'Title and Content')
  expect_warning(ReporteRs::addMarkdown(ppt$obj, text = "test"), "deprecated")
})

test_that("Add data to and write Powerpoint workbook" , {
  skip_if_not_installed("ReporteRs")
  ppt <- ppt_workbook()
  df <- data.frame(A = c("a", "b"), B = c(1, 2), stringsAsFactors = FALSE)

  to_ppt(df, ppt)

  fileName <- file.path(tempdir(), "ppt.pptx")
  write_data(ppt, fileName, overwrite = TRUE)

  expect_true(file.exists(fileName))
  unlink(fileName, recursive = TRUE, force = TRUE)

})

test_that("Add markdown to Powerpoint" , {
  skip_if_not_installed("ReporteRs")
  ppt <- ppt_workbook()
  rmd <- c("*This is added as markdown.*")

  to_ppt(rmd, ppt)

  fileName <- file.path(tempdir(), "ppt.pptx")
  write_data(ppt, fileName, overwrite = TRUE)

  expect_true(file.exists(fileName))
  unlink(fileName, recursive = TRUE, force = TRUE)

})