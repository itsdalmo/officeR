context("Utility functions")

test_that("join_str", {
  expect_error(join_str(1))
  expect_identical(join_str("A"), "A")
  expect_identical(join_str(c("A", "B")), "A and B")
})

test_that("get_default", {

  expect_error(get_default(1L))
  expect_error(get_default("pa"), exact = TRUE) # Matches both palette and pattern
  expect_identical(get_default("pal"), internal_defaults$palette)
  expect_identical(get_default("pale"), internal_defaults$palette)

})

test_that("clean_path", {
  expect_error(clean_path(c("test", "file")))
})

test_that("filename_no_ext", {
  expect_identical(filename_no_ext("test.sav"), "test")
})


