context("Utility functions")

test_that("any_fractions", {
  expect_true(any_fractions(c(1L, NA, 1.3)))
  expect_false(any_fractions(c(1L, NA, 11)))
})

test_that("join_str", {
  expect_error(join_str(1))
  expect_identical(join_str("A"), "A")
  expect_identical(join_str(c("A", "B")), "A and B")
})

test_that("trim_str", {
  x <- paste(letters[1:10], collapse = "")
  expect_error(trim_str(1))
  expect_identical(trim_str(x, n = 8), "abcde...")
  expect_identical(trim_str(x, n = 11, pad = " "), "abcdefghij ")
  expect_identical(trim_str(NA_character_), NA_character_)
})

test_that("clean scores", {

  expect_true(clean_score("1 aa") == "1")
  expect_identical(clean_score(c("1 aa", "bb 1", "10 cc")), c(1, NA, 10))

})

test_that("rescaling scores", {

  expect_true(clean_score("1 aa") == "1")
  expect_identical(clean_score(c("1 aa", "bb 1", "10 cc")), c(1, NA, 10))
  expect_identical(rescale_score(c("1", 10)), c(0, 100))

})

test_that("get_default", {

  expect_error(get_default(1L))
  expect_error(get_default("pa"), exact = TRUE) # Matches both palette and pattern
  expect_identical(get_default("pal"), internal_defaults$palette)
  expect_identical(get_default("pale"), internal_defaults$palette)

})

test_that("default_palette", {
  expect_identical(default_palette(), internal_defaults$palette)
})

test_that("default_latents", {
  expect_identical(default_latents(), internal_defaults$latents)
})

test_that("clean_path", {
  expect_error(clean_path(c("test", "file")))
})

test_that("filename_no_ext", {
  expect_identical(filename_no_ext("test.sav"), "test")
})


