context("sharepoint")

lnk <- "http://www.such-test.com/directory/file.txt"

test_that("sharepoint_link returns correct class" , {
  skip_if_not_installed("httr")
  sp <- sharepoint_link(lnk)

  expect_error(sharepoint_link(c("a", "b")), "must be a string")
  expect_s3_class(sp, "sharepoint_link")

})

test_that("We can call read_data on sharepoint links" , {
  # Slow test due to timeout.
  skip_if_not_installed("httr")
  sp <- sharepoint_link(lnk)

  # Set user/password to test
  httr:::set_envvar("sp_usr", "test", "session")
  httr:::set_envvar("sp_pwd", "test", "session")

  expect_error(read_data(sp), "Couldn't resolve")
})

test_that("sharepoint mount errors" , {
  expect_error(sp <- sharepoint_mount("A", "B"))
  expect_error(sp <- sharepoint_mount(lnk))
})

test_that("read_data on sharepoint mounts" , {
  skip_if_not(on_windows())
  sp <- sharepoint_mount(lnk)
  expect_error(read_data(sp), "Path does not exist")
})
