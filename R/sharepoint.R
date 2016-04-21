#' Prepare a sharepoint link for reading
#'
#' This function is minimalistic wrapper for \pkg{httr}, with the intention
#' of making it easy to read files from sharepoint using \code{\link{read_data}}.
#'
#' @param link A string starting with \code{http}.
#' @param file An object returned from \code{sharepoint_link}.
#' @param destination Optional: When sharepoint is not mounted, files are always
#' downloaded before being read into R. The file is temporary if you do not set
#' the destination.
#' @param ... Further arguments passed to \code{\link{read_data}}.
#' @author Kristian D. Olsen
#' @note This function requires \pkg{httr} if the drive is not mounted. User/password
#' will be requested and stored for the current R session. Imporant that you don't
#' share your .Rhistory after this!
#' @export
#' @examples
#' \dontrun{
#' # The gist of it:
#' read_data(sharepoint_link(your_server))
#' read_data(sharepoint_mount(your_server))
#' }

sharepoint_mount <- function(link) {
  if (!is_string(link))
    stop("'link' must be a string.")
  if (!on_windows())
    stop("Can only create paths to mounts on Windows.")
  structure(as_network_drive(link), class = c("sharepoint_mount", "character"))
}

#' @rdname sharepoint_mount
#' @export
sharepoint_link <- function(link) {
  if (!is_string(link))
    stop("'link' must be a string.")
  if (!requireNamespace("httr"))
    stop("'httr' is required to read sharepoint links.")
  structure(URLencode(link), class = c("sharepoint_link", "character"))
}

#' @rdname sharepoint_mount
#' @export
read_data.sharepoint_link <- function(file, destination = NULL, ...) {
  if (!requireNamespace("httr"))
    stop("'httr' is required to read sharepoint links.")
  if (tolower(tools::file_ext(file)) == "")
    stop("Cannot read directories. Sharepoint 2010 does not support REST.")

  # Always download the file before reading.
  if (is.null(destination)) {
    destination <- tempfile(fileext = paste0(".", tools::file_ext(file)))
    on.exit(unlink(destination, recursive = TRUE, force = TRUE), add = TRUE)
  }

  # Check if credentials are stored in session
  usr <- httr:::get_envvar("sp_usr")
  pwd <- httr:::get_envvar("sp_pwd")

  if (is.na(usr) || is.na(pwd)) {
    # Request user input
    usr <- readline("Enter username: ")
    pwd <- readline("Enter username: ")
    httr:::set_envvar("sp_usr", usr, "session")
    httr:::set_envvar("sp_pwd", pwd, "session")
  }

  # GET and check status
  resp <- httr::GET(file, httr::authenticate(usr, pwd, type = "ntlm"), httr::write_disk(destination))
  httr::stop_for_status(resp)

  # Read data
  read_data(destination, ...)

}

#' @rdname sharepoint_link
#' @export
read_data.sharepoint_mount <- function(file, ...) {
  NextMethod()
}