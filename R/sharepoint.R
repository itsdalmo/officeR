#' Prepare a sharepoint link for reading
#'
#' This function is minimalistic wrapper for \code{httr}, with the intention
#' of making it easy to read files from sharepoint using \code{read_data}.
#'
#' @param link A string starting with \code{http}.
#' @param mounted Set this to \code{TRUE} if you are on windows and have mounted
#' the sharepoint server as a network drive.
#' @param file An object returned from \code{sharepoint_link}.
#' @param destination Optional: When sharepoint is not mounted, files are always
#' downloaded before being read into R. The file is temporary if you do not set
#' the destination.
#' @param ... Further arguments passed to \code{read_data}.
#' @author Kristian D. Olsen
#' @note This function requires \code{httr} if the drive is not mounted. User/password
#' will be requested and stored for the current R session. Imporant that you don't
#' share your .Rhistory after this!
#' @export
#' @examples
#' \dontrun{
#' # The gist of it:
#' read_data(sharepoint_link(your_server))
#' }

sharepoint_link <- function(link, mounted = FALSE) {
  if (!is_string(link)) stop("'link' must be a string (character(1).")

  # Use mounted drives if they exist.
  if (on_windows() && mounted) {
    drive <- as_network_drive(link)
    if (file.exists(drive)) {
      link <- structure(drive, class = c("sharepoint_mount", class(link)))
    }
  } else {
    if (!requireNamespace("httr")) {
      stop("'httr' is required for reading from sharepoint.")
    }
    link <- structure(URLencode(link), class = c("sharepoint", class(link)))
  }
  link
}

#' @rdname sharepoint_link
#' @export
read_data.sharepoint <- function(file, destination = NULL, ...) {
  if (!requireNamespace("httr")) {
    stop("'httr' is required for reading from sharepoint.")
  }

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

  # Check SP version
  sp10 <- httr::modify_url(file, path = "_vti_bin/ListData.svc/")
  sp10 <- httr::GET(sp10, httr::authenticate(usr, pwd, type = "ntlm"))

  if (httr::status_code(sp10) == 200) {
    if (tools::file_ext(file) == "") {
      stop("Cannot read directories. Sharepoint 2010 does not support REST.")
    }
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
  read_data(file, ...)
}