#' Prepare a sharepoint link for reading
#'
#' This function is minimalistic wrapper for \code{httr}, with the intention
#' of making it easy to read files from sharepoint using \code{read_data}.
#'
#' @param link A string starting with \code{http}.
#' @param mounted Set this to \code{TRUE} if you are on windows and have mounted
#' the sharepoint server as a network drive.
#' @param x An object returned from \code{sharepoint_link}.
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
  if (!is.string(link)) stop("'link' must be a string (character(1).")

  # If we are on windows, check if the drive is mounted.
  if (on_windows() && mounted) {
    drive <- as_network_drive(link)
    if (file.exists(drive)) {
      link <- structure(drive, class = c("mounted_sharepoint", class(link)))
    }
  } else {
    if (!requireNamespace("httr")) {
      stop("'httr' is required for reading from sharepoint.")
    }
    # URL encode the link if not.
    link <- structure(URLencode(link), class = c("sharepoint", class(link)))
  }
  link
}

#' @rdname sharepoint_link
read_data.sharepoint <- function(x, destination = NULL, ...) {
  if (!requireNamespace("httr")) {
    stop("'httr' is required for reading from sharepoint.")
  }

  # Always download the file before reading.
  if (is.null(destination)) {
    destination <- tempfile(fileext = paste0(".", tools::file_ext(x)))
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

  # Encode URL string and GET
  resp <- httr::GET(x, httr::authenticate(usr, pwd, type = "ntlm"), httr::write_disk(destination))
  code <- httr::status_code(test)

  if (code == 200L) {
    # Success. Return the data.
    read_data(destination, ...)
  } else if (code == 401L) {
    stop("Unauthorized (401): Wrong username and/or password (or authentication).")
  } else if (code == 400L) {
    stop("Bad request (400): Something wrong with the function.")
  } else if (code == 404L) {
    stop("Not found (404): File does not exist.")
  } else {
    stop("Something went wrong. Status code: ", code)
  }

}

#' @rdname sharepoint_link
read_data.mounted_sharepoint <- function(x, ...) {
  read_data(x, ...)
}