# Convert http links into a sharepoint link on windows -------------------------
as_network_drive <- function(link) {
  if (!on_windows()) {
    warning("This function only returns a path for network drives on Windows.")
  }
  is_url <- grepl("^https?://.*[^/]\\.[a-z]+/.*", link, ignore.case = TRUE)
  if (!is_url) {
    stop("Input was not a URL:\n", link, "\n(Make sure to include http://).")
  }

  # Extract the domain and folder
  domain <- sub("^https?://(.[^/]*)/.*", "\\1", link)
  folder <- sub(paste0(".*", domain, "(.*)"), "\\1", link)

  # Return
  paste0("\\\\", domain, "@SSL/DavWWWRoot", folder)

}

# Check which OS we are on -----------------------------------------------------
on_windows <- function() {
  unname(Sys.info()["sysname"]) == "Windows"
}

on_osx <- function() {
  unname(Sys.info()["sysname"]) == "Darwin"
}

# normalizes paths and removes trailing /'s ------------------------------------
clean_path <- function(path) {
  if (!is_string(path)) stop("Path must be a string.")

  # Normalize if path is not absolute.
  if (!grepl("^(/|[A-Za-z]:|\\\\|https?|~)", path)) {
    path <- normalizePath(path, winslash = "/", mustWork = FALSE)
  }

  # Remove trailing slashes
  if (grepl("/$", path)) {
    path <- sub("/$", "", path)
  }

  path

}

# Retrieve the filename sans extension -----------------------------------------
basename_sans_ext <- function(file)  {
  tools::file_path_sans_ext(basename(file))
}

# Check if a vector represents date or time ------------------------------------
is_date <- function(x) {
  inherits(x, c("POSIXct", "POSIXt", "Date"))
}

is_percent <- function(x) {
  is.numeric(x) && all(x <= 1L & x >= 0L)
}

# Check if x is a string (length 1 character vector.) --------------------------
is_string <- function(x) {
  is.character(x) && length(x) == 1
}

# See if a list or data.frame contains any labelled vectors. -------------------
any_labelled <- function(x) {
  any(vapply(x, haven::is.labelled, logical(1)))
}

# Check if an object is named (not NULL, "" or NA) -----------------------------
is_named <- function(x) {
  !is.null(names(x)) && !any(names(x) == "")
}

# Like is.list, except it does not return true for data.frame ------------------
is_list <- function(x) {
  inherits(x, "list")
}

# Hadley's %||% ----------------------------------------------------------------
`%||%` <- function(a, b) if (!is.null(a)) a else b
