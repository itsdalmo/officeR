# Convert http links into a sharepoint link on windows -------------------------
as_network_drive <- function(link) {
  if (!on_windows()) {
    warning("This function only returns a path for network drives on Windows.")
  }
  is_url <- grep("^https?://.*[^/]\\.[a-z]+/.*", link, ignore.case = TRUE)
  if (!is_url) {
    stop("Input was not a URL:\n", link, "\n(Make sure to include http://).")
  }

  # Extract the domain and folder
  domain <- sub("^https?://(.[^/]*)/.*", "\\1", link)
  folder <- sub(paste0(".*", domain, "(.*)"), "\\1", link)

  # Return
  paste0("\\\\", domain, "@SSL/DavWWWRoot", folder)

}

# Turn a character vector into a listing
# E.g. Turn c("A", "B") into "A and B" -----------------------------------------
str_list <- function(x, conjunction = "and", quote = "'") {
  stopifnot(is.character(x))
  if (!is.null(quote)) {
    # Pasting quotes because sQuote and dQuote depend on the useFancyQuotes option.
    x <- paste0(quote, x, quote)
  }
  if (length(x) == 1L) {
    return(x)
  }
  paste(paste0(x[1:(length(x)-1)], collapse = ", "), conjunction, x[length(x)])
}

# Check which OS we are on -----------------------------------------------------
on_windows <- function() {
  unname(Sys.info()["sysname"]) == "Windows"
}

on_osx <- function() {
  unname(Sys.info()["sysname"]) == "Darwin"
}

# match_all returns ALL matching indices for x in table, ordered by x ----------
match_all <- function(x, table) {
  unlist(lapply(x, function(x) which(table == x)))
}

# normalizes paths and removes trailing /'s ------------------------------------
clean_path <- function(path) {
  if (!is.string(path)) stop("Path must be a string.")

  # Normalize if path is not absolute.
  if (!grepl("^(/|[A-Za-z]:|\\\\|~)", path)) {
    path <- normalizePath(path, winslash = "/", mustWork = FALSE)
  }

  # Remove trailing slashes
  if (grepl("/$", path)) {
    path <- sub("/$", "", path)
  }

  path

}

# Retrieve the filename sans extension -----------------------------------------
filename_no_ext <- function(file)  {
  sub(paste0("(.*)\\.", tools::file_ext(file), "$"), "\\1", basename(file))
}

# Check if a vector represents date or time ------------------------------------
is_date <- function(x) {
  inherits(x, c("POSIXct", "POSIXt", "Date"))
}

is_percent <- function(x) {
  is.numeric(x) && all(x <= 1L & x >= 0L)
}

# Check if x is a string (length 1 character vector.) --------------------------
is.string <- function(x) {
  is.character(x) && length(x) == 1
}

# See if a list or data.frame contains any labelled vectors. -------------------
any_labelled <- function(x) {
  any(vapply(x, inherits, what = "labelled", logical(1)))
}

# Like is.list, except it does not return true for data.frame ------------------
is.list2 <- function(x) {
  inherits(x, "list")
}

# Hadley's %||% ----------------------------------------------------------------
`%||%` <- function(a, b) if (!is.null(a)) a else b

