#' Get default values used internally.
#'
#' This function retrieves internally defined values where the names match the
#' \code{string} argument (not case sensitive). \code{default_latents} and
#' \code{default_palette} are identical to \code{get_default("latents")}.
#'
#' @param string A string which matches the default values you would like to return.
#' @param exact If the \code{string} matches more than one default value and exact
#' is set to \code{TRUE} an error occurs. The error lists the full name of all
#' matching internal defaults. Set this to \code{FALSE} to instead return all
#' matches instead.
#' @author Kristian D. Olsen
#' @export
#' @examples
#' lats <- get_default("latents")
#' pal <- get_default("palette")
#'
#' identical(lats, default_latents())
#' identical(pal, default_palette())

get_default <- function(string, exact = TRUE) {
  if (!is.string(string)) {
    stop("Expecting a string (character(1)) input for argument 'string'.")
  }

  id <- stri_detect(names(internal_defaults), fixed = string, ignore_case = TRUE)
  id <- names(internal_defaults)[id]
  if (length(id) > 1L && exact) {
    stop("The string matched more than one element:\n", join_str(stri_c("'", id, "'")))
  }

  res <- internal_defaults[id]
  if (length(res) == 1L)
    res <- res[[1]]

  res

}

#' @rdname get_default
#' @export
default_palette <- function() get_default("palette", exact = TRUE)

#' @rdname get_default
#' @export
default_latents <- function() get_default("latents", exact = TRUE)

# Default values ---------------------------------------------------------------

internal_defaults <- list(

  # Default palette
  palette = c("#F8766D", "#00BFC4", "#808080", "#00BF7D", "#9590FF", "#A3A500", "#EA8331"),

  # List (nested) of regex patterns used internally.
  pattern = list(
    detect_scale = "^[0-9]{1,2}[[:alpha:][:punct:] ]*",
    extract_scale = "^[0-9]{1,2}\\s*=?\\s*([[:alpha:]]*)",

    rmd = list(
      chunk_start = "^```\\{r",
      chunk_end = "```$",
      chunk_eval = ".*eval\\s*=\\s*((.[^},]+|.[^}]+\\))),?.*",
      inline = "`r[ [:alnum:][:punct:]][^`]+`",
      section = "^#[^#]",
      slide = "^##[^#]"),

    code = list(
      yaml = "^##\\+ ---",
      inline = "`r[ [:alnum:][:punct:]][^`]+`",
      title = "^##\\+\\s*#{1,2}[^#]",
      text = "^##\\+\\s.*"))

)