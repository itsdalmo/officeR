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

  # CSI latent names
  latents =   c("image", "expect", "prodq", "servq", "value", "epsi", "loyal"),

  # The inner model used in PLS-PM. The 'plspm' package expects a matrix.
  model = rbind(image  = c(0,0,0,0,0,0,0),
                expect = c(1,0,0,0,0,0,0),
                prodq  = c(1,1,0,0,0,0,0),
                servq  = c(1,1,1,0,0,0,0),
                value  = c(0,0,1,1,0,0,0),
                epsi   = c(1,0,1,1,1,0,0),
                loyal  = c(0,0,0,0,0,1,0)),

  # Common names for variables associated with a given latent.
  associations = list(image  = "q4",
                      expect = "q5",
                      prodq  = "q7p",
                      servq  = "q7s",
                      value  = "q8",
                      epsi   = c("q3", "q6", "q16"),
                      loyal  = c("q10", "q15", "q15b")),

  # Translations used for generating reports.
  translation = list(
    required = c("image",
                 "expect",
                 "prodq",
                 "servq",
                 "value",
                 "epsi",
                 "loyal",
                 "mainentity",
                 "subentity",
                 "manifest",
                 "difference",
                 "question",
                 "contrast_average",
                 "average",
                 "study_average",
                 "spring",
                 "fall"),

    norwegian = stringi::stri_unescape_unicode(
      c("Image/inntrykk",
        "Forventninger",
        "Produktkvalitet",
        "Servicekvalitet",
        "Verdi for pengene",
        "Kundetilfredshet",
        "Lojalitet",
        "Bank",
        "Avdeling",
        "Kode",
        "Diff",
        "Sp\\u00f8rsm\\u00e5l",
        "Snitt nasjonal",
        "Snitt",
        "Snitt webstudien",
        "V\\u00e5r",
        "H\\u00f8st")
    ),

    danish = stringi::stri_unescape_unicode(
      c("Image",
        "Forventninger",
        "Produktkvalitet",
        "Servicekvalitet",
        "V\\u00e6rdi for pengene",
        "Kundetilfredshed",
        "Loyalitet",
        "Bank",
        "Afdeling",
        "Kode",
        "Diff",
        "Sp\\u00f8rgsm\\u00e5l",
        "Snit national",
        "Snit",
        "Snit webstudien",
        "For\\u00e5r",
        "Efter\\u00e5r")
    )),

  # Default config values for the .config private field in Survey's.
  config = list(
    required = c("name",
                 "segment",
                 "year",
                 "period",
                 "method",
                 "cutoff",
                 "model"),

    value = c("Example", NA, format(Sys.time(), "%Y"), NA, "web", NA, NA)),

  # Locations for package-internal files.
  template = list(
    beamer = list(
      dir = "rmd/beamer",
      files = "beamer_template.tex")),

  theme = list(
    beamer = list(
      dir = "rmd/beamer",
      files = c("beamercolorthememetropolis.sty",
                "beamerfontthememetropolis.sty",
                "beamerthemem.sty",
                "logo.eps",
                "beamer_preamble.tex"))),

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