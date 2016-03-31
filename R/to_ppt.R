#' Pass data to Powerpoint workbooks
#'
#' \code{to_ppt} allows you to pass R objects to an open \code{pptWorkbook},
#' and write it later with \code{write_data}. The \code{pptWorkbook} can be created
#' by calling \code{ppt_workbook}. \code{to_ppt} will always format the output
#' and add it to a new slide.
#'
#' You can use this function to pass \code{plot}, \code{data.frame} (\code{table} and
#' \code{matrix} will be coerced), and \code{character} objects to the workbook.
#' Strings are assumed to be markdown.
#'
#' @param x A \code{data.frame}, \code{string} or \code{plot}
#' @param wb A \code{pptWorkbook}.
#' @param title Title to use for the new slide.
#' @param subtitle Subtitle.
#' @param template Path to a \code{pptx} template. Default is the template provided
#' in this package.
#' @param font Default font (Replaces the \code{ReporteRs-default-font} option).
#' @param fontsize Default fontsize (Replaces the \code{ReporteRs-fontsize} option).
#' @author Kristian D. Olsen
#' @note This function requires \code{ReporteRs}. The \code{pptWorkbook} object
#' is a thin R6 wrapper around ReporteR's \code{pptx}, and allow us to use
#' \code{to_ppt} in chained expressions, since the workbook is mutable.
#' @export
#' @examples
#' if (require(ReporteRs)) {
#'  wb <- ppt_workbook()
#'  df <- data.frame("String" = c("A", "B"), "Int" = c(1:2L), "Percent" = c(0.5, 0.75))
#'
#'  # The workbook is mutable, so we don't have to assign result.
#'  to_ppt(df, wb, title = "Example data", subtitle = "")
#'
#'  # Data is first argument, so we can use it with dplyr.
#'  # df %>% to_ppt(wb, title = "Example data", subtitle = "")
#'
#'  # Save the data
#'  write_data(wb, "Example table.pptx")
#' }

to_ppt <- function(x, wb, title = NULL, subtitle = NULL) {
  if (!requireNamespace("ReporteRs")) {
    stop("This function requires 'ReporteRs'.")
  } else if (!inherits(wb, "pptWorkbook")) {
    stop ("'wb' should be a pptWorkbook. See help(to_ppt).")
  }
  UseMethod("to_ppt")
}

#' @rdname to_ppt
#' @export
ppt_workbook <- function(template = NULL, font = "Calibri", fontsize = 10L) {
  # Set ReporteRs options.
  options("ReporteRs-default-font" = font)
  options("ReporteRs-fontsize" = fontsize)

  pptWorkbook$new(template)
}

#' @export
write_data.pptWorkbook <- function(x, file, ...) {
  if (!requireNamespace("ReporteRs")) {
    stop("'ReporteRs' required to write pptWorkbook.")
  }
  ReporteRs::writeDoc(x$obj, file = file)
}

#' @rdname to_ppt
#' @export
to_ppt.data.frame <- function(x, wb, title = NULL, subtitle = NULL) {
  wb$add_table(format_flextable(x), title %||% " ", subtitle %||% " ")
}


#' @export
to_ppt.matrix <- function(x, wb, ...) {
  warning("Coercing ", class(x), " to data.frame.")
  to_ppt(as.data.frame(x, stringsAsFactors = FALSE), wb, ...)
}

#' @export
to_ppt.table <- to_ppt.matrix

#' @rdname to_ppt
#' @export
to_ppt.FlexTable <- function(x, wb, title = NULL, subtitle = NULL) {
  wb$add_table(x, title %||% " ", subtitle %||% " ")
}

#' @rdname to_ppt
#' @export
to_ppt.ggplot <- function(x, wb, title = NULL, subtitle = NULL) {
  wb$add_plot(x, title %||% " ", subtitle %||% " ")
}

# Plots returned by evaluate::evaluate().
#' @export
to_ppt.recodedplot <- to_ppt.ggplot

#' @rdname to_ppt
#' @export
to_ppt.character <- function(x, wb, title = NULL, subtitle = NULL) {
  stopifnot(is_string(x))
  wb$add_markdown(x, title %||% " ", subtitle %||% " ")
}

# Workbook for powerpoint (R6 Class) -------------------------------------------
# This exists because ReporteRs does not use a mutable object for documents,
# and I want to_ppt/to_excel to have identical interfaces.
#' @importFrom R6 R6Class
pptWorkbook <- R6::R6Class("pptWorkbook",
  public = list(
    obj = NULL,

    initialize = function(template = NULL) {
      if (!requireNamespace("ReporteRs")) {
        stop("'ReporteRs' required to create a pptWorkbook.")
      }
      template <- template %||% system.file("ppt", "template.pptx", package = "officeR")
      self$obj <- ReporteRs::pptx(template = clean_path(template))
    },

    add_table = function(x, title, subtitle) {
      self$obj <- ReporteRs::addSlide(self$obj, slide.layout = "Title and Content")
      self$obj <- ReporteRs::addTitle(self$obj, title)
      self$obj <- ReporteRs::addFlexTable(self$obj, x)
      self$obj <- ReporteRs::addParagraph(self$obj, subtitle)
      invisible(self)
    },

    add_plot = function(x, title, subtitle) {
      self$obj <- ReporteRs::addSlide(self$obj, slide.layout = 'Title and Content')
      self$obj <- ReporteRs::addTitle(self$obj, title)
      self$obj <- ReporteRs::addPlot(self$obj, fun = print, x = x, bg = "transparent")
      self$obj <- ReporteRs::addParagraph(self$obj, subtitle)
      invisible(self)
    },

    add_markdown = function(x, title, subtitle) {
      self$obj <- ReporteRs::addSlide(self$obj, slide.layout = 'Title and Content')
      self$obj <- ReporteRs::addTitle(self$obj, title)
      self$obj <- ReporteRs::addMarkdown(self$obj, text = x)
      self$obj <- ReporteRs::addParagraph(self$obj, subtitle)
      invisible(self)
    },

    print = function() {
      print(self$obj)
    }

  )
)

# Create a flextable with the correct theme ------------------------------------
format_flextable <- function(df) {
  # Create the flextable
  ft <- ReporteRs::FlexTable(
    data = df,
    header.columns = TRUE,
    add.rownames = FALSE,
    body.par.props = ReporteRs::parProperties(padding = 0L),
    body.cell.props = ppt_body_cell(),
    body.text.props = ppt_body_text(),
    header.par.props = ReporteRs::parProperties(padding = 0L),
    header.cell.props = ppt_header_cell(),
    header.text.props = ppt_header_text()
  )

  # Center numeric columns
  num_columns <- which(vapply(df, is.numeric, logical(1), USE.NAMES = FALSE))
  center <- ReporteRs::parProperties(text.align = "center")
  ft[, num_columns] <- center
  ft[, num_columns, to = "header"] <- center

  # Color border for the last row
  ft[nrow(df), ] <- ppt_last_row()

  # Set fixed withs (to avoid overdimensioned tables)
  # (Assumes fixed with for sheets)
  fw <- rep(9L/ncol(df), ncol(df))
  ft <- ReporteRs::setFlexTableWidths(ft, fw)

  # Return
  ft

}

# PPT theme --------------------------------------------------------------------
ppt_header_text <- function() {
  ReporteRs::textProperties(
    font.size = getOption("ReporteRs-fontsize"),
    font.family = getOption("ReporteRs-default-font"),
    font.weight = "bold"
  )
}

ppt_header_cell <- function() {
  ReporteRs::cellProperties(
    border.left.style = "none",
    border.right.style = "none",
    border.top.color = "#0094A5",
    border.top.style = "solid",
    border.top.width = 3L,
    border.bottom.color = "#0094A5",
    border.bottom.style = "solid",
    border.bottom.width = 1L,
    padding.top = 2L,
    padding.bottom = 1L,
    background.color = "#CCE9EB"
  )
}

# Body styles
ppt_body_text <- function() {
  ReporteRs::textProperties(
    font.size = getOption("ReporteRs-fontsize") - 2L,
    font.family = getOption("ReporteRs-default-font"),
    font.weight = "normal"
  )
}

ppt_body_cell <- function() {
  ReporteRs::cellProperties(
    border.left.style = "none",
    border.right.style = "none",
    border.top.style = "none",
    border.bottom.color = "#BFBFBF",
    border.bottom.style = "solid",
    border.bottom.width = 1L,
    padding.top = 1L,
    padding.bottom = 1L,
    background.color = "transparent"
  )
}

# Separate cell style for bottom row
ppt_last_row <- function() {
  ReporteRs::cellProperties(
    border.left.style = "none",
    border.right.style = "none",
    border.top.style = "none",
    border.bottom.color = "#0094A5",
    border.bottom.style = "solid",
    border.bottom.width = 2L,
    padding.top = 1L,
    padding.bottom = 1L,
    background.color = "transparent"
  )
}
