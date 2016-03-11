#' Powerpoint template
#'
#' This function uses \code{ReporteRs} and \code{evaluate} to generate powerpoint
#' slides based on the same rmarkdown syntax that creates the pdf reports. I.e.
#' it is a dumbed down version of normal rmarkdown, where only \code{#} and \code{##}
#' are understood as being section titles and new frames/subtitles respectively.
#'
#' @param entity Name of the entity.
#' @param dir Base directory (from .Rmd template)
#' @param envir The environment in which to evaluate the rmarkdown code.
#' @author Kristian D. Olsen
#' @note This function is used by \code{generate_report}.

generate_ppt <- function(entity, dir, envir) {
  
  # Read in the .Rmd file
  rmd <- file.path(dir, "Markdown", stri_c(entity, ".Rmd"))
  rmd <- readLines(rmd, encoding = "UTF-8")
  
  # Convert the rmd to "r"
  if (!any(stri_detect(rmd, regex = "##\\+"), na.rm = TRUE)) {
    rmd <- rmd_to_r(rmd, write = FALSE)
  }
  
  # Separate out YAML
  is_yaml <- which(stri_detect(rmd, regex = default$pattern$code$yaml))
  if (!length(is_yaml)) {
    stop("Could not find the YAML frontmatter.", call. = FALSE)
  } else {
    yaml <- extract_yaml(rmd)
    rmd <- rmd[-c(is_yaml[1]:is_yaml[2])]
  }
  
  # Evaluate the markdown
  res <- evaluate_rmd(rmd, envir = envir)
  
  # Create report and add the first slide
  doc <- ReporteRs::pptx(template = file.path(dir, "PPT", "ppt_template.pptx"))
  doc <- ReporteRs::addSlide(doc, slide.layout = "Title Slide")
  doc <- ReporteRs::addTitle(doc, value = yaml$title)
  doc <- ReporteRs::addParagraph(doc, value = yaml$subtitle)
  doc <- ReporteRs::addParagraph(doc, value = yaml$author)
  doc <- ReporteRs::addParagraph(doc, value = yaml$date)
  
  # Add title etc
  doc <- to_ppt(doc, res)
  
  ReporteRs::writeDoc(doc, file = file.path(dir, "PPT", stri_c(entity, ".pptx")))
  
  invisible()
  
}

to_ppt <- function(doc, res) {
  
  # Get the types of data in our results
  type <- result_types(res)
  
  # Get the actual results
  title <- " "; subtitle <- " "
  
  for (i in seq_along(res)) {
    if (type[i] == "title") {
      title <- stri_replace(res[[i]], "", regex = "^#[^#]")
    } else if (type[i] == "subtitle") {
      subtitle <- stri_replace(res[[i]], "", regex = "^##[^#]")
    } else if (type[i] == "table") {
      doc <- ReporteRs::addSlide(doc, slide.layout = 'Title and Content')
      doc <- ReporteRs::addTitle(doc, title)
      doc <- ReporteRs::addFlexTable(doc, format_flextable(res[[i]]))
      doc <- ReporteRs::addParagraph(doc, subtitle)
    } else if (type[i] == "flextable") {
      doc <- ReporteRs::addSlide(doc, slide.layout = 'Title and Content')
      doc <- ReporteRs::addTitle(doc, title)
      doc <- ReporteRs::addFlexTable(doc, res[[i]])
      doc <- ReporteRs::addParagraph(doc, subtitle)
    } else if (type[i] == "markdown") {
      doc <- ReporteRs::addSlide(doc, slide.layout = 'Title and Content')
      doc <- ReporteRs::addTitle(doc, title)
      doc <- ReporteRs::addMarkdown(doc, text = res[[i]])
      doc <- ReporteRs::addParagraph(doc, subtitle)
    } else if (type[i] == "recordedplot") {
      doc <- ReporteRs::addSlide(doc, slide.layout = 'Title and Content')
      doc <- ReporteRs::addTitle(doc, title)
      doc <- ReporteRs::addPlot(doc, fun = print, x = res[[i]], bg = "transparent")
      doc <- ReporteRs::addParagraph(doc, subtitle)
    } else if (type[i] == "footnote") {
      warning("Footnotes are not supported and will not be included.", call. = FALSE)
    } else if (type[i] == "latextable") {
      warning("Latex tables are not supported and will not be included.", call. = FALSE)
    } else if (type[i] == "warning") {
      # evaluate::replay(res[[i]])
    } else if (type[i] == "error") {
      evaluate::replay(res[[i]])
    }
    
  }

  return(doc)
  
}

result_types <- function(res) {
  
  vapply(res, function(x) {
    
    if (length(x) == 0L || is.na(x)) return("unknown")
    
    if (is.character(x)) {
      if (stri_detect(x, regex = "^#[^#].*")) {
        "title" 
      } else if (stri_detect(x, regex = "^##[^#].*")) {
        "subtitle"
      } else if (stri_detect(x, regex = "begin\\{table\\}")) {
        "latextable"
      } else if (stri_detect(x, regex = "footnoteextra\\{")) {
        "footnote"
      } else {
        "markdown"
      }
    } else if (inherits(x, "recordedplot")) {
      "recordedplot"
    } else if (inherits(x, "source")) {
      "source"
    } else if (inherits(x, "data.frame")) {
      "table"
    } else if (inherits(x, "FlexTable")) {
      "flextable"
    } else if (inherits(x, "error")) {
      "error"
    } else if (inherits(x, "warning")) {
      "warning"
    } else {
      "unknown"
    }}, character(1))
  
}

copy_ppt_theme <- function(dir) {
  
  files <- system.file("ppt/ppt_template.pptx", package="reporttool")
  
  lapply(files, function(x, dir) { 
    file.copy(x, file.path(dir, basename(x)), overwrite=FALSE) 
  }, dir)
  
  invisible()
  
}
