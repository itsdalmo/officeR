# #' Evaluate a .Rmd file (after converting with rmd_to_r)
# #'
# #' Evaluate a \code{.Rmd} file in a given environment and return the results,
# #' including text with inline code (e.g. \code{`r 1+1`} becomes 2).
# #'
# #' @param rmd Path to a rmarkdown file.
# #' @param envir The environment in which to evaluate the rmarkdown file.
# #' @author Kristian D. Olsen
# #' @export
# #' @examples
# #' evaluate_rmd("test.Rmd", envir = parent.frame())
#
# evaluate_rmd <- function(rmd, envir = parent.frame()) {
#
#   pattern <- default$pattern$code
#
#   # Remove empty lines before evaluating
#   rmd <- rmd[rmd != ""]
#
#   # Get indexes
#   titles <- which(stri_detect(rmd, regex = pattern$title))
#   inlines <- which(stri_detect(rmd, regex = pattern$inline))
#   chunks <- which(!stri_detect(rmd, regex = pattern$text))
#   text <- setdiff(1:length(rmd), c(titles, inlines, chunks))
#
#   avoid_inline <- c(titles, chunks)
#   avoid_chunks <- c(inlines, titles, text)
#
#   results <- list()
#   n <- length(rmd)
#   index <- 0
#
#   for (i in seq_along(rmd)) {
#     if (i >= index) {
#
#       is_inline <- i %in% inlines
#       is_title <- i %in% titles
#       is_text <- i %in% text
#
#       if (is_inline || is_text && !is_title) {
#         index <- suppressWarnings(min(avoid_inline[avoid_inline > i], na.rm = TRUE))
#         if (is.infinite(index)) index <- n + 1
#
#         # Eval inline code and append to results
#         res <- eval_inline(rmd[i:(index-1)], envir = envir)
#         res <- stri_c(res[res != " "], collapse = "\n")
#         results <- c(results, as.list(res))
#
#       } else if (is_title) {
#         index <- i
#
#         # Eval titles and append to results
#         res <- eval_inline(rmd[i], envir = envir)
#         results <- c(results, res)
#
#       } else {
#         index <- suppressWarnings(min(avoid_chunks[avoid_chunks > i], na.rm = TRUE))
#         if (is.infinite(index)) index <- n + 1
#
#         # Eval chunks and append to results
#         res <- evaluate::evaluate(rmd[i:(index-1)], envir = envir, output_handler = ppt_handler)
#         results <- c(results, res)
#       }
#
#     }
#
#   }
#
#   # Return
#   results
#
# }

# # UTILITIES  -------------------------------------------------------------------
#
# eval_inline <- function(lines, envir = parent.frame()) {
#
#   pattern <- default$pattern$rmd
#
#   lines <- lapply(lines, function(x) {
#
#     is_inline <- stri_detect(x, regex = pattern$inline)
#     if (length(is_inline) == 0L || !is_inline) return(x) # Return early if it does not contain inline
#
#     inline <- unlist(stri_extract_all(x, regex = pattern$inline))
#     expr <- stri_replace_all(inline, "", regex = "`r\\s?|\\s?`")
#
#     for (i in seq_along(expr)) {
#       res <- as.character(eval(parse(text = expr[i]), envir = envir))
#       res <- if (!is.na(res) && length(res) > 0L) stri_c(res, collapse = " ") else " "
#       x <- stri_replace(x, replacement = res, fixed = inline[i])
#     }
#
#     x
#
#   })
#
#   # Remove ##+ and empty strings
#   lines <- stri_replace(unlist(lines), "", regex = "^##\\+ ")
#   lines <- lines[lines != ""]
#
#   lines
#
# }
#
# value_handler <- function(x, visible = TRUE) {
#
#   if (!visible) return()
#
#   if (inherits(x, "FlexTable") || is.data.frame(x)) {
#     x
#   } else {
#     print(x)
#   }
#
# }
#
# ppt_handler <- evaluate::new_output_handler(value = value_handler)
#
#
# #' Convert .Rmd to .R
# #'
# #' This function converts a \code{.Rmd} to a \code{.R} file (similar to \code{knitr::purl})
# #' by replacing chunk delimiters and putting \code{eval} chunk options inside \code{if}
# #' statements. Also comments out (\code{##+}) content that is not inside chunks.
# #'
# #' @param rmd Path to a rmarkdown file.
# #' @param encoding The encoding of both the input .Rmd file and the output.
# #' @author Kristian D. Olsen
# #' @return A .R file in the same directory and same name as the input .Rmd file.
# #' @note \code{UTF-8} is the recommended encoding for scripts and .Rmd files.
# #' @export
# #' @examples
# #' rmd_to_r("Example report.Rmd", encoding = "latin1")
#
# rmd_to_r <- function(rmd, encoding = "UTF-8", write = TRUE) {
#
#   # Assume strings are a path
#   if (is.string(rmd)) {
#     path <- clean_path(rmd)
#
#     if (!stri_trans_tolower(tools::file_ext(rmd)) == "rmd") {
#       stop("This function only accepts a .Rmd file", call. = FALSE)
#     } else {
#       rmd <- readLines(path, encoding = encoding)
#     }
#   } else {
#     write <- FALSE
#   }
#
#   # Get default patterns
#   pattern <- default$pattern$rmd
#
#   # Identify chunks
#   chunk_start <- which(stri_detect(rmd, regex = pattern$chunk_start))
#   chunk_end <- which(stri_detect(rmd, regex = pattern$chunk_end))
#
#   if (length(chunk_start) == length(chunk_end)) {
#     chunk_index <- Map(':', chunk_start, chunk_end)
#   } else {
#     stop("The .Rmd file contains unused chunk start/end indicators", call. = FALSE)
#   }
#
#   # Get and comment out all lines with content that is not in a chunk
#   not_chunk <- setdiff(1:length(rmd), unlist(chunk_index))
#   not_chunk <- setdiff(not_chunk, which(rmd == ""))
#
#   rmd[not_chunk] <- stri_c("##+ ", rmd[not_chunk])
#
#   # Comment out text, replace chunk delims and indicate chunknumber
#   n <- length(rmd)
#
#   for (i in seq_along(chunk_index)) {
#
#     # Update indices if the document has been extended
#     idx <- chunk_index[[i]]+(length(rmd)-n)
#     chunk <- replace_chunk_delim(rmd[idx])
#
#     # Piece together the chunk (w/delimiters) and the rest of the document
#     if (max(idx) < length(rmd)) {
#       last_line <- rmd[(max(idx)+1):length(rmd)]
#     } else if (chunk[length(chunk)] == "") {
#       last_line <- NULL
#     } else {
#       last_line <- ""
#     }
#
#     rmd <- c(if (min(idx) > 1) { rmd[1:(min(idx)-1)] }, chunk, if (!is.null(last_line)) { last_line })
#
#   }
#
#   # Save or return the cleaned Rmd file
#   if (write) {
#     path <- file(stri_c(tools::file_path_sans_ext(path), ".R"), encoding = "UTF-8")
#     on.exit(close(path), add = TRUE)
#     writeLines(rmd, path)
#   } else {
#     return(rmd)
#   }
#
#
# }
#
#
# # UTILITIES  -------------------------------------------------------------------
#
# replace_chunk_delim <- function(lines) {
#
#   # Get default pattern
#   pattern <- default$pattern$rmd
#
#   # Identify which (if any) chunks contain eval options
#   chunk_eval <- which(stri_detect(lines, regex = pattern$chunk_eval))
#   chunk_end <- which(stri_detect(lines, regex = pattern$chunk_end))
#
#   if (length(chunk_eval) > 0L) {
#
#     # Find the correct chunk end for each chunk eval
#     chunk_end <- vapply(chunk_eval, function(x) chunk_end[chunk_end > x][1], numeric(1))
#
#     # Replace chunk-delimiters with if-functions
#     lines[chunk_eval] <- stri_c("if (", stri_replace(lines[chunk_eval], "$1", regex = pattern$chunk_eval), ") {")
#     lines[chunk_end] <- "}"
#
#     # Double indendt the content in if-functions
#     chunk_index <- unlist(Map(':', chunk_eval, chunk_end))
#
#     content <- setdiff(chunk_index, c(chunk_eval, chunk_end))
#     lines[content] <- stri_c("  ", lines[content])
#
#   }
#
#   # Remove chunk start/endings without eval options
#   lines <- stri_replace_all(lines, "", regex = stri_c(pattern$chunk_start, ".*"))
#   lines <- stri_replace_all(lines, "", regex = stri_c(pattern$chunk_end, ".*"))
#
#   # Return
#   lines
#
# }
#
# extract_yaml <- function(rmd) {
#
#   # Patterns
#   pattern <- default$pattern$code
#
#   # Separate out YAML
#   yaml <- which(stri_detect(rmd, regex = pattern$yaml))
#   yaml <- yaml - c(-1, 1)
#   yaml <- rmd[yaml[1]:yaml[2]]
#
#   # Create report and add the first slide
#   nms <- vapply(yaml, stri_replace, replacement = "$1", regex = "^##\\+ ([[:alnum:]]*):.*", character(1))
#   nms <- unname(nms[!stri_detect(nms, regex = "^##\\+")])
#
#   yaml <- vapply(yaml, stri_replace, replacement = "$1", regex = "##\\+ [a-zA-Z]*:\\s*(.*)", character(1))
#   yaml <- unname(yaml[!stri_detect(yaml, regex = "^##\\+")])
#   yaml <- stri_replace_all(yaml, "", regex = "\"")
#
#   if (length(yaml) != length(nms)) {
#     stop("Problem parsing YAML frontmatter.", call. = FALSE)
#   }
#
#   yaml <- setNames(yaml, nms)
#   yaml <- as.list(yaml)
#
#   # Return
#   yaml
#
# }
