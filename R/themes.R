# #' ggplot2 theme
# #'
# #' Wrapper to make styling plots for a report less verbose.
# #'
# #' @param legend Default is bottom. Passed directly to theme call for ggplot2.
# #' @author Kristian D. Olsen
# #' @return A ggplot2 theme object
# #' @import ggplot2
# #' @export
# #' @examples
# #' p <- ggplot(df, aes(x=manifest, y=score, group=mainentity)) + geom_bar()
# #' p + plot_theme()
#
# plot_theme <- function(legend="bottom"){
#
#   theme(
#
#     # Make the background of the entire plot transparent
#     panel.background = element_rect(fill = "transparent", colour = NA),
#     plot.background = element_rect(fill = "transparent", colour = NA),
#     legend.background = element_rect(fill = "transparent", colour = NA),
#     legend.key = element_rect(fill = "transparent", colour = NA),
#
#     # Margins should also be transparent
#     panel.border = element_rect(fill = "transparent", colour=NA),
#
#     # Fix the border and axis lines/ticks
#     axis.ticks = element_blank(),
#
#     # Format axis text size
#     axis.text.x = element_text(size=9,colour="#23373b"),
#     axis.text.y = element_text(size=9,colour="#23373b"),
#     axis.title.y = element_blank(),
#     axis.title.x = element_blank(),
#
#     # Format the grid (only horizontal lines)
#     panel.grid.major.x = element_blank(),
#     panel.grid.major.y = element_line(colour="#D0D0D0",size=.5),
#
#     # Facet-title formatting
#     strip.text = element_text(colour="white"),
#     strip.background = element_rect(fill="#23373b"),
#
#     # Change plot margins
#     panel.margin = grid::unit(2, "lines"), # facet margins
#     plot.margin = grid::unit(c(0, 0, 0, 0), "cm"),
#
#     # Legend position and format
#     legend.position = legend,
#     legend.title = element_blank()
#   )
#
# }
#
# #' @rdname plot_theme
# #' @export
# plot_shared_legend <- function(...) {
#
#   plots <- list(...)
#
#   grobs <- ggplotGrob(plots[[1]] + theme(legend.position = "bottom"))$grobs
#   legends <- grobs[[which(lapply(grobs, function(x) x$name) == "guide-box")]]
#   heights <- sum(legends$height)
#
#   gridExtra::grid.arrange(
#     do.call(gridExtra::arrangeGrob, lapply(plots, function(x) {
#       x + theme(legend.position="none")})),
#     legends,
#     ncol = 1,
#     heights = grid::unit.c(grid::unit(1, "npc") - heights, heights))
#
# }
#
#
# # ReporteRs flextable theme
# format_flextable <- function(df, fontsize = 9) {
#
#   # Specify the properties for the table
#   tp <- ReporteRs::textProperties(font.size = fontsize)
#   bp <- ReporteRs::borderProperties(style="none")
#
#   # Create flextable with proper fontsizes
#   ft <- ReporteRs::FlexTable(df, body.text.props = tp, header.text.props = tp)
#
#   # Set columnwidths (assumes question text is found in the second column)
#   ft <- ReporteRs::setFlexTableWidths(ft, widths=c(.5, 5, 1.3, 1.3, .5))
#
#   # Add borderproperties
#   ft <- ReporteRs::setFlexTableBorders(ft, inner.vertical=bp, inner.horizontal=bp,
#                                        outer.vertical=bp, outer.horizontal=bp)
#
#   # Center text in columns with scores
#   ft[, 3:ncol(df)] <-  ReporteRs::parProperties(text.align = "center")
#   ft[, 3:ncol(df), to = "header"] <-  ReporteRs::parProperties(text.align = "center")
#
#   # Zebrastyle the rows
#   ft <- ReporteRs::setZebraStyle(ft, odd = "#DDDDDD", even = "#FFFFFF")
#
#   # Return the flextable
#   return(ft)
#
# }
#
# # Function used to format style and values in xlsx wrappers --------------------
#
# format_xlsx <- function(df, wb, sheet, table_row, style = TRUE, values = TRUE) {
#
#   # Get row and column index
#   title_row <- table_row - 1
#   last_row <- table_row + nrow(df)
#   all_rows <- table_row:last_row
#   all_cols <- 1:ncol(df)
#
#   # Optional styling of the table
#   if (isTRUE(style)) {
#     # Add title, style the cells and merge them
#     openxlsx::addStyle(wb, sheet, xlsx_style_title, rows = title_row, cols = all_cols, gridExpand = TRUE)
#     openxlsx::mergeCells(wb, sheet, rows = rep(title_row, length(all_cols)), cols = all_cols)
#
#     # Style the entire table
#     openxlsx::addStyle(wb, sheet, xlsx_style_table, rows = all_rows, cols = all_cols, gridExpand = TRUE)
#
#     # Alter the style of first row
#     openxlsx::addStyle(wb, sheet, xlsx_style_firstrow, rows = table_row, cols = all_cols, gridExpand = TRUE)
#
#     # ..  and last row
#     openxlsx::addStyle(wb, sheet, xlsx_style_lastrow, rows = last_row, cols = all_cols, gridExpand = TRUE)
#   }
#
#   # Optional formatting for numeric values
#   if (isTRUE(values)) {
#
#     # Get column indicies
#     fct_columns <- which(vapply(df, is.factor, logical(1)))
#     chr_columns <- which(vapply(df, is.character, logical(1)))
#     int_columns <- which(vapply(df, is.integer, logical(1)))
#     pct_columns <- which(vapply(df, function(x) all(is.numeric(x) && x <= 1 && x >= 0), logical(1)))
#     num_columns <- which(vapply(df, is.numeric, logical(1)))
#
#     num_columns <- setdiff(num_columns, c(int_columns, pct_columns))
#
#     # Format numeric to 1 decimal place
#     if (length(num_columns)) {
#       openxlsx::addStyle(wb, sheet, xlsx_numeric, rows = all_rows, cols = num_columns,
#                          gridExpand = TRUE, stack = TRUE)
#     }
#
#     # Format integer to 0 decimal places
#     if (length(int_columns)) {
#       openxlsx::addStyle(wb, sheet, xlsx_integer, rows = all_rows, cols = int_columns,
#                          gridExpand = TRUE, stack = TRUE)
#     }
#
#     # Format percentages
#     if (length(pct_columns)) {
#       openxlsx::addStyle(wb, sheet, xlsx_percent, rows = all_rows, cols = pct_columns,
#                          gridExpand = TRUE, stack = TRUE)
#     }
#
#     # Change the text orientation of character columns
#     if (length(chr_columns) || length(fct_columns)) {
#       openxlsx::addStyle(wb, sheet, xlsx_character, rows = all_rows, cols = c(chr_columns, fct_columns),
#                          gridExpand = TRUE, stack = TRUE)
#     }
#
#   }
#
# }
#
#
# # openxlsx styles --------------------------------------------------------------
# xlsx_numeric <- openxlsx::createStyle(numFmt = "0.0")
# xlsx_integer <- openxlsx::createStyle(numFmt = "0")
# xlsx_percent <- openxlsx::createStyle(numFmt = "0.0%")
# xlsx_character <- openxlsx::createStyle(halign = "left")
#
# xlsx_style_title <- openxlsx::createStyle(
#   fontName = "Trebuchet MS",
#   fontSize = 10,
#   fontColour = "#000000",
#   border = "Top",
#   borderColour = "#0094A5",
#   borderStyle = "medium",
#   fgFill = "#7DC6CC",
#   halign = "left",
#   valign = "center",
#   textDecoration = "Bold",
#   wrapText = TRUE
# )
#
# xlsx_style_table <- openxlsx::createStyle(
#   fontName = "Trebuchet MS",
#   fontSize = 8,
#   fontColour = "#000000",
#   border = "Bottom",
#   borderColour = "#BFBFBF",
#   borderStyle = "thin",
#   fgFill = "#FFFFFF",
#   halign = "center",
#   valign = "center"
# )
#
# xlsx_style_firstrow <- openxlsx::createStyle(
#   fontName = "Trebuchet MS",
#   fontSize = 8,
#   fontColour = "#000000",
#   border = "Bottom",
#   borderColour = "#0094A5",
#   borderStyle = "thin",
#   fgFill = "#CCE9EB",
#   halign = "center",
#   valign = "center"
# )
#
# xlsx_style_lastrow <- openxlsx::createStyle(
#   fontName = "Trebuchet MS",
#   fontSize = 8,
#   fontColour = "#000000",
#   border = "Bottom",
#   borderColour = "#0094A5",
#   borderStyle = "medium",
#   fgFill = "#CCE9EB",
#   halign = "center",
#   valign = "center"
# )
#
# xlsx_style_lastrow <- openxlsx::createStyle(
#   fontName = "Trebuchet MS",
#   fontSize = 8,
#   fontColour = "#000000",
#   border = "Bottom",
#   borderColour = "#0094A5",
#   borderStyle = "medium",
#   fgFill = "#FFFFFF",
#   halign = "center",
#   valign = "center"
# )
