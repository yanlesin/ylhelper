library(openxlsx)

#' @title Write Dataframe to XLSX with width set at "auto"
#'
#' @description Write Dataframe to XLSX with column width set to Auto
#' @param wb workbook
#' @param sheet sheet name or number
#' @param dataframe dataframe to write
#' @keywords XLSX columns width auto dataframe
#' @export
#' @examples
#' Write_XLSX_auto_width(wb, sheet, dataframe)

Write_XLSX_auto_width <- function(wb, sheet, dataframe){
  writeDataTable(wb,sheet, dataframe, keepNA = TRUE)
  setColWidths(wb, sheet, cols = 1:ncol(dataframe), widths = "auto")
}

#' @title Write Dataframe to XLSX with specified format to specified columns
#'
#' @description Write Dataframe to XLSX with specified format to specified columns
#' @param wb workbook
#' @param sheet sheet name or number
#' @param dataframe dataframe to write
#' @param style Excel style definition
#' @param cols Columns to apply format
#' @keywords XLSX columns dataframe format
#' @export
#' @examples
#' Write_XLSX_coll_format(wb, sheet, dataframe, style, cols)

Write_XLSX_coll_format <- function(wb, sheet, dataframe, style, cols){
  writeDataTable(wb,sheet, dataframe, keepNA = TRUE)
  addStyle(wb, sheet, style = style, rows = 2:(nrow(dataframe)+1), cols = cols, gridExpand = TRUE)
}

#' @title Apply format to dataframe columns
#'
#' @description Apply format to dataframe columns starting from second row up to end of dataframe
#' @param wb workbook
#' @param sheet sheet name or number
#' @param dataframe dataframe to write
#' @param style Excel style definition
#' @param cols Columns to apply format
#' @keywords XLSX column format
#' @export
#' @examples
#' coll_formats(wb, sheet, dataframe, style, cols)

coll_formats <- function(wb, sheet, dataframe, style, cols){
  addStyle(wb, sheet, style = style, rows = 2:(nrow(dataframe)+1), cols = cols, gridExpand = TRUE)
}

#' @title Set width of columns
#'
#' @description Set width of columns
#' @param wb workbook
#' @param sheet sheet name or number
#' @param cols Columns to apply format
#' @param width Width of column in points
#' @keywords XLSX column width
#' @export
#' @examples
#' coll_width(wb, sheet, cols, width)

coll_width <- function(wb, sheet, cols, width){
  setColWidths(wb, sheet, cols = cols, widths = width)
}
