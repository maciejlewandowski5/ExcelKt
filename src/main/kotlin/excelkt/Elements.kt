package excelkt

import org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK
import org.apache.poi.xssf.usermodel.*
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.util.*

@DslMarker
annotation class ExcelElementMarker

@ExcelElementMarker
abstract class ExcelElement {
    abstract val xssfWorkbook: XSSFWorkbook
    fun createCellStyle(f: XSSFCellStyle.() -> Unit = { }): XSSFCellStyle = xssfWorkbook.createCellStyle().apply(f)
    fun createFont(f: XSSFFont.() -> Unit = { }): XSSFFont = xssfWorkbook.createFont().apply(f)
}

class Workbook(
        override val xssfWorkbook: XSSFWorkbook,
        private val style: XSSFCellStyle?
) : ExcelElement() {
    fun sheet(name: String? = null, style: XSSFCellStyle? = null, init: Sheet.() -> Unit) =
            Sheet(
                    xssfWorkbook = xssfWorkbook,
                    style = style ?: this.style,
                    name = name
            ).apply(init)
}

class Sheet(
        override val xssfWorkbook: XSSFWorkbook,
        private val style: XSSFCellStyle?,
        name: String?
) : ExcelElement() {
    val xssfSheet = name?.let(xssfWorkbook::createSheet) ?: xssfWorkbook.createSheet()
    private var currentRowIndex = 0

    fun row(style: XSSFCellStyle? = null, initi: (Row.() -> Unit)? = null) {
        return Row(
                xssfWorkbook = xssfWorkbook,
                style = style ?: this.style,
                xssfSheet = xssfSheet,
                index = currentRowIndex++
        ).let { row ->
            initi?.also {
                row.apply(it)
            }
        }

    }
}

class Row(
        override val xssfWorkbook: XSSFWorkbook,
        private val style: XSSFCellStyle?,
        xssfSheet: XSSFSheet,
        index: Int
) : ExcelElement() {
    val xssfRow = xssfSheet.createRow(index)
    private var currentCellIndex = 0

    fun cell(content: Any? = null, style: XSSFCellStyle? = null) {
        Cell(
                xssfWorkbook = xssfWorkbook,
                style = style ?: this.style,
                content = content,
                xssfRow = xssfRow,
                index = currentCellIndex++
        )
    }
}

class Cell(
        override val xssfWorkbook: XSSFWorkbook,
        private val style: XSSFCellStyle?,
        content: Any?,
        xssfRow: XSSFRow,
        index: Int
) : ExcelElement() {
    init {

        xssfRow.createCell(index).run {
            content?.let {
                when (it) {
                    is Formula -> setCellFormula(it.content)
                    is Boolean -> setCellValue(it)
                    is Number -> setCellValue(it.toDouble())
                    is Date -> setCellValue(it)
                    is Calendar -> setCellValue(it)
                    is LocalDate -> setCellValue(Date.from(it.atStartOfDay(ZoneId.systemDefault()).toInstant()))
                    is LocalDateTime -> setCellValue(Date.from(it.atZone(ZoneId.systemDefault()).toInstant()))
                    else -> setCellValue(it.toString())
                }
            } ?: this.setCellType(CELL_TYPE_BLANK)

            this@Cell.style?.let {
                cellStyle = it
            }
        }
    }
}

data class Formula(val content: String)
