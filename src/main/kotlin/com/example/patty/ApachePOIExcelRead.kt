package com.example.patty

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.util.ResourceUtils
import java.io.File
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.io.IOException


class ApachePOIExcelRead {

    private var descriptionIndex: Int = -1;
    private var firstHeader: Row? = null
    private var secondHeader: Row? = null
    private val dishes: MutableMap<String, MutableList<Dish>> = mutableMapOf()
    private val file = ResourceUtils.getFile("classpath:01112021_normales.xls")
    fun read() {
        try {
            val workbook: Workbook = WorkbookFactory.create(file)
            val datatypeSheet: Sheet = workbook.getSheetAt(0)
            val iterator = datatypeSheet.iterator().withIndex()
            while (iterator.hasNext()) {
                val currentRowVal = iterator.next()
                val currentRow: Row = currentRowVal.value
                val cellIterator: Iterator<IndexedValue<Cell>> = currentRow.iterator().withIndex()
                when (currentRowVal.index) {
                    0 -> firstHeader = currentRow
                    1 -> {
                        secondHeader = currentRow;
                        getDescriptionIndex(cellIterator)
                    }
                    else -> if (descriptionIndex != -1) break;
                }
                if (currentRowVal.index == 1) {
                    getDescriptionIndex(cellIterator)
                }
            }
            identifyDifferentDishes(
                datatypeSheet.iterator().asSequence().toList()
                    .subList(2, datatypeSheet.iterator().asSequence().toList().size)
            )
            writeToFile()
        } catch (e: FileNotFoundException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }

    private fun getDescriptionIndex(cellIterator: Iterator<IndexedValue<Cell>>) {
        while (cellIterator.hasNext()) {
            val currentCellVal: IndexedValue<Cell> = cellIterator.next()
            val currentCell: Cell = currentCellVal.value
            if (currentCell.cellType === CellType.STRING) {
                println(currentCell.stringCellValue.toString())
                if (currentCell.stringCellValue.toString() == "Descripcion") {
                    descriptionIndex = currentCellVal.index
                    break
                }
            }
        }
    }

    private fun identifyDifferentDishes(rows: List<Row>) {
        rows.forEach {
            val dishValues = it.iterator().asSequence().toList()
            if (dishValues.size > 1) {
                val dishDescription = dishValues[descriptionIndex].stringCellValue.toString()
                val descriptionFirstWord = dishDescription.split(" ")[0]
                val key = if (descriptionFirstWord.last().isDigit()) descriptionFirstWord.substring(
                    0,
                    descriptionFirstWord.length - 1
                ) else descriptionFirstWord
                val dish = mapRowToDish(it)
                dishes.merge(key, mutableListOf(dish), ::addListToAnother)
            }
        }
    }

    private fun addListToAnother(list: MutableList<Dish>, list2: MutableList<Dish>): MutableList<Dish> {
        list.addAll(list2)
        return list
    }

    private fun writeToFile() {
        val workbook = XSSFWorkbook()
        val sortedDishes = dishes.toSortedMap()
        sortedDishes.forEach { mapDishElement ->
            val sheet = workbook.createSheet(mapDishElement.key)
            addSheetHeaders(sheet, mapDishElement)
            var rowIndex = 1
            mapDishElement.value.sortBy { it.description }
            var temDishName = ""
            mapDishElement.value.forEach { dish ->
                val rowToWrite = sheet.createRow(++rowIndex)
                temDishName = if (temDishName == "" || temDishName == dish.description) {
                    createRow(rowToWrite, dish)
                    dish.description
                } else {
                    createEmptyRow(rowToWrite)
                    val rowToWrite2 = sheet.createRow(++rowIndex)
                    createRow(rowToWrite2, dish)
                    ""
                }
            }
            createPartDishSheet(mapDishElement, workbook, "PROTEINA", "GUARNICION")
            createPartDishSheet(mapDishElement, workbook, "ARROZ", "SOPA")
            createPartDishSheet(mapDishElement, workbook, "ENSALADA")
            createPartDishSheet(mapDishElement, workbook, "BEBIDA", "POSTRE")
        }
        val out = FileOutputStream(File("01112021.xlsx"))
        workbook.write(out)
        out.close()
        println("gfgcontribute.xlsx written successfully on disk.")
    }

    private fun createPartDishSheet(
        mapDishElement: Map.Entry<String, MutableList<Dish>>,
        workbook: XSSFWorkbook,
        vararg options: String
    ) {
        if (mapDishElement.key == "ALMUERZO" || mapDishElement.key == "MERIENDA") {
            val optionsJoinedWithDash = options.joinToString("-")
            val dishesFilteredByOptionsInGroups =
                mapDishElement.value.filter { dishElement -> options.any { dishElement.itemOp.contains(it) } }
                    .groupBy { Pair(it.dishDesc, it.itemOp) }
                    .toSortedMap(compareBy<Pair<String, String>> { it.first }.thenBy { it.second })
            val sheet = workbook.createSheet("${mapDishElement.key} $optionsJoinedWithDash")
            addSheetHeaders(sheet, mapDishElement, optionsJoinedWithDash)
            var rowIndex2 = 1
            dishesFilteredByOptionsInGroups.values.forEach { dishList ->
                dishList.forEachIndexed { index, dish ->
                    val rowToWrite = sheet.createRow(++rowIndex2)
                    createRow(rowToWrite, dish)
                    dish.dishDesc
                    if (index == (dishList.size - 1)) {
                        val totalWeight = dishList.sumOf { it.weight }
                        val totalNoPa = dishList.sumOf { it.noPa }
                        val totalRow = sheet.createRow(++rowIndex2)
                        createTotalRow(totalRow, totalWeight, totalNoPa)
                    }
                }
            }
        }
    }

    private fun addSheetHeaders(
        sheet: XSSFSheet,
        mapDishElement: Map.Entry<String, MutableList<Dish>>,
        customTitleForFirstHeader: String = ""
    ) {
        val firstRow = sheet.createRow(0)
        val firstRowCell = firstRow.createCell(0)
        firstRowCell.setCellValue(
            "${
                firstHeader!!.iterator().asSequence().toList()[0].stringCellValue
            } $customTitleForFirstHeader ${mapDishElement.key}"
        )
        val secondRow = sheet.createRow(1)
        val headers = secondHeader!!.iterator().asSequence().toList()
        headers.forEachIndexed { index, cell ->
            val newCell = secondRow.createCell(index)
            newCell.setCellValue(cell.stringCellValue)
        }
    }

    private fun createRow(rowToWrite: XSSFRow, dish: Dish) {
        val serviceDateCell = rowToWrite.createCell(0)
        serviceDateCell.setCellValue(dish.serviceDate)
        val itemCell = rowToWrite.createCell(1)
        itemCell.setCellValue(dish.item)
        val descriptionCell = rowToWrite.createCell(2)
        descriptionCell.setCellValue(dish.description)
        val itemOpCell = rowToWrite.createCell(3)
        itemOpCell.setCellValue(dish.itemOp)
        val dishDescCell = rowToWrite.createCell(4)
        dishDescCell.setCellValue(dish.dishDesc)
        val noPaCell = rowToWrite.createCell(5)
        noPaCell.setCellValue(dish.noPa.toDouble())
        val weightCell = rowToWrite.createCell(6)
        weightCell.setCellValue(dish.weight)
    }

    private fun createEmptyRow(rowToWrite: XSSFRow) {
        val emptyCell = rowToWrite.createCell(0)
        emptyCell.setCellValue("")
    }

    private fun createTotalRow(rowToWrite: XSSFRow, weight: Double, noPa: Int) {
        val totalCell = rowToWrite.createCell(4)
        totalCell.setCellValue("TOTALES")
        val totalNoPaCell = rowToWrite.createCell(5)
        totalNoPaCell.setCellValue(noPa.toDouble())
        val totalWeightCell = rowToWrite.createCell(6)
        totalWeightCell.setCellValue(weight)

    }

    private fun mapRowToDish(row: Row): Dish {
        val rowList = row.iterator().asSequence().toList()
        return Dish(
            rowList[0].stringCellValue,
            rowList[1].stringCellValue,
            rowList[2].stringCellValue,
            rowList[3].stringCellValue,
            rowList[4].stringCellValue,
            rowList[5].numericCellValue.toInt(),
            rowList[6].numericCellValue
        )
    }
}