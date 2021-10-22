package com.example.patty

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication

@SpringBootApplication
class PattyApplication

fun main(args: Array<String>) {
    runApplication<PattyApplication>(*args)
    val reader = ApachePOIExcelRead()
    reader.read()
}
