package com.github.thleu.kexcelapi

import org.hamcrest.Matchers.*
import org.junit.Assert.assertThat
import org.junit.Rule
import org.junit.Test
import org.junit.experimental.runners.Enclosed
import org.junit.experimental.theories.DataPoints
import org.junit.experimental.theories.Theories
import org.junit.experimental.theories.Theory
import org.junit.rules.ExpectedException
import org.junit.rules.TemporaryFolder
import org.junit.runner.RunWith
import java.io.FileNotFoundException
import java.io.IOException
import java.util.*

@RunWith(Enclosed::class)
class KExcelTest {
    class Normal_open {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        @Throws(Exception::class)
        fun openFileNameTest() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            val wb = KExcel.open(file.canonicalPath)
            assertThat(wb, `is`(notNullValue()))
            wb.close()
        }
    }

    class Abnormal_open {
        @Test(expected = FileNotFoundException::class)
        fun openFileNameTest() {
            KExcel.open("/dummy")
        }
    }

    class Abnormal_open_can_not_write_to_file {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test(expected = IOException::class)
        fun openFileNameTest() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            file.setReadOnly()
            val wb = KExcel.open(file.canonicalPath)
            KExcel.write(wb, file.canonicalPath)
        }
    }

    class Obtaining_workbook_with_sheet_index {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        @Throws(Exception::class)
        fun openFileNameTest() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            val wb = KExcel.open(file.canonicalPath)
            assertThat(wb[0], `is`(not(nullValue())))
            wb.close();
        }
    }

    class Obtaining_workbook_with_sheet_name {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        @Throws(Exception::class)
        fun openFileNameTest() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            val wb = KExcel.open(file.canonicalPath)
            assertThat(wb["sheet1"], `is`(not(nullValue())))
            wb.close();
        }
    }

    @RunWith(Theories::class)
    class Normal_cellIndexToCellLabel {
        data class Fixture(val x: Int, val y: Int, val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            assertThat(KExcel.cellIndexToCellLabel(fixture.x, fixture.y), `is`(fixture.cellLabel))
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture(0, 0, "A1"),
                    Fixture(1, 0, "B1"),
                    Fixture(2, 0, "C1"),
                    Fixture(26, 0, "AA1"),
                    Fixture(27, 0, "AB1"),
                    Fixture(28, 0, "AC1")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_reading_cell_with_label {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toStr(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("B2", "あいうえお"),
                    Fixture("C3", "123"),
                    Fixture("D4", "150.51"),
                    Fixture("C2", "123")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_reading_cell_with_index {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val x: Int, val y: Int, val expected: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.x, fixture.y].toStr(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture(1, 1, "あいうえお"),
                    Fixture(2, 2, "123"),
                    Fixture(3, 3, "150.51"),
                    Fixture(2, 1, "123")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_toStr {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toStr(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("B2", "あいうえお"),
                    Fixture("C2", "123"),
                    Fixture("D2", "150.51"),
                    Fixture("F2", "true"),
                    Fixture("G2", "123150.51"),
                    Fixture("H2", ""),
                    Fixture("I2", ""),
                    Fixture("J2", "あいうえお123")
            )
        }
    }

    @RunWith(Theories::class)
    class Abnormal_toStr {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        data class Fixture(val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(UnsupportedOperationException::class.java)
                sheet[fixture.cellLabel].toStr()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("E2")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_toInt {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: Int)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toInt(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("B3", 456),
                    Fixture("C3", 123),
                    Fixture("D3", 150),
                    Fixture("G3", 369),
                    Fixture("J3", 456123)
            )
        }
    }

    @RunWith(Theories::class)
    class Abnormal_toInt {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        data class Fixture(val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toInt()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("B2"),
                    Fixture("E3"),
                    Fixture("F3"),
                    Fixture("H3"),
                    Fixture("I3"),
                    Fixture("K3")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_toDouble {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: Double)

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toDouble(), `is`(closeTo(fixture.expected, 0.00001)))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("B4", 123.456),
                    Fixture("C4", 123.0),
                    Fixture("D4", 150.51),
                    Fixture("G4", 50.17),
                    Fixture("J4", 123123.456)
            )
        }
    }

    @RunWith(Theories::class)
    class Abnormal_toDouble {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        data class Fixture(val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toDouble()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("B2"),
                    Fixture("E4"),
                    Fixture("F4"),
                    Fixture("H4"),
                    Fixture("I4"),
                    Fixture("K4")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_toBoolean {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: Boolean)

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toBoolean(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("F5", true),
                    Fixture("G5", false)
            )
        }
    }

    @RunWith(Theories::class)
    class Abnormal_toBoolean {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        data class Fixture(val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toBoolean()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("B5"),
                    Fixture("C5"),
                    Fixture("D5"),
                    Fixture("E5"),
                    Fixture("K5")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_toDate {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: Date)

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toDate(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("E6", TestUtil.getDate(2015, 12, 1)),
                    Fixture("G6", TestUtil.getDate(2015, 12, 3))
            )
        }
    }

    @RunWith(Theories::class)
    class Abnormal_toDate {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        data class Fixture(val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toDate()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("A6"),
                    Fixture("B6"),
                    Fixture("C6"),
                    Fixture("D6"),
                    Fixture("F6"),
                    Fixture("H6"),
                    Fixture("I6"),
                    Fixture("J6"),
                    Fixture("K6")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_toTime {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: Date)

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toDate(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("E7", TestUtil.getTime(10, 10, 30)),
                    Fixture("G7", TestUtil.getTime(12, 34, 30))
            )
        }
    }

    @RunWith(Theories::class)
    class Abnormal_toTime {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        data class Fixture(val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toDate()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("A7"),
                    Fixture("B7"),
                    Fixture("C7"),
                    Fixture("D7"),
                    Fixture("F7"),
                    Fixture("H7"),
                    Fixture("I7"),
                    Fixture("J7"),
                    Fixture("K7")
            )
        }
    }

    @RunWith(Theories::class)
    class Normal_toDateTime {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        data class Fixture(val cellLabel: String, val expected: Date)

        @Theory
        @Throws(Exception::class)
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet[fixture.cellLabel].toDate(), `is`(fixture.expected))
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("E8", TestUtil.getDateTime(2015, 12, 1, 10, 10, 30)),
                    Fixture("G8", TestUtil.getDateTime(2015, 12, 3, 10, 10, 30))
            )
        }
    }

    @RunWith(Theories::class)
    class Abnormal_toDateTime {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        data class Fixture(val cellLabel: String)

        @Theory
        fun test(fixture: Fixture) {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalAccessException::class.java)
                sheet[fixture.cellLabel].toDate()
            }
        }

        companion object {
            @DataPoints
            @JvmField
            val PARAMs = arrayOf(
                    Fixture("A8"),
                    Fixture("B8"),
                    Fixture("C8"),
                    Fixture("D8"),
                    Fixture("F8"),
                    Fixture("H8"),
                    Fixture("I8"),
                    Fixture("J8"),
                    Fixture("K8")
            )
        }
    }

    class Normal_set_String {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        fun test() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]
                sheet["A1"] = "あいうえお"

                KExcel.write(workbook, file.canonicalPath)
            }
            // 書き込んだものを再読込する
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet["A1"].toStr(), `is`("あいうえお"))
            }
        }
    }

    class Normal_set_String_index {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        fun test() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]
                sheet[0, 0] = "あいうえお"
                sheet[0, 1] = "かきくけこ"
                sheet[1, 2] = "さしすせそ"

                KExcel.write(workbook, file.canonicalPath)
            }
            // 書き込んだものを再読込する
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet["A1"].toStr(), `is`("あいうえお"))
                assertThat(sheet["A2"].toStr(), `is`("かきくけこ"))
                assertThat(sheet["B3"].toStr(), `is`("さしすせそ"))
            }
        }
    }

    class Normal_set_Int {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        fun test() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]
                sheet["A1"] = 12345

                KExcel.write(workbook, file.canonicalPath)
            }
            // 書き込んだものを再読込する
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet["A1"].toInt(), `is`(12345))
            }
        }
    }

    class Normal_set_Double {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        fun test() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]
                sheet["A1"] = 150.51

                KExcel.write(workbook, file.canonicalPath)
            }
            // 書き込んだものを再読込する
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet["A1"].toDouble(), `is`(closeTo(150.51, 0.00001)))
            }
        }
    }

    class Normal_set_Boolean {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        fun test() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]
                sheet["A1"] = true

                KExcel.write(workbook, file.canonicalPath)
            }
            // 書き込んだものを再読込する
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet["A1"].toBoolean(), `is`(true))
            }
        }
    }

    class Normal_set_Date {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Test
        fun test() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]
                sheet["A1"] = TestUtil.getDateTime(2017, 9, 23, 13, 32, 24)

                KExcel.write(workbook, file.canonicalPath)
            }
            // 書き込んだものを再読込する
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                assertThat(sheet["A1"].toDate(), `is`(TestUtil.getDateTime(2017, 9, 23, 13, 32, 24)))
            }
        }
    }

    class Abnormal_set {
        @Rule
        @JvmField
        val tempFolder = TemporaryFolder()

        @Rule
        @JvmField
        val thrown = ExpectedException.none()

        @Test
        fun test() {
            val file = TestUtil.getTempWorkbookFile(tempFolder, "book1.xlsx")
            KExcel.open(file.canonicalPath).use { workbook ->
                val sheet = workbook[0]

                thrown.expect(IllegalArgumentException::class.java)
                sheet["A1"] = file
            }
        }
    }
}
