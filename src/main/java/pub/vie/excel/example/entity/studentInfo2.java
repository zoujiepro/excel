package pub.vie.excel.example.entity;

import pub.vie.excel.common.annotation.ExcelEntity;
import pub.vie.excel.common.annotation.ExcelField;
import pub.vie.excel.common.constant.ExcelCellIndex;

import java.util.Date;

@ExcelEntity(classPathSource = "students.xlsx", sheetAt = 0, skip = 1, limitRow = -1)
public class studentInfo2 {

    @ExcelField(colIndex = ExcelCellIndex.CELL_INDEX_A)
    private String id;

    @ExcelField(colIndex = ExcelCellIndex.CELL_INDEX_B)
    private String name;

//    @ExcelField(colIndex = ExcelCellIndex.CELL_INDEX_C)
//    private int age;

    @ExcelField(colIndex = ExcelCellIndex.CELL_INDEX_D)
    private Date entranceDate;

}