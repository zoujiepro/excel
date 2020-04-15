package pub.vie.excel.example;

import pub.vie.excel.example.entity.studentInfo;
import pub.vie.excel.example.entity.studentInfo2;
import pub.vie.excel.read.ExcelReader;

import java.util.List;


/**
 * @Descrption :
 * @Author: zoujie
 * @Date: 2020-4-15
 */
public class Example {

    public static final String virtualPath = "students.xlsx";


    public static void main(String[] args) {
        Example example = new Example();
        example.testReadWithAnnotation();
    }


    public void testRead() {
        ExcelReader<studentInfo> reader = new ExcelReader<>();
        List<studentInfo> read = reader.read(ExcelReader.getStreamOnClassPath(virtualPath),1, studentInfo.class);

        if (read != null) {
            System.out.println(read.size());
        }
    }

    public void testReadWithAnnotation() {
        ExcelReader<studentInfo2> reader = new ExcelReader<>();
        List<studentInfo2> read = reader.read(studentInfo2.class);

        if (read != null) {
            System.out.println(read.size());
        }
    }

}
