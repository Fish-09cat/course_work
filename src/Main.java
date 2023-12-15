import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

import javax.swing.*;
import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String inputFilePath = "ИБС-21.xlsx";
        String outputFilePath = "Результат.xlsx";

        try {
            FileInputStream fis = new FileInputStream(inputFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            List<Student> students = analyzeData(sheet);

            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Results");

            writeStatistics(outputSheet, students);

            createChart(students);

            FileOutputStream fos = new FileOutputStream(outputFilePath);
            outputWorkbook.write(fos);
            fos.close();

            System.out.println("Анализ успешно выполнен. Результаты сохранены в файле " + outputFilePath);

            workbook.close();
            outputWorkbook.close();
            fis.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<Student> analyzeData(Sheet sheet) {
        List<Student> students = new ArrayList<>();

        for (Row row : sheet) {
            Cell nameCell = row.getCell(0);
            Cell ratingCell = row.getCell(1);

            String name = nameCell.getStringCellValue();
            int rating;

            if (ratingCell.getCellType() == CellType.NUMERIC) {
                rating = (int) ratingCell.getNumericCellValue();
            } else {
                continue;
            }

            students.add(new Student(name, rating));
        }

        return students;
    }

    private static void writeStatistics(Sheet outputSheet, List<Student> students) {
        int excellentCount = 0;
        int goodCount = 0;
        int satisfactoryCount = 0;
        int failCount = 0;
        double totalScore = 0;
        int studentCount = students.size();
        int maxRating = 0;

        List<String> excellentStudents = new ArrayList<>();
        List<String> goodStudents = new ArrayList<>();
        List<String> satisfactoryStudents = new ArrayList<>();
        List<String> failStudents = new ArrayList<>();

        for (Student student : students) {
            String name = student.getName();
            int rating = student.getRating();

            if (rating == 5) {
                excellentCount++;
                excellentStudents.add(name);
            } else if (rating == 4) {
                goodCount++;
                goodStudents.add(name);
            } else if (rating == 3) {
                satisfactoryCount++;
                satisfactoryStudents.add(name);
            } else if (rating == 2) {
                failCount++;
                failStudents.add(name);
            }

            if (maxRating < rating) {
                maxRating = rating;
            }

            totalScore += rating;
        }

        double averageScore = totalScore / studentCount;

        Row headerRow = outputSheet.createRow(0);
        headerRow.createCell(0).setCellValue("Количество отличников");
        headerRow.createCell(1).setCellValue("Количество хорошистов");
        headerRow.createCell(2).setCellValue("Количество троечников");
        headerRow.createCell(3).setCellValue("Количество не допущенных");
        headerRow.createCell(4).setCellValue("Средний балл");
        headerRow.createCell(5).setCellValue("Максимальная оценка");

        Row dataRow = outputSheet.createRow(1);
        dataRow.createCell(0).setCellValue(excellentCount);
        dataRow.createCell(1).setCellValue(goodCount);
        dataRow.createCell(2).setCellValue(satisfactoryCount);
        dataRow.createCell(3).setCellValue(failCount);
        dataRow.createCell(4).setCellValue(averageScore);
        dataRow.createCell(5).setCellValue(maxRating);

        int rowNum = 3;
        for (String name : excellentStudents) {
            Row row = outputSheet.createRow(rowNum++);
            row.createCell(0).setCellValue(name);
            row.createCell(1).setCellValue(5);
        }
        for (String name : goodStudents) {
            Row row = outputSheet.createRow(rowNum++);
            row.createCell(0).setCellValue(name);
            row.createCell(1).setCellValue(4);
        }
        for (String name : satisfactoryStudents) {
            Row row = outputSheet.createRow(rowNum++);
            row.createCell(0).setCellValue(name);
            row.createCell(1).setCellValue(3);
        }
        for (String name : failStudents) {
            Row row = outputSheet.createRow(rowNum++);
            row.createCell(0).setCellValue(name);
            row.createCell(1).setCellValue(2);
        }
    }

    private static void createChart(List<Student> students) {
        int excellentCount = 0;
        int goodCount = 0;
        int satisfactoryCount = 0;
        int failCount = 0;

        for (Student student : students) {
            int rating = student.getRating();

            if (rating == 5) {
                excellentCount++;
            } else if (rating == 4) {
                goodCount++;
            } else if (rating == 3) {
                satisfactoryCount++;
            } else if (rating == 2) {
                failCount++;
            }
        }

        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        dataset.addValue(excellentCount, "Отлично", "Отлично");
        dataset.addValue(goodCount, "Хорошо", "Хорошо");
        dataset.addValue(satisfactoryCount, "Удовлетворительно", "Удовлетворительно");
        dataset.addValue(failCount, "Не допущен", "Не допущен");

        JFreeChart barChart = ChartFactory.createBarChart(
                "Статистика оценок",
                "Оценка",
                "Количество",
                dataset,
                PlotOrientation.VERTICAL,
                true,
                true,
                false
        );

        JFrame frame = new JFrame("Статистика оценок");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);

        ChartPanel chartPanel = new ChartPanel(barChart);
        frame.getContentPane().add(chartPanel, BorderLayout.CENTER);
        frame.setVisible(true);
    }
}

