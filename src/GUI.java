import javax.swing.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.Iterator;

public class GUI extends JFrame {

    private static final String EXCEL_FILE_PATH = "C:/Users/acer/Downloads/databas.xlsx";
    private JPanel panel;
    private JTextField textField1;
    private JTextField textField2;
    private JTextField textField3;
    private JTextField textField4;
    private JTextField textField5;
    private JTextField textField6;
    private JTextField textField7;
    private JTextField textField8;
    private JTextField searchTextField;

    public GUI() {
        setTitle("Лабораторна робота 6-7");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(800, 600);
        setLocationRelativeTo(null); // Розміщення по центру екрану
        panel = new JPanel(new GridLayout(0, 2));


        // Ініціалізація текстових полів
        textField1 = new JTextField(10);
        textField2 = new JTextField(10);
        textField3 = new JTextField(10);
        textField4 = new JTextField(10);
        textField5 = new JTextField(10);
        textField6 = new JTextField(10);
        textField7 = new JTextField(10);
        textField8 = new JTextField(10);
        searchTextField = new JTextField(10);

        // Додавання кнопok
        JButton saveButton = new JButton("Створити нову команду");
        JButton saveFileButton = new JButton("Створити звіт(.xlsx)");
        JButton searchButton = new JButton("Пошук");



        // Обробник події для кнопok
        saveButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
                    saveDataToExcel(
                            textField1.getText(),
                            textField2.getText(),
                            textField3.getText(),
                            textField4.getText(),
                            textField5.getText(),
                            textField6.getText(),
                            textField7.getText(),
                            textField8.getText()
                    );
                    JOptionPane.showMessageDialog(null, "Дані успішно збережено у файл.");
                } catch (IOException ex) {
                    ex.printStackTrace();
                    JOptionPane.showMessageDialog(null, "Сталася помилка при збереженні даних.");
                }
            }
        });

        saveFileButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
                    saveFileToExcel();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });

        searchButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                String searchValue = searchTextField.getText();
                try {
                    searchByCriteria("Команда", searchValue);
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });


        // Додавання лейблів, текстових полів та кнопok до панелі та контейнера

        panel.add(new JLabel("Команда"));
        panel.add(textField1);
        panel.add(new JLabel("Зіграні матчі"));
        panel.add(textField2);
        panel.add(new JLabel("Перемоги"));
        panel.add(textField3);
        panel.add(new JLabel("Нічиї"));
        panel.add(textField4);
        panel.add(new JLabel("Поразки"));
        panel.add(textField5);
        panel.add(new JLabel("Забито"));
        panel.add(textField6);
        panel.add(new JLabel("Пропущено"));
        panel.add(textField7);
        panel.add(new JLabel("Очки"));
        panel.add(textField8);
        panel.add(new JLabel("Введіть назву команди, яку потрібно знайти"));
        panel.add(searchTextField);

        getContentPane().add(panel);
        getContentPane().add(saveButton);
        getContentPane().add(saveFileButton);
        getContentPane().add(searchButton);

        // Встановлення менеджера компонування FlowLayout
        setLayout(new FlowLayout());
    }

    //Додаємо дані до БД
    private static void saveDataToExcel(String data1, String data2, String data3, String data4,
                                        String data5, String data6, String data7, String data8) throws IOException {
        FileInputStream inputStream = new FileInputStream(EXCEL_FILE_PATH);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        int lastRowIndex = sheet.getLastRowNum();
        Row row = sheet.createRow(lastRowIndex + 1);

        Cell cell1 = row.createCell(0);
        cell1.setCellValue(data1);

        Cell cell2 = row.createCell(1);
        cell2.setCellValue(data2);

        Cell cell3 = row.createCell(2);
        cell3.setCellValue(data3);

        Cell cell4 = row.createCell(3);
        cell4.setCellValue(data4);

        Cell cell5 = row.createCell(4);
        cell5.setCellValue(data5);

        Cell cell6 = row.createCell(5);
        cell6.setCellValue(data6);

        Cell cell7 = row.createCell(6);
        cell7.setCellValue(data7);

        Cell cell8 = row.createCell(7);
        cell8.setCellValue(data8);

        FileOutputStream outputStream = new FileOutputStream(EXCEL_FILE_PATH);
        workbook.write(outputStream);

        workbook.close();
        inputStream.close();
        outputStream.close();
    }

    //Формування звіту у ексель файл, який зберігаємо у директорії, яку обирає користувач
    private void saveFileToExcel() throws IOException {
        FileInputStream inputStream = new FileInputStream(EXCEL_FILE_PATH);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        inputStream.close();

        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showSaveDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            String filePath = selectedFile.getAbsolutePath();

            if (!filePath.endsWith(".xlsx")) {
                filePath += ".xlsx";
            }

            FileOutputStream outputStream = new FileOutputStream(filePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            JOptionPane.showMessageDialog(null, "Файл успішно збережено.");
        }
    }

    // Пошук команди за її назвою
    private static void searchByCriteria(String columnName, String searchValue) throws IOException {
        FileInputStream inputStream = new FileInputStream(EXCEL_FILE_PATH);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        boolean found = false;

        int columnIndex = -1; // Індекс стовпця, де шукаємо
        Row headerRow = sheet.getRow(0);

        // Знаходимо індекс стовпця за назвою columnName
        for (Cell cell : headerRow) {
            if (cell.getStringCellValue().equals(columnName)) {
                columnIndex = cell.getColumnIndex();
                break;
            }
        }
        DataFormatter dataFormatter = new DataFormatter();
        if (columnIndex != -1) {
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next();

            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                String cellValue = dataFormatter.formatCellValue(cell);

                if (cellValue.equals(searchValue)) {
                    found = true;
                    break;
                }
            }
        }


        workbook.close();
        inputStream.close();

        if (found) {
            JOptionPane.showMessageDialog(null,"Слово '" + searchValue + "' було знайдено у стовпці '" + columnName + "'.");
        } else {
            JOptionPane.showMessageDialog(null, "Слово '" + searchValue + "' не знайдено у стовпці '" + columnName + "'." );
        }
    }
}
