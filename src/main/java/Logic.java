import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import javax.swing.*;
import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.HashMap;

/**
 * This program makes different list for canteen, transport ...
 */
class First {
    /**
     * This workbook imported from the Excel working shield.
     */
    private static HSSFWorkbook workbook;
    /**
     * Allow us to use current date.
     */
    private static final LocalDate TODAY = LocalDate.now();
    /**
     * JFileChooser allow us to import Excel files
     * when user out of office or files are relocated.
     */
    private static final JFileChooser OPENFILE = new JFileChooser();

    public static void main(String[] args) throws IOException {

        String[] months = {"", "Январь", "Февраль",
                "Март", "Апрель", "Май", "Июнь",
                "Июль", "Август", "Сентябрь",
                "Октябрь", "Ноябрь", "Декабрь"};


        int month = TODAY.getMonthValue();
        workbook = getGrafik(months[month]);


        JFrame frame = new JFrame("Развозка");
        JButton btnTransport = new JButton("Развозка");
        JButton btnMedic = new JButton("Мед.осмотр");
        JButton btnDinner = new JButton("Обеды + молоко");
        JButton btnRefresh = new JButton("Обновить данные");

        JTextField dateField = new JTextField(
                Integer.toString(TODAY.plusDays(1).
                        getDayOfMonth()), 10);
        JTextField monthField = new JTextField(months[month], 10);
        dateField.setBorder(BorderFactory.
                createTitledBorder("Дата развозки"));
        monthField.setBorder((BorderFactory.createTitledBorder("Месяц")));
        JPanel panel = new JPanel();

        frame.setSize(new Dimension(400, 150));
        frame.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        panel.setLayout(new FlowLayout());
        dateField.setSize(25, 10);
        btnTransport.setSize(25, 10);
        btnMedic.setSize(25, 10);
        btnDinner.setSize(25, 10);
        btnRefresh.setSize(25, 10);

        frame.add(panel);
        panel.add(btnTransport);
        panel.add(dateField);
        panel.add(monthField);
        panel.add(btnMedic);
        panel.add(btnDinner);
        panel.add(btnRefresh);

        btnTransport.addActionListener(e -> {
            try {
                createTransportTable(dateField, workbook);

            } catch (IOException ioException) {
                JOptionPane.showMessageDialog(null,
                        "Вам не хватает данных"
                                + " чтобы составить сиписки"
                                + " на развозку.");
                ioException.printStackTrace();
            }
        });

        btnMedic.addActionListener(e -> {
            try {
                medic(months[month], TODAY, workbook);
            } catch (IOException ioException) {
                ioException.printStackTrace();
            }
        });

        btnDinner.addActionListener(e -> {
            try {
                dinner(workbook, TODAY);
            } catch (IOException ioException) {
                ioException.printStackTrace();
            }
        });

        btnRefresh.addActionListener(e -> {
            try {
                workbook = getGrafik(months[month]);
            } catch (IOException fileNotFoundException) {
                fileNotFoundException.printStackTrace();
            }
            JOptionPane.showMessageDialog(null,
                    "Новые данные получены!");
        });

        frame.setVisible(true);
    }

    // ********************************************************************
    public static HSSFWorkbook getGrafik(
            final String month1) throws IOException {
        HSSFWorkbook wb;
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(
                    "M:\\Рабочее время_1\\"
                            + "Technical\\Розлив\\График сменности "
                            + TODAY.getYear() + "\\" + month1 + ".xls");
            wb = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            String path = null;
            while (path == null) {
                int ret = OPENFILE.showDialog(
                        null, "Открыть график");
                if (ret == JFileChooser.APPROVE_OPTION) {
                    path = OPENFILE.getSelectedFile().getPath();
                }
                if (path == null) {
                    JOptionPane.showMessageDialog(
                            null, "Для работы программы Вам"
                                    + " необходимо загрузить"
                                    + " \"График сменности\"");
                }
            }
            inputStream = new FileInputStream(path);
            wb = new HSSFWorkbook(inputStream);
        } finally {
            assert inputStream != null;
            inputStream.close();
        }
        return wb;
    }

    // **************************************************
    public static void dinner(HSSFWorkbook grafik,
                              LocalDate today) throws IOException {
        HSSFWorkbook workbookDinner = importTable(
                "D:\\ExelExample\\обеды-розлив.xls",
                "Открыть файл обедов");

        assert workbookDinner != null;
        HSSFSheet sheetDinner = workbookDinner.getSheet("обеды");
        HSSFSheet sheetMilk = workbookDinner.
                getSheet("молоко (вредники)");
        HSSFSheet sheetSource = grafik.getSheet("август");

        long monday = today.plusDays(
                8 - today.getDayOfWeek().getValue()).getDayOfMonth();

        int rowOffset = 11;
        int cloumnOffset = 7;

        int startIndDinnerCol2000 = 14;
        int startIndMilkCol800 = 2;

        for (int i = 0; i < 7; i++) {
            int col = i + (int) (monday + cloumnOffset);
            int startIndDinnerRow2000 = 22;
            int startIndMilkRow800 = 8;
            int startIndMilkRow2000 = 29;
            int cnt = 0;

            String date = today.plusDays(
                    8 - today.getDayOfWeek().getValue() + i).toString();
            sheetDinner.getRow(4).getCell(14 + i).setCellValue(date);
            sheetMilk.getRow(5).getCell(2 + i).setCellValue(date);

            for (int row = rowOffset; row < 55; row++) {
                Cell name = sheetSource.getRow(row).getCell(3);
                Cell shift = sheetSource.getRow(row).getCell(col);
                Cell milk800 = sheetMilk.getRow(
                        startIndMilkRow800).getCell(startIndMilkCol800);
                Cell milk2000 = sheetMilk.getRow(
                        startIndMilkRow2000).getCell(startIndMilkCol800);
                Cell dinn = sheetDinner.getRow(
                        startIndDinnerRow2000).getCell(startIndDinnerCol2000);
                if (shift != null && shift.getCellType()
                        == Cell.CELL_TYPE_NUMERIC) {
                    if (shift.getNumericCellValue() == 1.0 || shift.getNumericCellValue() == 8.0) {
                        milk800.setCellValue(name.getStringCellValue());
                        startIndMilkRow800++;
                        cnt++;
                    }
                    if (shift.getNumericCellValue() == 2.0) {
                        dinn.setCellValue(name.getStringCellValue());
                        milk2000.setCellValue(name.getStringCellValue());
                        startIndDinnerRow2000++;
                        startIndMilkRow2000++;
                    }
                }
            }
            if (i < 5) {
                sheetDinner.getRow(7).getCell(14 + i).setCellValue(cnt + 1);
            } else {
                sheetDinner.getRow(7).getCell(14 + i).setCellValue(cnt);
            }
            startIndMilkCol800++;
            startIndDinnerCol2000++;
        }

        FileOutputStream outputStreamDinner = null;
        try {
            outputStreamDinner = new FileOutputStream(
                    "M:\\Technical\\KARAGANDA FILLLING LINE\\"
                            + "Юля\\обеды-розлив1.xls");
            workbookDinner.write(outputStreamDinner);
            outputStreamDinner.close();
        } catch (Exception e) {
            outputStreamDinner = getFileOutputStream(
                    workbookDinner,
                    "Сохраните файл \"обеды\"");
        } finally {
            assert outputStreamDinner != null;
            outputStreamDinner.close();
        }

        JOptionPane.showMessageDialog(
                null, "Всё готово, обеды"
                        + " можно отправлять!");

    }


    //**********************************************************
    public static void medic(String month, LocalDate d,
                             HSSFWorkbook workbook) throws IOException {
        System.out.println(workbook);
        HSSFSheet sheet = workbook.getSheet("август");
        HSSFWorkbook workbookMedic = importTable("M:\\Technical\\"
                + "KARAGANDA FILLLING LINE\\Юля\\мед. осмотр  "
                + d.getYear() + ".xls", "Открыть мед. осмотр");
        HSSFWorkbook workbookWorkers = importTable(
                "Отчет по количеству персонала на заводе "
                        + "Караганда 02.10.2021.xls",
                "Открыть отчет по персоналу");
        HSSFWorkbook workbookReserve = importTable(
                "M:\\Technical\\KARAGANDA FILLLING LINE"
                        + "\\Юля\\Наемники "
                        + d.getYear() + "\\Наёмники " + month
                        + " " + d.getYear() + ".xls",
                "Открыть заявку по наёмникам");

        assert workbookReserve != null;
        HSSFSheet sheetReserve = workbookReserve.
                getSheetAt(d.getDayOfMonth() - 1);
        assert workbookWorkers != null;
        HSSFSheet sheetWorkers = workbookWorkers.
                getSheet("Технический департамент");
        Row rowWorkersDay = sheetWorkers.getRow(10);
        Row rowWorkersNight = sheetWorkers.getRow(25);

        assert workbookMedic != null;
        HSSFSheet sheetMedic = workbookMedic.getSheet(month);
        HashMap<String, Integer> hmMwedic = new HashMap<>();
        for (int i = 10; i < 42; ++i) {
            hmMwedic.put(sheetMedic.getRow(i).
                    getCell(1).getStringCellValue(), i);
        }

        int columnOffset = 7;
        int rowOffset = 8;
        CellStyle dayOff = sheetMedic.getRow(47).getCell(7).getCellStyle();
        CellStyle dayShift = sheetMedic.getRow(49).getCell(7).getCellStyle();
        CellStyle nightShift = sheetMedic.getRow(51).getCell(7).getCellStyle();
        CellStyle sick = sheetMedic.getRow(53).getCell(7).getCellStyle();
        CellStyle vocation = sheetMedic.getRow(55).getCell(7).getCellStyle();

        int cntEfesDay = 0;
        int cntEfesNight = 0;
        int cntElitDay = 0;
        int cntElitNight = 0;


        for (int i = rowOffset; i < 55; i++) {
            String key = sheet.getRow(i).getCell(3).getStringCellValue();
            Cell shiftCell = sheet.getRow(i).
                    getCell(d.getDayOfMonth() + columnOffset);
            if (shiftCell != null && shiftCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                if (shiftCell.getNumericCellValue() == 2.0) {
                    if (i < 48) {
                        cntEfesNight++;
                    } else {
                        cntElitNight++;
                    }
                } else {
                    if (i < 48) {
                        cntEfesDay++;
                    } else {
                        cntElitDay++;
                    }
                }
            }
            if (hmMwedic.containsKey(key)) {
                Cell cell = sheetMedic.getRow(hmMwedic.get(key))
                        .getCell(d.getDayOfMonth() + 2);
                if (shiftCell != null && shiftCell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    if (shiftCell.getNumericCellValue() == 2.0) {
                        cell.setCellValue("з");
                        cell.setCellStyle(nightShift);
                    } else {
                        cell.setCellValue("з");
                        cell.setCellStyle(dayShift);
                    }
                }
                if (shiftCell != null && shiftCell.getCellType() == Cell.CELL_TYPE_STRING) {
                    if (shiftCell.getStringCellValue().equals("в")) {
                        cell.setCellValue("в");
                        cell.setCellStyle(dayOff);
                    }
                    if (shiftCell.getStringCellValue().equals("о")) {
                        cell.setCellValue("о");
                        cell.setCellStyle(vocation);
                    }
                    if (shiftCell.getStringCellValue().equals("б")) {
                        cell.setCellValue("б");
                        cell.setCellStyle(sick);
                    }
                }
            }
        }

        rowWorkersDay.getCell(3).setCellValue(cntEfesDay);
        rowWorkersDay.getCell(4).setCellValue(sheetReserve.
                getRow(20).getCell(1).getNumericCellValue());
        rowWorkersDay.getCell(5).setCellValue(cntElitDay);
        rowWorkersNight.getCell(3).setCellValue(cntEfesNight);
        rowWorkersNight.getCell(4).setCellValue(sheetReserve.
                getRow(20).getCell(2).getNumericCellValue());
        rowWorkersNight.getCell(5).setCellValue(cntElitNight);
        FileOutputStream outputStreamMedic = null;
        try {
            outputStreamMedic = new FileOutputStream(
                    "M:\\Technical\\KARAGANDA FILLLING"
                            + " LINE\\Юля\\мед. осмотр  "
                            + d.getYear() + ".xls");
            workbookMedic.write(outputStreamMedic);
            outputStreamMedic.close();
        } catch (Exception e) {
            outputStreamMedic = getFileOutputStream(
                    workbookMedic, "Схранить мед. осмотр");
        } finally {
            assert outputStreamMedic != null;
            outputStreamMedic.close();
        }

        FileOutputStream outputStreamWorkers = null;
        try {
            outputStreamWorkers = new FileOutputStream("M:Technical\\"
                    + "KARAGANDA FILLLING LINE\\Юля\\Количество людей\\"
                    + d.getYear() + "\\"
                    + month + "\\Отчет по количеству персонала на заводе Караганда"
                    + d.getDayOfMonth() + "." + d.getMonthValue()
                    + "." + d.getYear() + ".xls");
            workbookWorkers.write(outputStreamWorkers);

        } catch (
                Exception e) {
            outputStreamWorkers = getFileOutputStream(
                    workbookWorkers,
                    "Сохранить отчет по количеству персонала");
        } finally {
            assert outputStreamWorkers != null;
            outputStreamWorkers.close();
        }
        JOptionPane.showMessageDialog(null,
                "Всё готово, можно отправлять!");
    }

    //***************************************************
    public static void createTransportTable(
            final JTextField dateF,
            final HSSFWorkbook workbook) throws IOException {
        LocalDate d = LocalDate.now();
        final int columnOffset = 7;
        final int rowOffset = 8;
        final int lastRowInd = 55;
        final int nameColumn = 3, transportDemandColumn = 4;
        int startRow800 = 4;
        int startRow2015 = 4;
        int startRow2000 = 4;
        int startRow2000yesterday = 4;
        final int nColAdrList = 6, nRowAdrList = 30;
        HSSFWorkbook workbookTransport = importTable(
                "21.09.2021.xls", "Открыть заявку на развозку");
        HashMap<String, String[]> adrList = new HashMap<>();
        HSSFSheet sheet = workbook.getSheet("август");
        assert workbookTransport != null;
        // получаем лист Приезд 08-00
        HSSFSheet sheetTransport800 = workbookTransport.
                getSheet("Приезд 08-00");
        // получаем лист Приезд 20-00
        HSSFSheet sheetTransport2000 = workbookTransport.
                getSheet("Приезд 20-00");
        // получаем лист Приезд 08-00
        HSSFSheet sheetTransport815 = workbookTransport.
                getSheet("Развоз 08-15");
        // получаем лист Приезд 20-15
        HSSFSheet sheetTransport2015 = workbookTransport.
                getSheet("Развоз 20-15");
        // список адресов
        HSSFSheet sheetAdr = workbookTransport.
                getSheet("Лист1");


//        Переносим данные из Лист1 в Hash таблицу
        for (int rowi = 2; rowi < nRowAdrList; ++rowi) {
            String[] adr = new String[nColAdrList];
            for (int j = 0; j < adr.length; ++j) {
                Cell cell = (sheetAdr.getRow(rowi).getCell(j + 1));

                if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    adr[j] = cell.getStringCellValue();
                }
            }
            adrList.put(sheetAdr.getRow(rowi).getCell(1).
                    getStringCellValue(), adr);
        }

        int desiredDate = Integer.parseInt(dateF.getText());

        for (int i = rowOffset; i < lastRowInd; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(desiredDate + columnOffset);
            Cell cellYesterday = row.getCell(desiredDate - 1 + columnOffset);

            if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC && row.
                    getCell(transportDemandColumn).
                    getStringCellValue().equals("развоз")) {
                if (cell.getNumericCellValue() == 1.0) {
                    Row trRow800 = sheetTransport800.getRow(startRow800);
                    Row trRow2015 = sheetTransport2015.getRow(startRow2015);
                    String res = row.getCell(nameColumn).getStringCellValue();
                    String[] adrData = adrList.get(res);
                    for (int k = 0; k < adrData.length; ++k) {
                        Cell trCell800 = trRow800.getCell(k + 1);
                        Cell trCell2015 = trRow2015.getCell(k + 1);
                        trCell800.setCellValue(adrData[k]);
                        trCell2015.setCellValue(adrData[k]);
                    }
                    startRow800++;
                    startRow2015++;
                } else if (cell.getNumericCellValue() == 2.0) {
                    String res1 = row.getCell(nameColumn).getStringCellValue();
                    Row trRow2000 = sheetTransport2000.getRow(startRow2000);
                    String[] adrData = adrList.get(res1);
                    for (int k = 0; k < adrData.length; ++k) {
                        Cell trCell2000 = trRow2000.getCell(k + 1);
                        trCell2000.setCellValue(adrData[k]);
                    }
                    startRow2000++;
                } else {
                    Row trRow800 = sheetTransport800.getRow(startRow800);
                    String res = row.getCell(nameColumn).getStringCellValue();
                    String[] adrData = adrList.get(res);
                    for (int k = 0; k < adrData.length; ++k) {
                        Cell trCell800 = trRow800.getCell(k + 1);
                        trCell800.setCellValue(adrData[k]);
                    }
                    startRow800++;
                }
            }
            if (cellYesterday != null && cellYesterday.getCellType() == Cell.CELL_TYPE_NUMERIC && row.
                    getCell(transportDemandColumn).getStringCellValue().
                    equals("развоз") && cellYesterday.
                    getNumericCellValue() == 2.0) {
                Row trRow815 = sheetTransport815.
                        getRow(startRow2000yesterday);
                String[] adrData = adrList.
                        get(row.getCell(nameColumn).getStringCellValue());

                for (int k = 0; k < adrData.length; ++k) {
                    Cell trCELL815 = trRow815.getCell(k + 1);
                    trCELL815.setCellValue(adrData[k]);
                }
                startRow2000yesterday++;
            }
        }

        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream("M:\\Technical\\"
                    + "KARAGANDA FILLLING LINE\\Юля\\Адрес"
                    + "   " + d.getYear() + "\\"
                    + desiredDate + "." + d.getMonthValue()
                    + "." + d.getYear() + ".xls");
            workbookTransport.write(fileOut);
            JOptionPane.showMessageDialog(null,
                    "Файл развозки создан,"
                            + " можно отправлять!");
        } catch (Exception e) {
            fileOut = getFileOutputStream(workbookTransport,
                    "Coхранить файл развозки");
        } finally {
            assert fileOut != null;
            fileOut.close();
        }

    }

    public static FileOutputStream getFileOutputStream(
            final HSSFWorkbook workbook,
            final String title) throws IOException {
        String path = null;
        int ret = OPENFILE.showDialog(null, title);
        if (ret == JFileChooser.APPROVE_OPTION) {
            path = OPENFILE.getSelectedFile().getPath();
        }
        if (ret == JFileChooser.CANCEL_OPTION) {
            return null;
        }
        assert path != null;
        FileOutputStream fileOut = new FileOutputStream(path);
        workbook.write(fileOut);
        return fileOut;
    }

    /**
     * This method import Excel table to our program
     */
    public static HSSFWorkbook importTable(
            final String path,
            final String title) throws IOException {
        HSSFWorkbook curWorkbook;
        FileInputStream inputStream = null;
        try {
            inputStream = new FileInputStream(path);
            curWorkbook = new HSSFWorkbook(inputStream);
        } catch (Exception e) {
            String path1 = null;
            int ret = OPENFILE.showDialog(null, title);
            if (ret == JFileChooser.APPROVE_OPTION) {
                path1 = OPENFILE.getSelectedFile().getPath();
            }
            if (ret == JFileChooser.CANCEL_OPTION) {
                System.out.println("Отмена");
                return null;
            }
            assert path1 != null;
            inputStream = new FileInputStream(path1);
            curWorkbook = new HSSFWorkbook(inputStream);
        } finally {
            assert inputStream != null;
            inputStream.close();
        }
        return curWorkbook;
    }
}
