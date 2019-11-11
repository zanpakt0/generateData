package main.java;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.Math;
import java.text.SimpleDateFormat;
import java.util.Date;


class Human {
    private String name;
    private String sex;
    private String surName;
    private String patronymic;
    private String dateOfBirth;
    private int age;
    private String cityBirth;
    private int postIndex;
    private String country;
    private String region;
    private String cityResidence;
    private String street;
    private int house;
    private int flat;

    Human(String name, String sex) {
        this.name = name;
        this.sex = sex;
    }

    String getName() {
        return name;
    }

    String getSex() {
        return sex;
    }

    String getSurName() {
        return surName;
    }

    String getPatronymic() {
        return patronymic;
    }

    String getDateOfBirth() {
        return dateOfBirth;
    }

    int getAge() {
        return age;
    }

    String getCityBirth() {
        return cityBirth;
    }

    int getPostIndex() {
        return postIndex;
    }

    String getCountry() {
        return country;
    }

    String getRegion() {
        return region;
    }

    String getCityResidence() {
        return cityResidence;
    }

    String getStreet() {
        return street;
    }

    int getHouse() {
        return house;
    }

    int getFlat() {
        return flat;
    }

    void setSurname() throws IOException {
        HSSFWorkbook wb;
        if (this.sex.equals("МУЖ")) {
            wb = new HSSFWorkbook(new FileInputStream("manSurname.xls"));
        } else {
            wb = new HSSFWorkbook(new FileInputStream("womanSurname.xls"));
        }
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        this.surName = wb.getSheetAt(0).getRow(index).getCell(0).getRichStringCellValue().getString();
    }

    void setPatronymic() throws IOException {
        HSSFWorkbook wb;
        if (this.sex.equals("МУЖ")) {
            wb = new HSSFWorkbook(new FileInputStream("manPatronymic.xls"));
        } else {
            wb = new HSSFWorkbook(new FileInputStream("womanPatronymic.xls"));
        }
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        this.patronymic = wb.getSheetAt(0).getRow(index).getCell(0).getRichStringCellValue().getString();
    }

    void setDateOfBirthAndAge() {
        Date currentDate;
        currentDate = new Date((long) (Math.random() * (new Date().getTime())));
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        SimpleDateFormat yearFormat = new SimpleDateFormat("y");
        this.dateOfBirth = dateFormat.format(currentDate);
        Date dTime = new Date(new Date().getTime() - currentDate.getTime());
        this.age = Integer.parseInt(yearFormat.format(dTime)) - 1970;
    }

    void setCityBirth() throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream("city.xls"));
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        this.cityBirth = wb.getSheetAt(0).getRow(index).getCell(0).getRichStringCellValue().getString();
    }

    void setPostIndex() {
        this.postIndex = (int) (100000 + Math.random() * 900000);
    }

    void setCountry() throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream("country.xls"));
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        this.country = wb.getSheetAt(0).getRow(index).getCell(0).getRichStringCellValue().getString();
    }

    void setRegion() throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream("region.xls"));
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        this.region = wb.getSheetAt(0).getRow(index).getCell(0).getRichStringCellValue().getString();
    }

    void setCityResidence() throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream("city.xls"));
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        this.cityResidence = wb.getSheetAt(0).getRow(index).getCell(0).getRichStringCellValue().getString();
    }

    void setStreet() throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream("street.xls"));
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        this.street = wb.getSheetAt(0).getRow(index).getCell(0).getRichStringCellValue().getString();
    }

    void setHouse() {
        this.house = (int) (1 + Math.random() * 300);
    }

    void setFlat() {
        this.flat = (int) (1 + Math.random() * 500);
    }
}

class Generator {

    static Human[] generate(int count) throws IOException {
        Human[] humans = new Human[count];
        for (int i = 0; i < count; i++) {
            Human human = getNameAndSex();
            human.setSurname();
            human.setPatronymic();
            human.setDateOfBirthAndAge();
            human.setCityBirth();
            human.setPostIndex();
            human.setCountry();
            human.setRegion();
            human.setCityResidence();
            human.setHouse();
            human.setFlat();
            human.setStreet();
            humans[i] = human;
        }
        return humans;
    }

    private static Human getNameAndSex() {
        HSSFWorkbook wb = null;
        try {
            wb = new HSSFWorkbook(new FileInputStream("names.xls"));
        } catch (IOException e) {
            e.printStackTrace();
        }
        assert wb != null;
        HSSFSheet sheet = wb.getSheetAt(0);
        int index = (int) (Math.random() * wb.getSheetAt(0).getPhysicalNumberOfRows());
        HSSFRow row = sheet.getRow(index);
        return new Human(row.getCell(0).getRichStringCellValue().getString(), row.getCell(1).getRichStringCellValue().getString());
    }

}

class RighterToExcel {

    static void righter(Human[] humans, String fileName) throws IOException {
        String[] column = {"Имя", "Фамилия", "Отчество", "Возраст", "Пол", "Дата рождения", "Место рождения",
                "Индекс", "Страна", "Область", "Город", "Улица", "Дом", "Квартира",};
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("Sheet0");
        int rowNum = 0;
        HSSFRow row = sheet.createRow(rowNum);
        for (int i = 0; i < 14; i++) {
            row.createCell(i).setCellValue(column[i]);
        }
        for (Human human : humans) {
            createSheet(sheet, ++rowNum, human);
        }
        String path = String.valueOf(new File(fileName).getCanonicalFile());
        FileOutputStream out = new FileOutputStream(fileName);
        wb.write(out);
        System.out.println("Файл создан. Путь: " + new File(path).getAbsolutePath());
        out.close();

    }

    private static void createSheet(HSSFSheet sheet, int rowNum, Human human) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(human.getName());
        row.createCell(1).setCellValue(human.getSurName());
        row.createCell(2).setCellValue(human.getPatronymic());
        row.createCell(3).setCellValue(human.getAge());
        row.createCell(4).setCellValue(human.getSex());
        row.createCell(5).setCellValue(human.getDateOfBirth());
        row.createCell(6).setCellValue(human.getCityBirth());
        row.createCell(7).setCellValue(human.getPostIndex());
        row.createCell(8).setCellValue(human.getCountry());
        row.createCell(9).setCellValue(human.getRegion());
        row.createCell(10).setCellValue(human.getCityResidence());
        row.createCell(11).setCellValue(human.getStreet());
        row.createCell(12).setCellValue(human.getHouse());
        row.createCell(13).setCellValue(human.getFlat());

    }
}