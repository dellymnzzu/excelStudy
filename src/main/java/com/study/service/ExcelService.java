package com.study.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.regex.Pattern;

@Service
public class ExcelService {
    private static final Pattern phonePattern = Pattern.compile(
        "^(01[016789]-\\d{3,4}-\\d{4}$)");  // 휴대폰: 010-1234-5678, 011-123-4567

    private static final Pattern namePattern = Pattern.compile("^([가-힣]{2,5})");
    // ^ : 문자열의 시작 , [가-힣]: 한글 음절(가~힣)만 허용, {2,5} : 2글자 이상 5글자 이하, $ : 문자열의 끝

    private static final Pattern agePattern = Pattern.compile("^\\d{1,3}");

    public void readExcel(MultipartFile file) throws IOException {
        InputStream inputStream = file.getInputStream();
        Workbook workbook = new XSSFWorkbook(inputStream);  // 2007버전 이후


        Sheet workSheet = workbook.getSheetAt(0);
        // 헤더 행 처리
        Row headerRow = workSheet.getRow(0);
        if (headerRow != null) {
            System.out.println("헤더 : ");
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                String cellValue = cell.getStringCellValue();
                System.out.println("cellValue : " + cellValue);
            }
            System.out.println();
        }
        int rowNumLength = workSheet.getLastRowNum();
        System.out.println("rowNumLength : " + rowNumLength);
        int maxRows = Math.min(rowNumLength, workSheet.getLastRowNum());
        for (int rowNum = 1; rowNum <= maxRows; rowNum++) {
            Row row = workSheet.getRow(rowNum);
            if (row == null) {
                System.out.println("행 " + (rowNum) + ": ");
                continue;
            }
            System.out.print("행 " + (rowNum) + ": ");
            for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                Cell cell = row.getCell(cellNum);
                String cellValue = getCellValueAsString(cell);
                // 컬럼별 정규식 검증 및 처리
                String processedValue = processColumnData(cellNum, cellValue);
                System.out.print(processedValue + " ");
            }
            System.out.println();
        }
        workbook.close();
        inputStream.close();


    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // 숫자가 정수인 경우 소수점 제거
                double numericValue = cell.getNumericCellValue();
                  return String.valueOf((long) numericValue);
            default:
                return "";
        }
    }


    private String processColumnData(int columnIndex, String value) {
        if (value == null || value.trim().isEmpty()) {
            return "";
        }
        
        String trimmedValue = value.trim();
        
        switch (columnIndex) {
            case 0: // 이름 컬럼
                return processName(trimmedValue);
            case 1: // 나이 컬럼
                return processAge(trimmedValue);
            case 2: // 전화번호 컬럼
                return processPhoneNumber(trimmedValue);
            default:
                return trimmedValue; // 기타 컬럼은 그대로
        }
    }

    private String processName(String name) {
        if (namePattern.matcher(name).matches()) {
            return name;
        } else {
            System.out.println("잘못된 이름 형식: " + name);
            return name;
        }
    }
    private String processAge(String age) {
        if (agePattern.matcher(age).matches()) {
            int ageValue = Integer.parseInt(age);
            if (ageValue >= 0 && ageValue <= 150) {
                return age;
            } else {
                System.out.println("에러 :  " + age);
                return age;
            }
        } else {
            System.out.println("잘못된 나이 형식: " + age);
            return age;
        }
    }
    

    private String processPhoneNumber(String phone) {
        if (phonePattern.matcher(phone).matches()) {
            String normalizedPhone =
                    phone.replaceAll("-", "");
            return normalizedPhone;
        } else {
            System.out.println("잘못된 전화번호 형식: " + phone);
            return phone;
        }
    }
}




