package com.example.excel_spring.controller;


import com.example.excel_spring.model.Product;
import org.apache.coyote.http11.upgrade.UpgradeProcessorExternal;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.awt.*;
import java.io.IOException;
import java.util.*;
import java.util.List;

@RestController
@RequestMapping("api/v1")
public class ImportExcelController {

    @PostMapping("/importExcel")

    public ResponseEntity<List<Product>> importExcelFile (@RequestParam("file")MultipartFile files) throws IOException{
    HttpStatus status = HttpStatus.OK;
    List<Product> productList = new ArrayList<>();

    XSSFWorkbook workbook = new XSSFWorkbook(files.getInputStream());
    XSSFSheet worksheet = workbook.getSheetAt(0);

        for (int index = 0; index < worksheet.getPhysicalNumberOfRows(); index++) {
        if (index > 0) {
            Product product = new Product();

            XSSFRow row = worksheet.getRow(index);
            Integer id = (int) row.getCell(0).getNumericCellValue();

            product.setId(id.toString());
            product.setProductName(((XSSFRow) row).getCell(1).getStringCellValue());
            product.setPrice(row.getCell(2).getNumericCellValue());
            product.setCategory(row.getCell(3).getStringCellValue());

            productList.add(product);
        }
    }

        return new ResponseEntity<>(productList, status);
}
}
