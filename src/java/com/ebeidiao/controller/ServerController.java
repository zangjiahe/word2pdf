package com.ebeidiao.controller;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

@RestController
@RequestMapping(value = "/office")
public class ServerController {
    @PostMapping(value = "/word2pdf")
    public ResponseEntity<byte[]> word2pdf(@RequestParam(name = "file") MultipartFile inputFile) throws IOException {
        // Ensure the input file is not empty and is of type DOCX
        if (inputFile.isEmpty()) {
            System.err.println("Invalid or missing DOCX file");
        }

        // Convert the Word document to PDF using documents4j
        ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
//        OutputStream outputStream = new FileOutputStream("C:/wwwroot/127.0.0.1/test.pdf");
        try (InputStream inputStream = inputFile.getInputStream()) {

            IConverter converter = LocalConverter.builder().build();
            converter.convert(inputStream)
                    .as(DocumentType.DOCX)
                    .to(pdfOutputStream)
                    .as(DocumentType.PDF)
                    .execute();
        } catch (Exception e) {
            System.err.println("Failed to convert DOCX to PDF:"+e.getMessage());
        }

        // Create a ResponseEntity with the converted PDF bytes and appropriate headers
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        String fileName = inputFile.getOriginalFilename() + ".pdf";
        headers.setContentDispositionFormData("attachment", fileName);
        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
        System.out.println(pdfOutputStream.toByteArray().length);
        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfOutputStream.toByteArray());
    }
}