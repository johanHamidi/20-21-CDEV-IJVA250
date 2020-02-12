package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.repository.ArticleRepository;
import com.example.demo.service.impl.ClientServiceImpl;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.util.List;
import java.util.Optional;
import java.util.Date;
import java.util.Date;
import java.text.SimpleDateFormat;
import java.util.Calendar;

/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController<fileOutputStream> {

    private ArticleService articleService;
    private ClientServiceImpl clientServiceImpl;
    private FactureService factureService;

    public HomeController(ArticleService articleService, ClientServiceImpl clientService, FactureService factureService) {
        this.articleService = articleService;
        this.clientServiceImpl = clientService;
        this.factureService = factureService;
    }

    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Article> articles = articleService.findAll();
        modelAndView.addObject("articles", articles);

        List<Client> toto = clientServiceImpl.findAllClients();
        modelAndView.addObject("clients", toto);

        List<Facture> factures = factureService.findAllFactures();
        modelAndView.addObject("factures", factures);

        return modelAndView;
    }

    @GetMapping("/articles/csv")
    public void articleCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("test/csv");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-articles.csv\"");
        PrintWriter writer = response.getWriter();

        writer.println("Libelle;Prix");
        List<Article> articles = articleService.findAll();

        for (Article article : articles){
            String line = article.getLibelle() + ";" + article.getPrix();
            writer.println(line);
        }
    }

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("test/csv");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();

        writer.println("Nom;Prenom;Age");
        List<Client> clients = clientServiceImpl.findAllClients();

        for (Client client : clients){
            LocalDate now = LocalDate.now();
            Integer age = client.getDateNaissance().until(now).getYears();

            String line = client.getNom() + ";" + client.getPrenom()+ ";" + age;
            writer.println(line);
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-clients.xlsx\"");

        List<Client> clients = clientServiceImpl.findAllClients();

        Workbook workbook = new XSSFWorkbook();

        //Style
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.PINK.getIndex());


        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        /*XSSFColor color = new XSSFColor(new java.awt.Color(0, 37, 128));
        headerCellStyle.setTopBorderColor(color);*/

        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellNom = headerRow.createCell(0);
        cellNom.setCellValue("Nom");
        cellNom.setCellStyle(headerCellStyle);

        Cell cellPrenom = headerRow.createCell(1);
        cellPrenom.setCellValue("Prenom");
        cellPrenom.setCellStyle(headerCellStyle);

        Cell cellAge = headerRow.createCell(2);
        cellAge.setCellValue("Age");
        cellAge.setCellStyle(headerCellStyle);


        int rowNum = 1;

        for (Client client : clients){

            LocalDate now = LocalDate.now();
            Integer age = client.getDateNaissance().until(now).getYears();

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(client.getNom());
            row.createCell(1).setCellValue(client.getPrenom());
            row.createCell(2).setCellValue(age);
        }

        workbook.write(response.getOutputStream());

    }


    @GetMapping("/articles/xlsx")
    public void articlesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-articles.xlsx\"");


        List<Article> articles = articleService.findAll();

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Articles");
        Row headerRow = sheet.createRow(0);

        Cell cellLibelle = headerRow.createCell(0);
        cellLibelle.setCellValue("Libelle");

        Cell cellPrix = headerRow.createCell(1);
        cellPrix.setCellValue("Prix");



        int rowNum = 1;

        for (Article article : articles){

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(article.getLibelle());
            row.createCell(1).setCellValue(article.getPrix());
        }

        workbook.write(response.getOutputStream());

    }

    @GetMapping("/factures/xlsx")
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-factures.xlsx\"");


        List<Facture> factures = factureService.findAllFactures();

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Factures");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("id");

        Cell cellClient = headerRow.createCell(1);
        cellClient.setCellValue("Client");



        int rowNum = 1;

        for (Facture facture : factures){

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(facture.getId());
            row.createCell(1).setCellValue(facture.getClient().getNom());
        }

        workbook.write(response.getOutputStream());

    }

}
