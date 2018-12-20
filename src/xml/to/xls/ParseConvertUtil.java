/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xml.to.xls;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.paint.Color;
import javax.swing.filechooser.FileSystemView;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.*;

/**
 *
 * @author Arkel
 */
public class ParseConvertUtil {

    private String dir = FileSystemView.getFileSystemView().getDefaultDirectory().getPath() + "/xml to xls";
    private CellStyle styleHead, styleNormal;

    public void readXml(File xmlFile) {
        try {
            File xlsFile = new File(dir + "/" + xmlFile.getName().substring(0, xmlFile.getName().length() - 3) + "xlsx");

            XSSFWorkbook workbook = new XSSFWorkbook();

            styleHead = workbook.createCellStyle();
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            styleHead.setFont(boldFont);
            styleHead.setAlignment(HorizontalAlignment.CENTER_SELECTION);
            styleHead.setFillBackgroundColor(IndexedColors.BLUE.getIndex());

            styleNormal = workbook.createCellStyle();
            styleNormal.setAlignment(HorizontalAlignment.CENTER_SELECTION);
            styleNormal.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
            ArrayList<XSSFSheet> listSheets = new ArrayList<>();

            DocumentBuilderFactory builderFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = builderFactory.newDocumentBuilder();
            Document document = builder.parse(xmlFile);
            document.getDocumentElement().normalize();
            NodeList listNode = document.getDocumentElement().getChildNodes();
            ArrayList<String> sheetsNames = new ArrayList<>();

            for (int i = 0; i < listNode.getLength(); i++) {
                final String sheetName = listNode.item(i).getNodeName().replaceAll("\r", "").replaceAll("\n", "").trim();
                if (!sheetsNames.contains(sheetName) && sheetName.toLowerCase() != "#text" && sheetName != " " && sheetName.length() != 0) {
                    sheetsNames.add(listNode.item(i).getNodeName());
                    listSheets.add(workbook.createSheet(sheetsNames.get(sheetsNames.size() - 1)));
                    XSSFRow row = listSheets.stream().filter(sheetl -> sheetl.getSheetName() == sheetName).findFirst().get().createRow(0);
                    setColsNames(row, listNode.item(i));
                }
            }
            for (int i = 0; i < listNode.getLength(); i++) {
                System.out.println(i);
                final String sheetName = listNode.item(i).getNodeName().replaceAll("\r", "").replaceAll("\n", "").trim();
                if (sheetName.toLowerCase() != "#text") {
                    if (!sheetsNames.contains(sheetName)) {
                        sheetsNames.add(listNode.item(i).getNodeName());
                        listSheets.add(workbook.createSheet(sheetsNames.get(sheetsNames.size() - 1)));
                        XSSFRow row = listSheets.stream().filter(sheetl -> sheetl.getSheetName() == sheetName).findFirst().get().createRow(0);
                        System.out.println("Row number: " + row.getRowNum());
                        makeCells(row, listNode.item(i));
                    } else {
                        XSSFSheet sheet = listSheets.get(sheetsNames.indexOf(listNode.item(i).getNodeName()));
                        XSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
                        // System.out.println("Row number: "+row.getRowNum());
                        makeCells(row, listNode.item(i));

                    }
                }
            }
            BufferedOutputStream boos = new BufferedOutputStream(new FileOutputStream(xlsFile));
            workbook.write(boos);
            Runtime.getRuntime().exec("explorer.exe /select," + xlsFile.getPath());
        } catch (Exception ex) {
            Logger.getLogger(ParseConvertUtil.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void readXls(File xlsFile) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(xlsFile);
            //XSSFWorkbookFactory.create(new ObjectInputStream(new FileInputStream(xlsFile)));
            File xmlFile = new File(dir + "/" + xlsFile.getName().substring(0, xlsFile.getName().lastIndexOf(".") + 1) + "xml");
            FileWriter fwr = new FileWriter(xmlFile);

            String xmlDatas = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                    + "<AsycudaWorld_Manifest xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n";
            //xmlDatas+="<"+workbook.getn
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                XSSFRow firstRow = sheet.getRow(0);
                xmlDatas += "<" + sheet.getSheetName().replace(" ", "_") + ">\n";
                for (int j = 1; j < sheet.getLastRowNum(); j++) {
                    XSSFRow row = sheet.getRow(j);
                    xmlDatas += "\t<" + sheet.getSheetName().replace(" ", "_") + "_DATA>\n";
                    for (int k = 0; k < row.getLastCellNum(); k++) {
                        XSSFCell cell = row.getCell(k);
                        try {
                            cell.setCellType(CellType.STRING);
                        } catch (Exception e) {
                            Logger.getLogger(ParseConvertUtil.class.getName()).log(Level.SEVERE, null, e);
                        }
                        firstRow.getCell(k).setCellType(CellType.STRING);
                        xmlDatas += "\t\t<" + firstRow.getCell(k).getStringCellValue().replace(" ", "_") + ">";
                        try {
                            xmlDatas += cell.getStringCellValue();
                        } catch (Exception e) {
                            Logger.getLogger(ParseConvertUtil.class.getName()).log(Level.SEVERE, null, e);
                        }
                        xmlDatas += "</" + firstRow.getCell(k).getStringCellValue().replace(" ", "_") + ">\n";
                    }
                    xmlDatas += "\t</" + sheet.getSheetName().replace(" ", "_") + "_DATA>\n";
                }
                xmlDatas += "</" + sheet.getSheetName().replace(" ", "_") + ">";
            }
            xmlDatas += "</AsycudaWorld_Manifest>";
            fwr.write(xmlDatas);
            fwr.close();
            Runtime.getRuntime().exec("explorer.exe /select," + xmlFile.getPath());
        } catch (Exception ex) {
            Alert alert = new Alert(Alert.AlertType.ERROR, "Fichier mal structuré", ButtonType.CLOSE);
            Logger.getLogger(ParseConvertUtil.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void makeCells(XSSFRow row, Node node) {
        node.normalize();
        String nodeValue = node.getNodeValue() == null ? "" : node.getNodeValue().replaceAll("\r", "").replaceAll("\n", "").trim();
        if (!node.hasChildNodes() && nodeValue != "" && nodeValue != " " && nodeValue.length() > 0) {
            System.out.println("node value " + node.getNodeValue());
            XSSFCell cell = row.createCell(row.getLastCellNum() >= 0 ? row.getLastCellNum() : 0);
            cell.setCellValue(node.getNodeValue().replaceAll("\r", "").replaceAll("\n", "").trim());
            cell.setCellStyle(styleNormal);
        } else {
            for (int i = 0; i < node.getChildNodes().getLength(); i++) {
                makeCells(row, node.getChildNodes().item(i));
            }
        }
    }

    public void setColsNames(XSSFRow row, Node node) {
        node.normalize();
        String nodeValue = node.getNodeValue() == null ? "" : node.getNodeValue().replaceAll("\r", "").replaceAll("\n", "").trim();
        if (!node.hasChildNodes() && nodeValue != "" && nodeValue != " " && nodeValue.length() > 0) {
            XSSFCell cell = row.createCell(row.getLastCellNum() >= 0 ? row.getLastCellNum() : 0);
            cell.setCellValue(node.getParentNode().getNodeName().replaceAll("\r", "").replaceAll("\n", "").trim());
            cell.setCellStyle(styleHead);
        } else {
            for (int i = 0; i < node.getChildNodes().getLength(); i++) {
                setColsNames(row, node.getChildNodes().item(i));
            }
        }
    }
}
