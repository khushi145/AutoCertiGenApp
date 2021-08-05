package com.example.autogeneratecertificates;

import android.content.Intent;
import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.appcompat.app.AppCompatActivity;

import com.tom_roush.pdfbox.pdmodel.PDDocument;
import com.tom_roush.pdfbox.pdmodel.PDPage;
import com.tom_roush.pdfbox.pdmodel.PDPageContentStream;
import com.tom_roush.pdfbox.pdmodel.font.PDType1Font;
import com.tom_roush.pdfbox.util.Matrix;
import com.tom_roush.pdfbox.util.PDFBoxResourceLoader;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;


public class TemplateActivity extends AppCompatActivity {

    TextView test;
    String[][] exceldata = new String[30][30];
    int name,college,course,position,society,competition,date,year;
    int row_num;
    Button go_to;
    String path,template,signatory1,signatory2,designation1,designation2;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_template);

        path = getIntent().getStringExtra("path");
        template=getIntent().getStringExtra("template");
        signatory1=getIntent().getStringExtra("signatory1");
        signatory2=getIntent().getStringExtra("signatory2");
        designation1=getIntent().getStringExtra("designation1");
        designation2=getIntent().getStringExtra("designation2");

        try {
            row_num = Integer.parseInt(getIntent().getStringExtra("entries"));
        }catch (NullPointerException e){
            e.printStackTrace();
        }

        switch(template){
            case "t1":
                readExcelData1();
                break;
            case "t2":
                readExcelData2();
                break;
            default:Toast.makeText(getApplicationContext(),"Error in template selection",Toast.LENGTH_SHORT).show();
        }

        test = (TextView) findViewById(R.id.test);
        go_to = findViewById( R.id.generate_btn );
        /*
        go_to.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Uri selectedUri = Uri.parse("file://"+Environment.DIRECTORY_DOWNLOADS+"/AutoCertiGen/");
                Intent i = new Intent(Intent.ACTION_VIEW);
                i.setDataAndType(selectedUri, "application/*");

                if (i.resolveActivityInfo(getPackageManager(), 0) != null)
                {
                    startActivity(i);
                }
                else
                {
                    Toast.makeText( getApplicationContext(), "Couldn't reach to the Destination", Toast.LENGTH_SHORT ).show();
                    // if you reach this place, it means there is no any file
                    // explorer app installed on your device
                }
            }
        } );*/
    }

    public void readExcelData1() {
        try {
            InputStream inputfile = getContentResolver().openInputStream(Uri.parse(path));
            XSSFWorkbook workbook = new XSSFWorkbook(inputfile);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            int count = 0;
            String temp;
            while (rowIterator.hasNext() && count < row_num+1) {
                Row row = rowIterator.next();
                Iterator<Cell> cx = row.cellIterator();
                for (int i=0; i<5; i++){
                    temp=cx.next().getStringCellValue();
                    exceldata[count][i]=temp;
                }
                count++;
            }
            inputfile.close();
            identifyColumn1();
            genPDF1();
        } catch (FileNotFoundException e) {
            test.setText("FilenotFound");
            e.printStackTrace();
        } catch (IOException e) {
            test.setText("IOException");
            e.printStackTrace();
        }
    }

    public void identifyColumn1(){
        for (int i=0; i<5; i++){
            if(exceldata[0][i].toLowerCase().equals("name")){
                name=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("date")){
                date=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("competition")){
                competition=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("position")){
                position=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("society")){
                society=i;
            }
            else{
                Toast.makeText(getApplicationContext(),"Columns of the excelsheet don't match the Template placeholders. Please upload another excel file",Toast.LENGTH_LONG).show();
            }
        }
    }

    private void genPDF1() {
        try {
            PDFBoxResourceLoader.init(getApplicationContext());
            AssetManager assetManager = getAssets();
            for (int i=1; i<row_num+1; i++){
                InputStream is = assetManager.open("t1.pdf");
                OutputStream newPDFfile = createFile(exceldata[i][name]);
                copy(is, newPDFfile);
                File PDFfile = new File(getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)+"/AutoCertiGen/"+exceldata[i][name]+".pdf");
                PDDocument pdf = PDDocument.load(PDFfile);
                PDPage page = pdf.getPage(0);
                PDPageContentStream contentStream = new PDPageContentStream(pdf, page,true,false);
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_BOLD, 150);
                contentStream.newLineAtOffset(1330, -1300);
                contentStream.showText(exceldata[i][name]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(600, -1500);
                contentStream.showText("for securing "+exceldata[i][position]+" place in the competition-"+exceldata[i][competition]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(1000, -1620);
                contentStream.showText("organized by the society: "+exceldata[i][society]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(550, -1740);
                contentStream.showText("under the aegis of the Annual Cultural Festival -Karvaan'21-");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(1300, -1860);
                contentStream.showText("held on "+exceldata[i][date]+".");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2200);
                contentStream.showText(signatory1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2500, -2200);
                contentStream.showText(signatory2);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2300);
                contentStream.showText(designation1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2500, -2300);
                contentStream.showText(designation2);
                contentStream.endText();
                contentStream.close();
                pdf.save(PDFfile);
                pdf.close();
                newPDFfile.close();
                is.close();
            }
            test.setText("Certificate Generation Completed!");
            Log.d("TemplateActivity", "Completed");
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void readExcelData2() {
        try {
            InputStream inputfile = getContentResolver().openInputStream(Uri.parse(path));
            XSSFWorkbook workbook = new XSSFWorkbook(inputfile);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            int count = 0;
            String temp;
            while (rowIterator.hasNext() && count < row_num+1) {
                Row row = rowIterator.next();
                Iterator<Cell> cx = row.cellIterator();
                for (int i=0; i<5; i++){
                    temp=cx.next().getStringCellValue();
                    exceldata[count][i]=temp;
                }
                count++;
            }

            inputfile.close();
            identifyColumn2();
            genPDF2();
        } catch (FileNotFoundException e) {
            test.setText("FilenotFound");
            e.printStackTrace();
        } catch (IOException e) {
            test.setText("IOException");
            e.printStackTrace();
        }
    }

    public void identifyColumn2(){
        for (int i=0; i<5; i++){
            if(exceldata[0][i].toLowerCase().equals("name")){
                name=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("year")){
                year=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("course")){
                course=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("position")){
                position=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("date")){
                date=i;
            }
            else{
                //add toast here for error msg ("column name don't match template
                // add another excel sheet")
            }
        }
    }

    private void genPDF2() {
        try {
            PDFBoxResourceLoader.init(getApplicationContext());
            AssetManager assManager = getAssets();

            for (int i=1; i<row_num+1; i++){
                InputStream is = assManager.open("t2.pdf");
                OutputStream newPDFfile = createFile(exceldata[i][name]);
                copy(is, newPDFfile);
                File PDFfile = new File(getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)+"/AutoCertiGen/"+exceldata[i][name]+".pdf");
                PDDocument pdf = PDDocument.load(PDFfile);

                PDPage page = pdf.getPage(0);
                PDPageContentStream contentStream = new PDPageContentStream(pdf, page,true,false);
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_BOLD, 100);
                contentStream.newLineAtOffset(900, -1300);
                contentStream.showText("This is to certify that Ms. "+exceldata[i][name]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(900, -1420);
                contentStream.showText("of "+exceldata[i][course]+" "+exceldata[i][year]+" year has secured");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(900, -1540);
                contentStream.showText(exceldata[i][position]+" position in the academic session 2020-2021.");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(1200, -1760);
                contentStream.showText("Presented this on: "+exceldata[i][date]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2200);
                contentStream.showText(signatory1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2500, -2200);
                contentStream.showText(signatory2);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2300);
                contentStream.showText(designation1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2500, -2300);
                contentStream.showText(designation2);
                contentStream.endText();
                contentStream.close();
                pdf.save(PDFfile);
                pdf.close();
                newPDFfile.close();
                is.close();
            }

            Log.d("TemplateActivity", "Completed");
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    private OutputStream createFile(String name) throws IOException {
        File f = new File(getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)+"/AutoCertiGen/" +name+".pdf");
        if (f.exists()){
            f.delete();
        }
        if (!f.getParentFile().exists()){
            f.getParentFile().mkdirs();
        }
        f.createNewFile();
        OutputStream newFile = new FileOutputStream(f);
        return newFile;
    }

    public static void copy(InputStream fis, OutputStream fos)
            throws IOException
    {
        byte[] buffer = new byte[1024];
        int read;
        while((read = fis.read(buffer)) != -1){
            fos.write(buffer, 0, read);
        }
    }
}