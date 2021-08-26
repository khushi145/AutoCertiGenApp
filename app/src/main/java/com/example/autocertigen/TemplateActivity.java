package com.example.autocertigen;

import android.annotation.SuppressLint;
import android.content.Intent;
import android.content.res.AssetManager;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;

import com.tom_roush.pdfbox.pdmodel.PDDocument;
import com.tom_roush.pdfbox.pdmodel.PDPage;
import com.tom_roush.pdfbox.pdmodel.PDPageContentStream;
import com.tom_roush.pdfbox.pdmodel.font.PDType1Font;
import com.tom_roush.pdfbox.pdmodel.graphics.image.LosslessFactory;
import com.tom_roush.pdfbox.pdmodel.graphics.image.PDImageXObject;
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
import java.util.NoSuchElementException;


public class TemplateActivity extends AppCompatActivity {

    TextView success,displayPath;
    String[][] exceldata = new String[2000][30];
    int name,course,position,society,competition,date,year;
    int row_num;
    Button go_to, goMain;
    boolean matchFlag,sizeFlag;
    String path,template,signatory1,signatory2,designation1,designation2,sign1Image,sign2Image;

    @SuppressLint("SetTextI18n")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_template);

        success = (TextView) findViewById(R.id.success);
        displayPath=(TextView)findViewById(R.id.pdfLoc);
        go_to = (Button) findViewById( R.id.goto_btn );
        goMain = (Button) findViewById( R.id.main_button );
        matchFlag=sizeFlag=true;

        new Thread(()->{
            path = getIntent().getStringExtra("path");
            template = getIntent().getStringExtra("template");
            signatory1 = getIntent().getStringExtra("signatory1");
            signatory2 = getIntent().getStringExtra("signatory2");
            designation1 = getIntent().getStringExtra("designation1");
            designation2 = getIntent().getStringExtra("designation2");
            sign1Image=getIntent().getStringExtra("sign1image");
            sign2Image=getIntent().getStringExtra("sign2image");

            Log.d("TAG-TEMPLATE","path1: "+sign1Image);
            Log.d("TAG-TEMPLATE","path2: "+sign2Image);

            try {
                row_num = Integer.parseInt(getIntent().getStringExtra("entries"));
            } catch (NullPointerException e) {
                e.printStackTrace();
            }

            switch (template) {
                case "t1":
                    readExcelData1();
                    break;
                case "t2":
                    readExcelData2();
                    break;
                default:
                    Toast.makeText(getApplicationContext(), "Error in template selection", Toast.LENGTH_SHORT).show();
            }
        }).start();

        go_to.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Uri selectedUri = Uri.parse(getFolder());
                Intent i = new Intent(Intent.ACTION_VIEW);
                i.setDataAndType(selectedUri,"*/*");

                if (i.resolveActivityInfo(getPackageManager(),0)!=null)
                {
                    startActivity(Intent.createChooser(i,"Choose"));
                }
                else
                {
                    Toast.makeText( getApplicationContext(), "Couldn't reach to the Destination", Toast.LENGTH_SHORT ).show();
                    // if you reach this place, it means there is no any file
                    // explorer app installed on your device
                }
            }
        } );

        goMain.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent iMain = new Intent(getApplicationContext(), MainActivity.class );
                iMain.addFlags(Intent.FLAG_ACTIVITY_CLEAR_TOP | Intent.FLAG_ACTIVITY_NEW_TASK);
                startActivity(iMain);
                TemplateActivity.this.finish();
            }
        });
    }

    @SuppressLint("SetTextI18n")
    public void readExcelData1() {
       try {
           InputStream inputfile = getContentResolver().openInputStream(Uri.parse(path));
           XSSFWorkbook workbook = new XSSFWorkbook(inputfile);
           XSSFSheet sheet = workbook.getSheetAt(0);
           Iterator<Row> rowIterator = sheet.iterator();
           int count = 0;
           String temp;
           try {
               while (rowIterator.hasNext() && count < row_num + 1) {
                   Row row = rowIterator.next();
                   Iterator<Cell> cx = row.cellIterator();
                   for (int i = 0; i < 5; i++) {
                       temp = cx.next().getStringCellValue();
                       exceldata[count][i] = temp;
                   }
                   count++;
               }
           }catch(NoSuchElementException e){
               e.printStackTrace();
               sizeFlag=false;
               success.setText("Please upload another excel file.");
               displayPath.setText("The number of entries in Excel file is less than the number of certificates to be generated.");
           }
           inputfile.close();
           if(sizeFlag) {
               identifyColumn1();
               if (count < row_num){
                   success.setText("Please upload another excel file.");
                   displayPath.setText("The number of entries in Excel file is less than the number of certificates to be generated.");
                   throw new IOException();
               }
               if (matchFlag) {
                   genPDF1();
               } else {
                   success.setText("Please upload another excel file.");
                   displayPath.setText("Columns of the excelsheet don't match the Template placeholders.");
               }
           }
       } catch (FileNotFoundException e) {
           e.printStackTrace();
       } catch (IOException e) {
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
                matchFlag=false;
            }
        }
    }

    @SuppressLint("SetTextI18n")
    private void genPDF1() {
        try {
            PDFBoxResourceLoader.init(getApplicationContext());
            AssetManager assetManager = getAssets();
            for (int i=1; i<row_num+1; i++){
                success.setText(i-1+"/"+row_num+" FILES GENERATED");
                InputStream is = assetManager.open("t1.pdf");
                OutputStream newPDFfile = createFile(exceldata[i][name],i);
                copy(is, newPDFfile);
                File PDFfile = new File(getFolder()+i+"_"+exceldata[i][name]+".pdf");
                Log.d("PATH",PDFfile.getAbsolutePath());
                Log.d("PATH",PDFfile.getPath());
                PDDocument pdf = PDDocument.load(PDFfile);
                PDPage page = pdf.getPage(0);
                PDPageContentStream contentStream = new PDPageContentStream(pdf, page,true,true);
                contentStream.transform(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_BOLD, 150);
                contentStream.newLineAtOffset(1330, -1300);
                contentStream.showText(exceldata[i][name]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(600, -1500);
                contentStream.showText("for securing "+exceldata[i][position]+" place in the competition-"+exceldata[i][competition]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(1000, -1620);
                contentStream.showText("organized by the society: "+exceldata[i][society]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(550, -1740);
                contentStream.showText("under the aegis of the Annual Cultural Festival -Karvaan'21-");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(1300, -1860);
                contentStream.showText("held on "+exceldata[i][date]+".");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2200);
                contentStream.showText(signatory1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2380, -2200);
                contentStream.showText(signatory2);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2300);
                contentStream.showText(designation1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2380, -2300);
                contentStream.showText(designation2);
                contentStream.endText();
                if(!sign1Image.equals("NULL")) {
                    InputStream ims1 = getContentResolver().openInputStream(Uri.parse(sign1Image));
                    Bitmap original = BitmapFactory.decodeStream(ims1);
                    Bitmap scaled = Bitmap.createScaledBitmap(original,450,250, false);
                    PDImageXObject pdImage1= LosslessFactory.createFromImage(pdf,scaled);
                    contentStream.drawImage(pdImage1,400,-2100,450,250);
                    ims1.close();
                }
                if(!sign2Image.equals("NULL")) {
                    InputStream ims2 = getContentResolver().openInputStream(Uri.parse(sign2Image));
                    Bitmap original = BitmapFactory.decodeStream(ims2);
                    Bitmap scaled = Bitmap.createScaledBitmap(original,450,250, false);
                    PDImageXObject pdImage2= LosslessFactory.createFromImage(pdf,scaled);
                    contentStream.drawImage(pdImage2, 2450, -2100, 450, 250);
                    ims2.close();
                }
                contentStream.close();
                pdf.save(PDFfile);
                pdf.close();
                newPDFfile.close();
                is.close();
            }
            success.setText("Certificate Generation Completed!");
            displayPath.setText("The files are located at the following location in Internal Storage:\n" +
                    getFolder());
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    @SuppressLint("SetTextI18n")
    public void readExcelData2() {
        try {
            InputStream inputfile = getContentResolver().openInputStream(Uri.parse(path));
            XSSFWorkbook workbook = new XSSFWorkbook(inputfile);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            int count = 0;
            String temp;
            try {
                while (rowIterator.hasNext() && count < row_num + 1) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cx = row.cellIterator();
                    for (int i = 0; i < 5; i++) {
                        temp = cx.next().getStringCellValue();
                        exceldata[count][i] = temp;
                    }
                    count++;
                }
            }catch(NoSuchElementException e){
                e.printStackTrace();
                sizeFlag=false;
                success.setText("Please upload another excel file.");
                displayPath.setText("The number of entries in Excel file is less than the number of certificates to be generated.");
            }
            inputfile.close();
            if(sizeFlag) {
                identifyColumn2();
                if (matchFlag) {
                    genPDF2();
                } else {
                    success.setText("Please upload another excel file.");
                    displayPath.setText("Columns of the excelsheet don't match the Template placeholders.");
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
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
                matchFlag=false;
            }
        }
    }

    @SuppressLint("SetTextI18n")
    private void genPDF2() {
        try {
            PDFBoxResourceLoader.init(getApplicationContext());
            AssetManager assManager = getAssets();

            for (int i=1; i<row_num+1; i++){
                success.setText(i-1+"/"+row_num+" FILES GENERATED");
                InputStream is = assManager.open("t2.pdf");
                OutputStream newPDFfile = createFile(exceldata[i][name],i);
                copy(is, newPDFfile);
                File PDFfile = new File(getFolder()+i+"_"+exceldata[i][name]+".pdf");
                PDDocument pdf = PDDocument.load(PDFfile);

                PDPage page = pdf.getPage(0);
                PDPageContentStream contentStream = new PDPageContentStream(pdf, page,true,false);
                contentStream.transform(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_BOLD, 100);
                contentStream.newLineAtOffset(900, -1300);
                contentStream.showText("This is to certify that Ms. "+exceldata[i][name]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(900, -1420);
                contentStream.showText("of "+exceldata[i][course]+" "+exceldata[i][year]+" year has secured");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(900, -1540);
                contentStream.showText(exceldata[i][position]+" position in the academic session 2020-2021.");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 100);
                contentStream.newLineAtOffset(1200, -1760);
                contentStream.showText("Presented this on: "+exceldata[i][date]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2200);
                contentStream.showText(signatory1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2500, -2200);
                contentStream.showText(signatory2);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(400, -2300);
                contentStream.showText(designation1);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 90);
                contentStream.newLineAtOffset(2500, -2300);
                contentStream.showText(designation2);
                contentStream.endText();

                if(!sign1Image.equals("NULL")) {
                    InputStream ims1 = getContentResolver().openInputStream(Uri.parse(sign1Image));
                    PDImageXObject pdImage1= LosslessFactory.createFromImage(pdf,BitmapFactory.decodeStream(ims1));
                    contentStream.drawImage(pdImage1,450,-2100,450,250);
                    ims1.close();
                }
                if(!sign2Image.equals("NULL")) {
                    InputStream ims2 = getContentResolver().openInputStream(Uri.parse(sign2Image));
                    PDImageXObject pdImage2= LosslessFactory.createFromImage(pdf,BitmapFactory.decodeStream(ims2));
                    contentStream.drawImage(pdImage2, 2450, -2100, 450, 250);
                    ims2.close();
                }

                contentStream.close();
                pdf.save(PDFfile);
                pdf.close();
                newPDFfile.close();
                is.close();
            }
            success.setText("Certificate Generation Completed!");
            displayPath.setText("The files are located at the following location in Internal Storage:\n" +
                    getFolder());
        }catch (IOException e) {
            e.printStackTrace();
        }
    }

    private OutputStream createFile(String name,int i) throws IOException {
        File f = new File(getFolder() + i + "_" + name + ".pdf");
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

    @NonNull
    private String getFolder() {
        return Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS) + "/AutoCertiGen/";
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