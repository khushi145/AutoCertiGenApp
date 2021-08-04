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
    int name,college,course,position,society;
    int row_num;
    Button go_to;
    String path;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_template);

        path = getIntent().getStringExtra("path");

        test = (TextView) findViewById(R.id.test);
        go_to = findViewById( R.id.generate_btn );

        try {
            row_num = Integer.parseInt(getIntent().getStringExtra("entries"));
        }catch (NullPointerException e){
            e.printStackTrace();
        }
        readExcelData();
        test.setText("Certificate Generation Completed!");

        /*go_to.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Uri selectedUri = Uri.parse( getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS).toString() );
                Intent i = new Intent(Intent.ACTION_VIEW);
                i.setDataAndType(selectedUri, "resource/folder");

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
        } );
         */
    }

    public void readExcelData() {
        new Thread(() -> {
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
                identifyColumn();
                genPDF();
            } catch (FileNotFoundException e) {
                test.setText("FilenotFound");
                e.printStackTrace();
            } catch (IOException e) {
                test.setText("IOException");
                e.printStackTrace();
            }
        }).start();
    }

    public void identifyColumn(){
        for (int i=0; i<5; i++){
            if(exceldata[0][i].toLowerCase().equals("name")){
                name=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("college")){
                college=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("course")){
                course=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("position")){
                position=i;
            }
            else if(exceldata[0][i].toLowerCase().equals("society")){
                society=i;
            }
            else{
                //add toast here for error msg ("column name don't match template
                // add another excel sheet")
            }
        }
    }

    private void genPDF() {
        try {
            PDFBoxResourceLoader.init(getApplicationContext());
            AssetManager assManager = getAssets();

            for (int i=1; i<row_num+1; i++){
                InputStream is = assManager.open("template1.pdf");
                OutputStream newPDFfile = createFile(exceldata[i][name]);
                copy(is, newPDFfile);
                File PDFfile = new File(getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)+exceldata[i][name]+".pdf");
                PDDocument pdf = PDDocument.load(PDFfile);

                PDPage page = pdf.getPage(0);
                PDPageContentStream contentStream = new PDPageContentStream(pdf, page,true,false);
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_BOLD, 150);
                contentStream.newLineAtOffset(1400, -1200);
                contentStream.showText(exceldata[i][name]);
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 120);
                contentStream.newLineAtOffset(1150, -1400);
                contentStream.showText("from "+exceldata[i][college]+" has secured");
                contentStream.endText();
                contentStream.beginText();
                contentStream.setTextMatrix(new Matrix(1f, 0f, 0f, -1f, 0f, 0f));
                contentStream.setFont(PDType1Font.TIMES_ROMAN, 120);
                contentStream.newLineAtOffset(1150, -1520);
                contentStream.showText(exceldata[i][position]+" position in "+exceldata[i][society]+".");
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