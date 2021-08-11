package com.example.autocertigen;

import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.os.Bundle;
import android.text.TextUtils;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;

@SuppressWarnings("deprecation")
public class MainActivity extends AppCompatActivity {
    Button browse;
    EditText entries;
    String path_xlsx;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        browse = findViewById(R.id.browse);
        entries = findViewById(R.id.entries);

        browse.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if (TextUtils.isEmpty(entries.getText())) {
                    //Toast.makeText( getApplicationContext(), "Number of Certificates reqquired" ).show();
                    entries.setError(" Please Enter Number of Certificates ");
                    entries.requestFocus();
                } else {
                    getPath();
                }
            }
        } );
    }

    protected void getPath() {
        Intent i = new Intent(Intent.ACTION_GET_CONTENT);
        String[] mimeTypes = {"application/vnd.ms-excel" , "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"};
        i.setType("*/*");
        i.putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes);
        startActivityForResult(i, 1);
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        switch (requestCode) {
            case 1:
                if (resultCode == RESULT_OK) {
                    path_xlsx = data.getData().toString();
                    Intent i = new Intent(getApplicationContext(), ScrollSelectActivity.class );
                    i.putExtra("path", path_xlsx);
                    i.putExtra("entries",entries.getText().toString());
                    startActivity(i);
                }
                break;
        }
    }

}