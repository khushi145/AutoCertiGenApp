package com.example.autogeneratecertificates;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.os.Bundle;
import android.text.TextUtils;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

public class AdditionalInfo extends AppCompatActivity {
    EditText signatory1, signatory2, designation1, designation2;
    Button generate;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate( savedInstanceState );
        setContentView( R.layout.activity_additional_info );
        signatory1 = findViewById( R.id.edit_signatory1 );
        designation1 = findViewById( R.id.edit_designation1 );
        signatory2 = findViewById( R.id.edit_signatory2 );
        designation2 = findViewById( R.id.edit_designation2 );
        generate = (Button)findViewById( R.id.generate_btn );

        generate.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                if(TextUtils.isEmpty(signatory1.getText()) && TextUtils.isEmpty(signatory2.getText()) && TextUtils.isEmpty(designation1.getText()) && TextUtils.isEmpty(designation2.getText()))
                {
                    Toast.makeText( getApplicationContext(), "Fields are Empty!",  Toast.LENGTH_SHORT).show();

                } else if (TextUtils.isEmpty(signatory1.getText()) ) {
                    //proceed with operation
                    signatory1.setError(" Please Enter Signatory Name ");
                    signatory1.requestFocus();
                } else if (TextUtils.isEmpty(designation1.getText())) {
                    designation1.setError( "Please Enter Designation Name " );
                    designation1.requestFocus();
                } else if (TextUtils.isEmpty(signatory2.getText())) {
                    signatory2.setError( "Please Enter Signatory Name " );
                    signatory2.requestFocus();
                } else if (TextUtils.isEmpty(designation2.getText())) {
                    designation2.setError( "Please Enter Designation Name " );
                    designation2.requestFocus();
                }
                else if(isInvalidInput()){
                    Toast.makeText( getApplicationContext(), "Re-enter the fields!",  Toast.LENGTH_SHORT).show();
                }
                else {
                    Intent i = new Intent(getApplicationContext(), TemplateActivity.class);
                    i.putExtra( "path", getIntent().getStringExtra( "path" ) );
                    i.putExtra( "entries", getIntent().getStringExtra( "entries" ) );
                    i.putExtra( "template", getIntent().getStringExtra( "template" ) );
                    i.putExtra( "signatory1", signatory1.getText().toString() );
                    i.putExtra( "designation1", designation1.getText().toString() );
                    i.putExtra( "signatory2", signatory2.getText().toString() );
                    i.putExtra( "designation2", designation2.getText().toString() );
                    startActivity( i );
                }
            }
        } );

    }
    public boolean isInvalidInput(){
        boolean flag=false;
        if (TextUtils.isDigitsOnly( signatory1.getText())) {
            signatory1.setError("Invalid Input: Input can't contain digits only");
            flag=true;
        }
        if (TextUtils.isDigitsOnly( designation1.getText() )) {
            signatory2.setError("Invalid Input: Input can't contain digits only");
            flag=true;
        }
        if (TextUtils.isDigitsOnly( signatory2.getText() )) {
            designation1.setError("Invalid Input: Input can't contain digits only");
            flag=true;
        }
        if (TextUtils.isDigitsOnly( designation2.getText() )) {
            designation2.setError("Invalid Input: Input can't contain digits only");
            flag=true;
        }
        return flag;
    }
}