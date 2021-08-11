package com.example.autocertigen;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.ImageView;

public class ScrollSelectActivity extends AppCompatActivity {
    ImageView t1, t2, t3;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate( savedInstanceState );
        setContentView( R.layout.activity_scroll_select );

        t1 = findViewById( R.id.t1_img );
        t2 = findViewById( R.id.t2_img );
        //t3 = findViewById( R.id.t3_img );

        t1.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                callIntent("t1");
            }
         });

        t2.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                callIntent("t2");
            }
        } );
/*
        t3.setOnClickListener( new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                callIntent("t3");
            }
        });
*/
    }

    protected void callIntent(String template){
        Intent i = new Intent(this, AdditionalInfo.class);
        i.putExtra("template", template);
        i.putExtra("path", getIntent().getStringExtra( "path" ));
        i.putExtra( "entries", getIntent().getStringExtra( "entries" ));
        startActivity( i );
    }
}