package com.example.autogeneratecertificates;

import android.graphics.Color;
import android.os.Bundle;
import android.view.View;

import androidx.appcompat.app.AppCompatActivity;

import gr.net.maroulis.library.EasySplashScreen;

public class SplashScreenActivity extends AppCompatActivity {
    static {
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLInputFactory",
                "com.fasterxml.aalto.stax.InputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLOutputFactory",
                "com.fasterxml.aalto.stax.OutputFactoryImpl"
        );
        System.setProperty(
                "org.apache.poi.javax.xml.stream.XMLEventFactory",
                "com.fasterxml.aalto.stax.EventFactoryImpl"
        );
    }
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate( savedInstanceState );

        EasySplashScreen config = new EasySplashScreen( SplashScreenActivity.this )
                .withFullScreen()
                .withTargetActivity( MainActivity.class )
                .withSplashTimeOut( 2000 )
                .withBackgroundResource(R.drawable.splash_launcher_background)
                .withHeaderText( "" )
                .withFooterText( "Copyright 2021" )
                .withBeforeLogoText( "" )
                .withAfterLogoText( "CertiGen" )
                .withLogo( R.mipmap.ic_launcher_round );

        config.getHeaderTextView().setTextColor( Color.BLACK );
        config.getAfterLogoTextView().setTextColor( Color.WHITE );
        config.getBeforeLogoTextView().setTextColor( Color.BLUE );
        config.getFooterTextView().setTextColor( Color.WHITE );

        View splashscreen = config.create();
        setContentView( splashscreen );
    }
}
