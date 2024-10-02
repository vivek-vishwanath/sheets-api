package org.example

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.json.gson.GsonFactory
import com.google.api.services.sheets.v4.Sheets
import com.google.auth.http.HttpCredentialsAdapter
import com.google.auth.oauth2.GoogleCredentials
import java.io.FileNotFoundException

object OAuth {

    // Located in src/main/resources
    const val CREDENTIALS_PATH = "credentials.json"
    const val APPLICATION_NAME = "QB Tournament ScoreSheets"

    val service: Sheets by lazy {
        val httpTransport = GoogleNetHttpTransport.newTrustedTransport()
        val jsonFactory = GsonFactory.getDefaultInstance()

        val stream = OAuth::class.java.classLoader.getResourceAsStream(CREDENTIALS_PATH)
            ?: throw FileNotFoundException("Resource not found at src/main/resources/$CREDENTIALS_PATH")
        val credential = GoogleCredentials.fromStream(stream)
            .createScoped(listOf("https://www.googleapis.com/auth/spreadsheets"))

        Sheets.Builder(httpTransport, jsonFactory, HttpCredentialsAdapter(credential))
            .setApplicationName(APPLICATION_NAME)
            .build()
    }
}