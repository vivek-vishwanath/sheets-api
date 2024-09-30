package org.example

import com.google.api.client.auth.oauth2.Credential
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.json.gson.GsonFactory
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.SheetsScopes
import com.google.auth.http.HttpCredentialsAdapter
import com.google.auth.oauth2.GoogleCredentials
import java.io.FileNotFoundException
import java.io.InputStreamReader

object OAuth {

    // Located in src/main/resources
    const val CREDENTIALS_PATH = "credentials.json"
    const val TOKENS_PATH = "/tokens"
    const val APPLICATION_NAME = "Quiz Bowl Tournament ScoreSheet Maker"

    val JSON_FACTORY = GsonFactory.getDefaultInstance()
    val HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport()
    val SCOPES = listOf(SheetsScopes.SPREADSHEETS, SheetsScopes.DRIVE)
    // val style = Format(fontSize = 14, fontFamily = "Trebuchet MS")

    val service: Sheets by lazy {
        val httpTransport = GoogleNetHttpTransport.newTrustedTransport()
        val jsonFactory = JSON_FACTORY

        val stream = OAuth::class.java.classLoader.getResourceAsStream(CREDENTIALS_PATH)
            ?: throw FileNotFoundException("Resource not found at src/main/resources/$CREDENTIALS_PATH")
        val credential = GoogleCredentials.fromStream(stream)
            .createScoped(listOf("https://www.googleapis.com/auth/spreadsheets"))

        Sheets.Builder(httpTransport, jsonFactory, HttpCredentialsAdapter(credential))
            .setApplicationName(APPLICATION_NAME)
            .build()
    }

    val anchor = service.spreadsheets().values()
}