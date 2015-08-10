package com.amannmalik.connector.googlesheets;

/**
 * Created by amann.malik on 8/7/2015.
 */

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpResponse;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.Arrays;

/**
 * @author amann.malik
 */

public class GoogleDriveConnector {

    private final Drive driveService;
    private static final Logger LOG = LoggerFactory.getLogger(GoogleDriveConnector.class);

    public GoogleDriveConnector(String serviceAccountEmail, File keyFile) throws GeneralSecurityException, IOException {
        HttpTransport transport = GoogleNetHttpTransport.newTrustedTransport();
        JacksonFactory factory = new JacksonFactory();
        GoogleCredential credential = new GoogleCredential.Builder()
                .setTransport(transport)
                .setJsonFactory(factory)
                .setServiceAccountId(serviceAccountEmail)
                .setServiceAccountScopes(Arrays.asList(DriveScopes.DRIVE))
                .setServiceAccountPrivateKeyFromP12File(keyFile)
                .build();
        driveService = new Drive.Builder(transport, factory, credential).build();
    }

    public File getWorkbook(String fileId) throws IOException {
        com.google.api.services.drive.model.File file = driveService.files().get(fileId).execute();
        String downloadUrl = file.getExportLinks().get("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        GenericUrl url = new GenericUrl(downloadUrl);

        HttpResponse resp = driveService.getRequestFactory().buildGetRequest(url).execute();
        byte[] buffer;
        try (InputStream inStream = resp.getContent()) {
            buffer = new byte[inStream.available()];
            inStream.read(buffer);
        }
        resp.disconnect();

        final java.io.File tempFile = java.io.File.createTempFile(fileId, ".xlsx");
        try (OutputStream outStream = new FileOutputStream(tempFile)) {
            outStream.write(buffer);
        }
        return tempFile;
    }
}
