import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class SheetsAndJava {
    private static Sheets sheetsService;
    private static String SPREADSHEET_ID = "1hMUg5xKtRNV_F48RgiSbENH4gwcAf6QqBz9niG-oOcs";

    private static Credential authorize() throws IOException, GeneralSecurityException {
        InputStream in = SheetsAndJava.class.getResourceAsStream("/credentials.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(
                JacksonFactory.getDefaultInstance(), new InputStreamReader(in)
        );

        List<String> scopes = Arrays.asList(SheetsScopes.SPREADSHEETS);

        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                GoogleNetHttpTransport.newTrustedTransport(), JacksonFactory.getDefaultInstance(),
                clientSecrets, scopes)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File("token")))
                .setAccessType("offline")
                .build();

        return new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver())
                .authorize("user");
    }

    private static Sheets getSheetsService() throws IOException, GeneralSecurityException {
        Credential credential = authorize();
        String APPLICATION_NAME = "Google Sheets Example";
        return new Sheets.Builder(GoogleNetHttpTransport.newTrustedTransport(),
                JacksonFactory.getDefaultInstance(), credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    private static void getValues(String range) throws IOException {
        ValueRange response = sheetsService.spreadsheets().values()
                .get(SPREADSHEET_ID, range)
                .execute();

        List<List<Object>> values = response.getValues();

        if(values == null || values.isEmpty()) {
            System.out.println("No data found.");
        } else {
            for (List row : values) {
                System.out.printf("%s | %s | %s\n", row.get(0), row.get(1), row.get(2));
            }
        }
    }

    public static void setValues(List<List<Object>> values) throws IOException {
        ValueRange appendBody = new ValueRange()
                .setValues(values);

        AppendValuesResponse appendResult = sheetsService.spreadsheets().values()
                .append(SPREADSHEET_ID, "sheet", appendBody)
                .setValueInputOption("USER_ENTERED")
                .setIncludeValuesInResponse(true)
                .execute();

        System.out.println(appendResult);
    }

    public static void updateValue(List<List<Object>> values, String range) throws IOException {
        ValueRange body = new ValueRange()
                .setValues(values);

        UpdateValuesResponse result = sheetsService.spreadsheets().values()
                .update(SPREADSHEET_ID, range, body)
                .setValueInputOption("RAW")
                .execute();

        System.out.println(result);
    }

    public static void deleteRows(int sheetId, int startIndex) throws IOException {
        DeleteDimensionRequest deleteRequest = new DeleteDimensionRequest()
                .setRange(
                        new DimensionRange()
                                .setSheetId(sheetId)
                                .setDimension("ROWS")
                                .setStartIndex(startIndex)
                );
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setDeleteDimension(deleteRequest));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, body).execute();
    }

    private static boolean hasWorkSheet(String title) throws IOException {
        Sheets.Spreadsheets.Get request = sheetsService.spreadsheets().get(SPREADSHEET_ID);
        Spreadsheet response = request.execute();
        List<Sheet> filteredSheets = response.getSheets().stream()
                .filter(sheet -> sheet.getProperties().getTitle().equals(title)).collect(Collectors.toList());

        if (filteredSheets.isEmpty()) {
            return false;
        }

        return filteredSheets.get(0) != null;
    }

    private static void updateWorksheetDimension(String worksheetTitle, int worksheetId, int row, int col) throws IOException {
        UpdateSheetPropertiesRequest updateSheetPropertiesRequest = new UpdateSheetPropertiesRequest()
                .setFields("*")
                .setProperties(new SheetProperties().setTitle(worksheetTitle).setSheetId(worksheetId).setGridProperties(new GridProperties().setColumnCount(col).setRowCount(row)));

        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setUpdateSheetProperties(updateSheetPropertiesRequest));

        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        sheetsService.spreadsheets().batchUpdate(SPREADSHEET_ID, body).execute();
    }

    public static void main(String[] args) {
        try {
            sheetsService = getSheetsService();

             String range = "sheet!A1:C3";
             getValues(range);

            List<List<Object>> valuesToAdd = Arrays.asList(
                    Arrays.asList("This", "was", "added", "from", "code!")
            );
            setValues(valuesToAdd);

            List<List<Object>> valuesToUpdate = Arrays.asList(
                    Arrays.asList("updated")
            );
            updateValue(valuesToUpdate, "A1");

            deleteRows(0, 3);

            hasWorkSheet("OI");

            updateWorksheetDimension("OI", 123, 2, 2);
        } catch (IOException | GeneralSecurityException e) {
            System.out.println("EXCEPTION:" + e.getMessage());
        }
    }
}
