package org.example;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.net.URL;
import java.nio.file.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DownloadPDFsUI extends JFrame {

    private JTextField filePathField;
    private JButton selectFileButton;
    private JButton downloadButton;

    // Dane bazy danych PostgreSQL
    private static final String DB_URL = "jdbc:postgresql://10.100.100.42:5432/hrappka";
    private static final String USER = "hrappka";
    private static final String PASS = "1UjJ7DIHXO3YpePh";

    public DownloadPDFsUI() {
        super("Pobieranie certyfikatów PDF");

        // Inicjalizacja interfejsu użytkownika
        filePathField = new JTextField(30);
        selectFileButton = new JButton("Wybierz plik Excel");
        downloadButton = new JButton("Pobierz certyfikaty");

        // Ustawienie layoutu
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        panel.add(filePathField);
        panel.add(selectFileButton);
        panel.add(downloadButton);
        add(panel);

        // Ustawienie działania przycisków
        selectFileButton.addActionListener(e -> chooseExcelFile());
        downloadButton.addActionListener(e -> downloadCertificates());

        // Konfiguracja okna
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(400, 150);
        setLocationRelativeTo(null); // Wyśrodkowanie okna
        setVisible(true);
    }

    private void chooseExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        int option = fileChooser.showOpenDialog(DownloadPDFsUI.this);
        if (option == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            filePathField.setText(selectedFile.getAbsolutePath());
        }
    }

    private void downloadCertificates() {
        String excelFilePath = filePathField.getText().trim();
        if (excelFilePath.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Wybierz plik Excel przed pobraniem certyfikatów.");
            return;
        }

        // Wczytaj numery certyfikatów z pliku Excel
        List<String> certificateNumbers = readCertificateNumbersFromExcel(excelFilePath);
        if (certificateNumbers.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Brak numerów certyfikatów do pobrania.");
            return;
        }

        String excelFileName = getFileNameWithoutExtension(excelFilePath);

        // Pobierz certyfikaty PDF z bazy danych i zapisz lokalnie
        try (Connection conn = DriverManager.getConnection(DB_URL, USER, PASS)) {
            String query = buildQuery(certificateNumbers.size());
            try (PreparedStatement pstmt = conn.prepareStatement(query)) {
                for (int i = 0; i < certificateNumbers.size(); i++) {
                    pstmt.setString(i + 1, certificateNumbers.get(i));
                }

                try (ResultSet rs = pstmt.executeQuery()) {
                    while (rs.next()) {
                        String usrName = rs.getString("usr_name");
                        String fileUrl = rs.getString("file_url");
                        String certNumber = rs.getString("cert_number");

                        // Utwórz folder dla użytkownika, jeśli nie istnieje
                        Path userDir = Paths.get( excelFileName, usrName);
                        if (!Files.exists(userDir)) {
                            Files.createDirectories(userDir);
                        }

                        // Pobierz plik PDF
                        try (InputStream in = new URL(fileUrl).openStream()) {
                            Path filePath = userDir.resolve(certNumber + ".pdf");
                            Files.copy(in, filePath, StandardCopyOption.REPLACE_EXISTING);
                            System.out.println("Pobrano i zapisano plik: " + filePath);
                        } catch (IOException ex) {
                            System.out.println("Nie udało się pobrać pliku: " + fileUrl);
                            ex.printStackTrace(); // Opcjonalnie wyświetl szczegóły wyjątku
                        }
                    }

                    JOptionPane.showMessageDialog(this, "Pobrano certyfikaty z bazy danych.");
                } catch (SQLException ex) {
                    System.out.println("Błąd wykonania zapytania SQL: " + ex.getMessage());
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }
        } catch (SQLException ex) {
            System.out.println("Błąd połączenia z bazą danych: " + ex.getMessage());
        }
    }

    private String getFileNameWithoutExtension(String filePath) {
        String fileName = Paths.get(filePath).getFileName().toString();
        int dotIndex = fileName.lastIndexOf('.');
        if (dotIndex != -1) {
            return fileName.substring(0, dotIndex);
        }
        return fileName;
    }

    private List<String> readCertificateNumbersFromExcel(String excelFilePath) {
        List<String> certificateNumbers = new ArrayList<>();
        try (InputStream inputStream = new FileInputStream(excelFilePath)) {
            Workbook workbook;
            if (excelFilePath.toLowerCase().endsWith(".xls")) {
                workbook = new HSSFWorkbook(inputStream); // Obsługa plików Excel .xls
            } else if (excelFilePath.toLowerCase().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(inputStream); // Obsługa plików Excel .xlsx
            } else {
                JOptionPane.showMessageDialog(this, "Nieobsługiwany typ pliku Excel: " + excelFilePath);
                return certificateNumbers;
            }

            Sheet sheet = workbook.getSheetAt(0); // Zakładam, że dane są na pierwszym arkuszu
            for (Row row : sheet) {
                Cell cell = row.getCell(12); // Numer certyfikatu w kolumnie M (indeks 12)
                if (cell != null) {
                    certificateNumbers.add(cell.getStringCellValue().trim());
                }
            }
        } catch (IOException | EncryptedDocumentException e) {
            JOptionPane.showMessageDialog(this, "Błąd wczytywania pliku Excel: " + e.getMessage());
        }
        return certificateNumbers;
    }

    private String buildQuery(int numberOfParams) {
        StringBuilder queryBuilder = new StringBuilder();
        queryBuilder.append("SELECT ")
                .append("usr_name, ")
                .append("cr_number, ")
                .append("cert_number, ")
                .append("file_url ")
                .append("FROM ( ")
                .append("SELECT ")
                .append("u.usr_name, ")
                .append("request_event.cr_number, ")
                .append("uc.cert_number, ")
                .append("CONCAT('https://hrappka.budhrd.eu/files/get/f_hash/', f.f_hash, '/h/', c.cmp_hash) AS file_url, ")
                .append("ROW_NUMBER() OVER(PARTITION BY u.usr_name, uc.cert_number ORDER BY f.f_hash) AS row_num ")
                .append("FROM ")
                .append("files f ")
                .append("INNER JOIN companies c ON f.f_company_fkey = c.cmp_id ")
                .append("INNER JOIN user_certificates uc ON f.f_entity_fkey = uc.cert_id ")
                .append("INNER JOIN users u ON f.f_entity_main_fkey = u.usr_id ")
                .append("INNER JOIN company_user_calendar_events cuce ON cuce.cuce_user_fkey = f.f_entity_main_fkey ")
                .append("INNER JOIN company_contractor_requests request_event ON request_event.cr_id = cuce.cuce_request_fkey ")
                .append("WHERE ")
                .append("f.f_entity_type = 'user-certificates' ")
                .append("AND f.f_deleted = false ")
                .append("AND uc.cert_deleted = false ")
                .append("AND uc.cert_end_date > CURRENT_DATE ")
                .append("AND cuce.cuce_deleted = false ")
                .append("AND uc.cert_type = 'CERTIFICATE_TYPE_WELDER' ")
                .append(") AS subquery ")
                .append("WHERE ")
                .append("row_num = 1 ")
                .append("AND cert_number IN (");

        // Dodaj parametry dla numerów certyfikatów
        for (int i = 0; i < numberOfParams; i++) {
            queryBuilder.append("?");
            if (i < numberOfParams - 1) {
                queryBuilder.append(", ");
            }
        }

        queryBuilder.append(") ")
                .append("GROUP BY ")
                .append("usr_name, ")
                .append("cr_number, ")
                .append("cert_number, ")
                .append("file_url;");

        return queryBuilder.toString();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new DownloadPDFsUI());
    }
}
