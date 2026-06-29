<?php
ini_set('memory_limit', '512M');
require 'vendor/autoload.php';

// Configuration
$excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx';
$delimiter = ','; // Délimiteur CSV

try {
    // 1. CONVERSION DIRECTE EN TABLEAU (plus fiable que via CSV)
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setReadDataOnly(true);
    $spreadsheet = $reader->load($excel_file);
    $worksheet = $spreadsheet->getActiveSheet(); // ou getSheetByName() si nécessaire

    // 2. CONNEXION MySQL
    $conn = new PDO("mysql:host=localhost;dbname=planning", "root", "");
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $stmt = $conn->prepare("INSERT INTO portefeuilles (...) VALUES (...)");

    // 3. LECTURE OPTIMISÉE
    $imported_count = 0;
    $highestRow = $worksheet->getHighestDataRow();
    $highestColumn = $worksheet->getHighestDataColumn();
    $columns = ['A','B','C','D','E','F','G','H','I','J']; // Ajustez selon vos colonnes

    for ($row = 2; $row <= $highestRow; $row++) {
        $rowData = [];
        foreach ($columns as $col) {
            $rowData[] = $worksheet->getCell($col.$row)->getValue();
        }

        // Vérification des données
        if (count(array_filter($rowData)) < 1) continue; // Ignore les lignes vides

        try {
            $stmt->execute([
                ':num' => 1,
                ':nom_pf' => 'Conseils et Communication_Isnelle HOUEKPONHOUNDE',
                ':entreprise' => $rowData[0] ?? null,
                ':responsable' => $rowData[1] ?? null,
                ':fonction' => $rowData[2] ?? null,
                ':tel' => $rowData[3] ?? null,
                ':email' => $rowData[4] ?? null,
                ':ville' => $rowData[5] ?? null,
                ':activite' => $rowData[6] ?? null,
                ':ifu' => $rowData[7] ?? null,
                ':qip' => $rowData[8] ?? null,
                ':rccm' => $rowData[9] ?? null
            ]);
            $imported_count++;
        } catch (PDOException $e) {
            echo "<p style='color:orange'>Erreur ligne $row: ".$e->getMessage()."</p>";
        }
    }

    echo "<h3>Résumé de l'importation</h3>";
    echo "<p style='color:green'>Importation réussie! $imported_count enregistrements importés.</p>";

} catch(Exception $e) {
    die("<p style='color:red'>ERREUR: ".$e->getMessage()."</p>");
}