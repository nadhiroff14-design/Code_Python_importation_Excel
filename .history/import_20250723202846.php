<?php
ini_set('memory_limit', '512M');
require 'vendor/autoload.php';

// Configuration de la base de données
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "planning";
$excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx';
$csv_temp_file = 'temp_import.csv';

// Vérification de l'existence du fichier Excel
if (!file_exists($excel_file)) {
    die("<p style='color:red'>ERREUR: Le fichier Excel '$excel_file' est introuvable. Voici les fichiers présents:<br>"
        . implode("<br>", scandir(__DIR__)) . "</p>");
}

try {
    // 1. CONVERSION EN CSV (mémoire optimisée)
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setReadDataOnly(true); // Ignore le formatage
    $reader->setReadEmptyCells(false); // Ignore les cellules vides
    
    // Charge uniquement la feuille souhaitée
    $spreadsheet = $reader->load($excel_file);
    $worksheet = $spreadsheet->getSheetByName('1 CONSEILS COM');
    
    if (!$worksheet) {
        die("<p style='color:red'>ERREUR: La feuille '1 CONSEILS COM' est introuvable. Feuilles disponibles:<br>"
            . implode("<br>", $spreadsheet->getSheetNames()) . "</p>");
    }

    // Sauvegarde temporaire en CSV
    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Csv($spreadsheet);
    $writer->save($csv_temp_file);
    
    // Libération mémoire
    unset($spreadsheet, $worksheet, $reader, $writer);

    // 2. IMPORT DEPUIS LE CSV
    $conn = new PDO("mysql:host=$servername;dbname=$dbname", $username, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $stmt = $conn->prepare("
        INSERT INTO portefeuilles (
            numero_portefeuille, nom_portefeuille, nom_Entreprise, nom_Responsable, 
            fonction, telephone, email, ville, activite_Principale, IFU, QIP, RCCM
        ) VALUES (
            :num, :nom_pf, :entreprise, :responsable, :fonction, :tel, 
            :email, :ville, :activite, :ifu, :qip, :rccm
        )
    ");

    $portefeuille_num = 1;
    $portefeuille_nom = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE';
    $imported_count = 0;
    $line_number = 0;

    if (($handle = fopen($csv_temp_file, 'r')) !== false) {
        while (($data = fgetcsv($handle)) !== false) {
            $line_number++;
            
            // Ignorer l'en-tête (première ligne)
            if ($line_number === 1) continue;
            
            // Vérification des colonnes
            if (count($data) < 10) {
                echo "<p style='color:orange'>Ligne $line_number: Nombre de colonnes insuffisant (".count($data).")</p>";
                continue;
            }

            // Nettoyage des données
            $data = array_map('trim', $data);
            
            try {
                $stmt->execute([
                    ':num' => $portefeuille_num,
                    ':nom_pf' => $portefeuille_nom,
                    ':entreprise' => $data[0] ?? null,
                    ':responsable' => $data[1] ?? null,
                    ':fonction' => $data[2] ?? null,
                    ':tel' => $data[3] ?? null,
                    ':email' => $data[4] ?? null,
                    ':ville' => $data[5] ?? null,
                    ':activite' => $data[6] ?? null,
                    ':ifu' => $data[7] ?? null,
                    ':qip' => $data[8] ?? null,
                    ':rccm' => $data[9] ?? null
                ]);
                $imported_count++;
            } catch (PDOException $e) {
                echo "<p style='color:orange'>Erreur ligne $line_number: ".$e->getMessage()."</p>";
            }
        }
        fclose($handle);
    }

    // Nettoyage du fichier temporaire
    if (file_exists($csv_temp_file)) {
        unlink($csv_temp_file);
    }

    echo "<h3>Résumé de l'importation</h3>";
    echo "<p style='color:green'>Importation réussie! $imported_count enregistrements importés depuis la feuille '1 CONSEILS COM'.</p>";

} catch(PDOException $e) {
    die("<p style='color:red'>ERREUR MySQL: " . $e->getMessage() . "</p>");
} catch(Exception $e) {
    // Nettoyage en cas d'erreur
    if (isset($handle) && is_resource($handle)) fclose($handle);
    if (file_exists($csv_temp_file)) unlink($csv_temp_file);
    die("<p style='color:red'>ERREUR: " . $e->getMessage() . "</p>");
}
?>