<?php
ini_set('memory_limit', '1G'); // Augmentation significative de la mémoire
set_time_limit(0); // Pas de limite de temps d'exécution
require 'vendor/autoload.php';

// Configuration
$excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx';
$sheet_name = '1 CONSEILS COM';
$chunk_size = 1000; // Nombre de lignes à traiter à la fois

try {
    // 1. Initialisation du Reader avec optimisation maximale
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setReadDataOnly(true);
    $reader->setReadEmptyCells(false);
    $reader->setLoadSheetsOnly([$sheet_name]);
    
    // 2. Calcul du nombre total de lignes sans tout charger
    $worksheetInfo = $reader->listWorksheetInfo($excel_file);
    $totalRows = $worksheetInfo[0]['totalRows'];
    $totalColumns = $worksheetInfo[0]['totalColumns'];
    
    echo "Fichier contient $totalRows lignes et $totalColumns colonnes\n";

    // 3. Connexion MySQL
    $conn = new PDO("mysql:host=localhost;dbname=planning", "root", "");
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $conn->setAttribute(PDO::ATTR_EMULATE_PREPARES, false);

    // 4. Traitement par blocs
    $imported_count = 0;
    for ($startRow = 2; $startRow <= $totalRows; $startRow += $chunk_size) {
        $endRow = min($startRow + $chunk_size - 1, $totalRows);
        
        // Chargement partiel du fichier
        $reader->setReadFilter(new ChunkReadFilter($startRow, $endRow));
        $spreadsheet = $reader->load($excel_file);
        $worksheet = $spreadsheet->getActiveSheet();
        
        // Traitement du bloc
        for ($row = $startRow; $row <= $endRow; $row++) {
            $data = [];
            for ($col = 'A'; $col <= $worksheet->getHighestColumn(); $col++) {
                $data[] = $worksheet->getCell($col.$row)->getValue();
            }
            
            if (count(array_filter($data)) < 1) continue; // Ignore les lignes vides

            try {
                $conn->beginTransaction();
                $stmt = $conn->prepare("INSERT INTO portefeuilles (...) VALUES (...)");
                $stmt->execute([
                    ':num' => 1,
                    ':nom_pf' => 'Conseils et Communication_Isnelle HOUEKPONHOUNDE',
                    ':entreprise' => $data[0] ?? null,
                    // ... autres bindings
                ]);
                $conn->commit();
                $imported_count++;
            } catch (PDOException $e) {
                $conn->rollBack();
                echo "Erreur ligne $row: ".$e->getMessage()."\n";
            }
        }
        
        // Libération mémoire
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        gc_collect_cycles();
        
        echo "Progression: ".min($endRow, $totalRows)."/$totalRows lignes traitées\n";
    }

    echo "Importation réussie! $imported_count enregistrements importés.\n";

} catch(Exception $e) {
    die("ERREUR: ".$e->getMessage());
}

// Filtre de lecture par blocs
class ChunkReadFilter implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
    private $startRow = 0;
    private $endRow = 0;

    public function __construct($startRow, $endRow) {
        $this->startRow = $startRow;
        $this->endRow = $endRow;
    }

    public function readCell($column, $row, $worksheetName = '') {
        return ($row >= $this->startRow && $row <= $this->endRow);
    }
}
?>