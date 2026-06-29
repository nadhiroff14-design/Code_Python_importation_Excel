<?php
ini_set('memory_limit', '1G');
set_time_limit(0);
require 'vendor/autoload.php';

// Configuration
$excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx';
$sheet_name = '1 CONSEILS COM';
$portefeuille_num = 1;
$portefeuille_nom = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE';

// Vérification du fichier
if (!file_exists($excel_file)) {
    die("Erreur : Fichier Excel introuvable.\nFichiers disponibles :\n".implode("\n", scandir(__DIR__)));
}

// Classe pour lecture par blocs
class ChunkReadFilter implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
    private $startRow;
    private $endRow;
    
    public function __construct($startRow, $endRow) {
        $this->startRow = $startRow;
        $this->endRow = $endRow;
    }
    
    public function readCell($column, $row, $worksheetName = '') {
        return ($row >= $this->startRow && $row <= $this->endRow);
    }
}

try {
    // 1. Initialisation du lecteur Excel
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setReadDataOnly(true);
    $reader->setReadEmptyCells(false);
    
    // Vérification que le fichier est lisible
    if (!$reader->canRead($excel_file)) {
        throw new Exception("Impossible de lire le fichier Excel. Format non supporté ou fichier corrompu.");
    }
    
    // Analyse du fichier pour obtenir les infos
    $worksheetInfo = $reader->listWorksheetInfo($excel_file);
    if (empty($worksheetInfo)) {
        throw new Exception("Aucune feuille trouvée dans le fichier Excel.");
    }
    
    $totalRows = $worksheetInfo[0]['totalRows'];
    $totalColumns = $worksheetInfo[0]['totalColumns'];
    
    echo "Fichier Excel détecté : $totalRows lignes, $totalColumns colonnes\n";

    // 2. Connexion MySQL
    $conn = new PDO("mysql:host=localhost;dbname=planning", "root", "");
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $conn->exec("SET NAMES utf8mb4");

    // 3. Préparation de la requête SQL
    $sql = "INSERT INTO portefeuilles (
                numero_portefeuille, nom_portefeuille, nom_Entreprise, 
                nom_Responsable, fonction, telephone, email, ville, 
                activite_Principale, IFU, QIP, RCCM, portefeuille
            ) VALUES (
                :num, :nom_pf, :entreprise, :responsable, :fonction, 
                :tel, :email, :ville, :activite, :ifu, :qip, :rccm, :pf
            )";
    $stmt = $conn->prepare($sql);

    // 4. Traitement par blocs de 500 lignes
    $chunk_size = 500;
    $imported_count = 0;
    $error_count = 0;
    
    for ($startRow = 2; $startRow <= $totalRows; $startRow += $chunk_size) {
        $endRow = min($startRow + $chunk_size - 1, $totalRows);
        
        echo "Traitement des lignes $startRow à $endRow... ";
        
        // Configuration du filtre
        $reader->setReadFilter(new ChunkReadFilter($startRow, $endRow));
        $reader->setLoadSheetsOnly([$sheet_name]);
        
        // Chargement du bloc
        $spreadsheet = $reader->load($excel_file);
        $worksheet = $spreadsheet->getActiveSheet();
        
        // Traitement des lignes
        for ($row = $startRow; $row <= $endRow; $row++) {
            try {
                // Lecture des données
                $data = [
                    'entreprise'    => $worksheet->getCell('A'.$row)->getValue(),
                    'responsable'   => $worksheet->getCell('B'.$row)->getValue(),
                    'fonction'      => $worksheet->getCell('C'.$row)->getValue(),
                    'telephone'     => $worksheet->getCell('D'.$row)->getValue(),
                    'email'        => $worksheet->getCell('E'.$row)->getValue(),
                    'ville'        => $worksheet->getCell('F'.$row)->getValue(),
                    'activite'     => $worksheet->getCell('G'.$row)->getValue(),
                    'ifu'          => $worksheet->getCell('H'.$row)->getValue(),
                    'qip'          => $worksheet->getCell('I'.$row)->getValue(),
                    'rccm'         => $worksheet->getCell('J'.$row)->getValue()
                ];
                
                // Ignorer les lignes vides
                if (empty(array_filter($data, function($v) { return $v !== null && $v !== ''; }))) {
                    continue;
                }
                
                // Insertion
                $stmt->execute([
                    ':num'         => $portefeuille_num,
                    ':nom_pf'      => $portefeuille_nom,
                    ':entreprise'  => $data['entreprise'],
                    ':responsable' => $data['responsable'],
                    ':fonction'    => $data['fonction'],
                    ':tel'         => $data['telephone'],
                    ':email'       => $data['email'],
                    ':ville'       => $data['ville'],
                    ':activite'    => $data['activite'],
                    ':ifu'         => $data['ifu'],
                    ':qip'         => $data['qip'],
                    ':rccm'        => $data['rccm'],
                    ':pf'          => $portefeuille_nom
                ]);
                
                $imported_count++;
            } catch (Exception $e) {
                $error_count++;
                echo "E";
                error_log("Erreur ligne $row: ".$e->getMessage());
                continue;
            }
        }
        
        // Nettoyage mémoire
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        gc_collect_cycles();
        
        echo " OK\n";
    }

    // 5. Résumé final
    echo "\nRésumé de l'importation :\n";
    echo "- Lignes importées avec succès : $imported_count\n";
    echo "- Erreurs rencontrées : $error_count\n";
    echo "- Total traité : ".($imported_count + $error_count)." sur ".($totalRows - 1)." lignes de données\n";

} catch (Exception $e) {
    die("\nERREUR CRITIQUE : " . $e->getMessage());
}
?>