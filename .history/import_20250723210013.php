<?php
ini_set('memory_limit', '1G');
set_time_limit(0);
require 'vendor/autoload.php';

// Définition du filtre de lecture par blocs
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

// Configuration
$excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx';
$sheet_name = '1 CONSEILS COM';
$chunk_size = 500; // Nombre de lignes à traiter à la fois
$portefeuille_num = 1;
$portefeuille_nom = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE';

try {
    // Initialisation du lecteur Excel
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setReadDataOnly(true);
    $reader->setReadEmptyCells(false);
    $reader->setLoadSheetsOnly([$sheet_name]);
    
    // Analyse du fichier
    $worksheetInfo = $reader->listWorksheetInfo($excel_file);
    if (empty($worksheetInfo)) {
        die("Aucune information sur les feuilles trouvée dans le fichier.");
    }
    
    $totalRows = $worksheetInfo[0]['totalRows'];
    $totalColumns = $worksheetInfo[0]['totalColumns'];
    
    echo "Fichier contient $totalRows lignes et $totalColumns colonnes\n";

    // Connexion MySQL
    $conn = new PDO("mysql:host=localhost;dbname=planning", "root", "");
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Préparation de la requête SQL
    $sql = "INSERT INTO portefeuilles (
                numero_portefeuille, nom_portefeuille, nom_Entreprise, 
                nom_Responsable, fonction, telephone, email, ville, 
                activite_Principale, IFU, QIP, RCCM, portefeuille
            ) VALUES (
                :num, :nom_pf, :entreprise, :responsable, :fonction, 
                :tel, :email, :ville, :activite, :ifu, :qip, :rccm, :pf
            )";
    $stmt = $conn->prepare($sql);

    // Traitement par blocs
    $imported_count = 0;
    for ($startRow = 2; $startRow <= $totalRows; $startRow += $chunk_size) {
        $endRow = min($startRow + $chunk_size - 1, $totalRows);
        
        echo "Traitement des lignes $startRow à $endRow...\n";
        
        $reader->setReadFilter(new ChunkReadFilter($startRow, $endRow));
        $spreadsheet = $reader->load($excel_file);
        $worksheet = $spreadsheet->getActiveSheet();
        
        for ($row = $startRow; $row <= $endRow; $row++) {
            // Mapping des colonnes Excel vers la base de données
            $data = [
                'entreprise'    => $worksheet->getCell('A'.$row)->getValue(),
                'responsable'   => $worksheet->getCell('B'.$row)->getValue(),
                'fonction'     => $worksheet->getCell('C'.$row)->getValue(),
                'telephone'     => $worksheet->getCell('D'.$row)->getValue(),
                'email'        => $worksheet->getCell('E'.$row)->getValue(),
                'ville'         => $worksheet->getCell('F'.$row)->getValue(),
                'activite'      => $worksheet->getCell('G'.$row)->getValue(),
                'ifu'           => $worksheet->getCell('H'.$row)->getValue(),
                'qip'           => $worksheet->getCell('I'.$row)->getValue(),
                'rccm'          => $worksheet->getCell('J'.$row)->getValue()
            ];
            
            // Ignorer les lignes vides
            if (empty(array_filter($data))) continue;

            try {
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
            } catch (PDOException $e) {
                echo "Erreur ligne $row: ".$e->getMessage()."\n";
            }
        }
        
        // Nettoyage mémoire
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        gc_collect_cycles();
    }

    echo "Importation terminée avec succès! $imported_count enregistrements importés.\n";

} catch(Exception $e) {
    die("ERREUR: ".$e->getMessage());
}
?>