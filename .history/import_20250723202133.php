<?php
ini_set('memory_limit', '512M');
require 'vendor/autoload.php';

// Configuration de la base de données
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "planning";
$excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx';

// Vérification de l'existence du fichier Excel
if (!file_exists($excel_file)) {
    die("<p style='color:red'>ERREUR: Le fichier Excel '$excel_file' est introuvable. Voici les fichiers présents:<br>"
        . implode("<br>", scandir(__DIR__)) . "</p>");
}

try {
    // Chargement du fichier Excel
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($excel_file);
    
    // Accès à la feuille spécifique
    $worksheet = $spreadsheet->getSheetByName('1 CONSEILS COM');
    if (!$worksheet) {
        die("<p style='color:red'>ERREUR: La feuille '1 CONSEILS COM' est introuvable. Feuilles disponibles:<br>"
            . implode("<br>", $spreadsheet->getSheetNames()) . "</p>");
    }

    // Connexion MySQL
    $conn = new PDO("mysql:host=$servername;dbname=$dbname", $username, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Préparation de la requête
    $stmt = $conn->prepare("
        INSERT INTO portefeuilles (
            numero_portefeuille, nom_portefeuille, nom_Entreprise, nom_Responsable, 
            fonction, telephone, email, ville, activite_Principale, IFU, QIP, RCCM
        ) VALUES (
            :num, :nom_pf, :entreprise, :responsable, :fonction, :tel, 
            :email, :ville, :activite, :ifu, :qip, :rccm
        )
    ");

    // Paramètres pour ce portefeuille
    $portefeuille_num = 1;
    $portefeuille_nom = 'Conseils et Communication_Isnelle HOUEKPONHOUNDE';
    $imported_count = 0;

    // Parcours des lignes (en commençant à la ligne 2 pour ignorer les en-têtes)
    foreach ($worksheet->getRowIterator(2) as $row) {
        $data = [];
        foreach ($row->getCellIterator() as $cell) {
            $data[] = $cell->getValue();
        }

        // Vérification du nombre minimal de colonnes
        if (count($data) < 10) {
            echo "<p style='color:orange'>Ligne {$row->getRowIndex()}: Nombre de colonnes insuffisant (".count($data).")</p>";
            continue;
        }

        // Mapping des colonnes
        $entreprise        = isset($data[0]) ? trim($data[0]) : '';
        $responsable      = isset($data[1]) ? trim($data[1]) : '';
        $fonction         = isset($data[2]) ? trim($data[2]) : '';
        $telephone        = isset($data[3]) ? trim($data[3]) : '';
        $email            = isset($data[4]) ? trim($data[4]) : '';
        $ville            = isset($data[5]) ? trim($data[5]) : '';
        $activite         = isset($data[6]) ? trim($data[6]) : '';
        $ifu              = isset($data[7]) ? trim($data[7]) : '';
        $qip              = isset($data[8]) ? trim($data[8]) : '';
        $rccm             = isset($data[9]) ? trim($data[9]) : '';

        // Insertion
        try {
            $stmt->execute([
                ':num' => $portefeuille_num,
                ':nom_pf' => $portefeuille_nom,
                ':entreprise' => $entreprise,
                ':responsable' => $responsable,
                ':fonction' => $fonction,
                ':tel' => $telephone,
                ':email' => $email,
                ':ville' => $ville,
                ':activite' => $activite,
                ':ifu' => $ifu,
                ':qip' => $qip,
                ':rccm' => $rccm
            ]);
            $imported_count++;
        } catch (PDOException $e) {
            echo "<p style='color:orange'>Erreur d'insertion ligne {$row->getRowIndex()}: ".$e->getMessage()."</p>";
        }
    }

    echo "<h3>Résumé de l'importation</h3>";
    echo "<p style='color:green'>Importation réussie! $imported_count enregistrements importés depuis la feuille '1 CONSEILS COM'.</p>";

} catch(PDOException $e) {
    die("<p style='color:red'>ERREUR MySQL: " . $e->getMessage() . "</p>");
} catch(Exception $e) {
    die("<p style='color:red'>ERREUR: " . $e->getMessage() . "</p>");
}
?>