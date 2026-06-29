<?php
require 'vendor/autoload.php';

// Configuration de la base de données
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "planning";
$excel_file = 'BASE_DE_PROSPECTION_CELTIIS_ABDALLAH-SURFACE-PRO8-GZ.xlsx';
// Mapping complet des 13 portefeuilles
$portefeuilles = [
    1 => ['file' => '1 CONSEILS COM.xlsx', 'nom' => 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'],
    /*2 => ['file' => '2 BTP.xlsx', 'nom' => 'BTP_Franck AROUNA'],
    3 => ['file' => '3 FINANCES.xlsx', 'nom' => 'Banque et finance_Carmen AWO'],
    4 => ['file' => '4 COMMERCE.xlsx', 'nom' => 'Commerce_Felix AKPLOGAN et Rafiou AKPONI AKOUNSOUHA'],
    5 => ['file' => '5 TRANSPORT.xlsx', 'nom' => 'Transport auto moto_Elvis KONGNON'],
    6 => ['file' => '6 ICT.xlsx', 'nom' => 'ICT_Emerick'],
    7 => ['file' => '7 HOTEL TOURISME.xlsx', 'nom' => 'Hôtel tourisme_Khader ASSOGBA et Géraud DJOI'],
    8 => ['file' => '8 CAFE BAR RESTAURANT.xlsx', 'nom' => 'Café bar restaurant_Arétas HOUANSOU et Herman AGOSSA'],
    9 => ['file' => '9 SANTE ET SOINS.xlsx', 'nom' => 'Santé et soins_Anne Marie DOSSA'],
    10 => ['file' => '10 ALIMENTATION.xlsx', 'nom' => 'Alimentation_Benedicte ANOUMOU'],
    11 => ['file' => '11 MODE ARTISANAT.xlsx', 'nom' => 'Mode et artisanat_Carole Tambamou'],
    12 => ['file' => '12 FORMATION.xlsx', 'nom' => 'Formation et Admin_Gloria TOSSE'],
    13 => ['file' => '13 SECURITE ET SERVICES.xlsx', 'nom' => 'Sécurité et services_Gatien ZENOUVOU']*/
];

try {
    // Connexion à la base de données
    $conn = new PDO("mysql:host=$servername;dbname=$dbname", $username, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    
    // Préparation de la requête d'insertion
    $stmt = $conn->prepare("
        INSERT INTO portefeuilles (
            numero_portefeuille, nom_portefeuille, nom_Entreprise, nom_Responsable, 
            fonction, telephone, email, ville, activite_Principale, IFU, QIP, RCCM
        ) VALUES (
            :num, :nom_pf, :entreprise, :responsable, :fonction, :tel, 
            :email, :ville, :activite, :ifu, :qip, :rccm
        )
    ");

    // Compteur global d'importations
    $total_imported = 0;

    foreach ($portefeuilles as $num => $pf) {
        echo "<h3>Traitement du portefeuille $num - {$pf['nom']}</h3>";
        
        // Vérification de l'existence du fichier
        if (!file_exists($pf['file'])) {
            echo "<p style='color:red'>Erreur: Fichier {$pf['file']} introuvable.</p>";
            continue;
        }

        try {
            // Chargement du fichier Excel
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($pf['file']);
            $worksheet = $spreadsheet->getActiveSheet();
            
            $imported_count = 0;
            
            foreach ($worksheet->getRowIterator(2) as $row) { // On commence à la ligne 2 (ligne 1 = en-têtes)
                $data = [];
                foreach ($row->getCellIterator() as $cell) {
                    $data[] = $cell->getValue();
                }
                
                // Vérification du nombre de colonnes
                if (count($data) < 10) {
                    echo "<p style='color:orange'>Ligne {$row->getRowIndex()}: Nombre de colonnes insuffisant</p>";
                    continue;
                }
                
                // Mapping des colonnes selon votre structure:
                // 0: NOM ENTREPRISES
                // 1: NOM DU POINT FOCAL (responsable)
                // 2: FONCTION
                // 3: Téléphone
                // 4: EMAIL
                // 5: VILLE
                // 6: ACTIVITE PRINCIPALE
                // 7: IFU
                // 8: QIP
                // 9: RCCM
                
                // Nettoyage des données
                $entreprise = isset($data[0]) ? trim($data[0]) : '';
                $responsable = isset($data[1]) ? trim($data[1]) : '';
                $fonction = isset($data[2]) ? trim($data[2]) : '';
                $telephone = isset($data[3]) ? trim($data[3]) : '';
                $email = isset($data[4]) ? trim($data[4]) : '';
                $ville = isset($data[5]) ? trim($data[5]) : '';
                $activite = isset($data[6]) ? trim($data[6]) : '';
                $ifu = isset($data[7]) ? trim($data[7]) : '';
                $qip = isset($data[8]) ? trim($data[8]) : '';
                $rccm = isset($data[9]) ? trim($data[9]) : '';

                // Insertion des données
                $stmt->execute([
                    ':num' => $num,
                    ':nom_pf' => $pf['nom'],
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
            }
            
            $total_imported += $imported_count;
            echo "<p>Succès: $imported_count enregistrements importés</p>";
            
        } catch (Exception $e) {
            echo "<p style='color:red'>Erreur lors du traitement du portefeuille $num: " . $e->getMessage() . "</p>";
            error_log("Erreur portefeuille $num: " . $e->getMessage());
            continue;
        }
    }
    
    echo "<h2>Résumé final</h2>";
    echo "<p>$total_imported enregistrements importés au total sur les 13 portefeuilles.</p>";
    
} catch(PDOException $e) {
    die("<p style='color:red'>Erreur de connexion à la base de données: " . $e->getMessage() . "</p>");
} catch(Exception $e) {
    die("<p style='color:red'>Erreur générale: " . $e->getMessage() . "</p>");
}
?>