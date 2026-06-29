<?php
require 'vendor/autoload.php';


$servername = "localhost";
$username = "votre_utilisateur";
$password = "votre_mot_de_passe";
$dbname = "planning";

// Mapping des numéros de portefeuille aux noms de fichiers et libellés
$portefeuilles = [
    1 => ['file' => '1 CONSEILS COM.xlsx', 'nom' => 'Conseils et Communication_Isnelle HOUEKPONHOUNDE'],
    2 => ['file' => '2 BTP.xlsx', 'nom' => 'BTP_Franck AROUNA'],
    2 => ['file' => '3 FINANCES.xlsx', 'nom' => 'Banque et finance_Carmen AWO'],
    2 => ['file' => '4 COMMERCE.xlsx', 'nom' => 'Commerce_Felix AKPLOGAN et Rafiou AKPONI AKOUNSOUHA'],
    2 => ['file' => '5 TRANSPORT.xlsx', 'nom' => 'Transport auto moto_Elvis KONGNON'],
    2 => ['file' => '6 ICT.xlsx', 'nom' => 'ICT_Emerick'],
    2 => ['file' => '7 HOTEL TOURISME.xlsx', 'nom' => 'Hôtel tourisme_Khader ASSOGBA et Géraud DJOI'],
    2 => ['file' => '8 CAFE BAR RESTAURANT.xlsx', 'nom' => 'Café bar restaurant_Arétas HOUANSOU et Herman AGOSSA'],
    2 => ['file' => '9 SANTE ET SOINS.xlsx', 'nom' => 'Santé et soins_ Anne marie DOSSA'],
    2 => ['file' => '10 ALIMENTATION.xlsx', 'nom' => 'Alimentation_Benedicte ANOUMOU'],
    2 => ['file' => '11 MODE ARTISANAT.xlsx', 'nom' => 'Mode et artisanat_Carole Tambamou'],
    2 => ['file' => '12 FORMATION.xlsx', 'nom' => 'Alimentation_Benedicte ANOUMOU'],
    2 => ['file' => '10 ALIMENTATION.xlsx', 'nom' => 'Alimentation_Benedicte ANOUMOU'],
    // ... ajoutez les 11 autres portefeuilles
];

try {
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

    foreach ($portefeuilles as $num => $pf) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($pf['file']);
        $worksheet = $spreadsheet->getActiveSheet();
        
        foreach ($worksheet->getRowIterator(2) as $row) { // Supposant ligne 1 = en-têtes
            $data = [];
            foreach ($row->getCellIterator() as $cell) {
                $data[] = $cell->getValue();
            }
            
            $stmt->execute([
                ':num' => $num,
                ':nom_pf' => $pf['nom'],
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
        }
        echo "Portefeuille $num importé.<br>";
    }
} catch(Exception $e) {
    die("Erreur: " . $e->getMessage());
}
?>