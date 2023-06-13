#!/usr/bin/php
<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use \PhpOffice\PhpSpreadsheet\Shared\Date;
use Symfony\Component\Mailer\Transport;
use Symfony\Component\Mailer\Mailer;
use Symfony\Component\Mime\Email;

$dotenv = Dotenv\Dotenv::createImmutable(__DIR__);
$dotenv->load();
$dotenv->required(['DB_HOST', 'DB_NAME', 'DB_USER', 'DB_PASS']);

$transport = Transport::fromDsn($_ENV['DSN']);
$mailer = new Mailer($transport);

$datej = ($_ENV['DATE_J']?$_ENV['DATE_J']:date('yyyymmdd'));
$datem = ($_ENV['DATE_M']?$_ENV['DATE_M']:date('yyyymmdd'));

$link = mysqli_connect($_ENV['DB_HOST'], $_ENV['DB_USER'], $_ENV['DB_PASS'], $_ENV['DB_NAME']);
$link->set_charset("utf8mb4");

$spreadsheet = new Spreadsheet();
$countWorksheet = 0;
addWorkSheet($datej);
addWorkSheet($datem);
addTotalWorkSheet($datem);

// Enregistrement du fichier
$writer = new Xlsx($spreadsheet);
try {
    $writer->save($_ENV['REPOSITORY'] . $datej . '.xlsx');
} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
    $email = (new Email())
        ->from($_ENV['FROM'])
        ->to($_ENV['TO'])
        //->cc('cc@example.com')
        //->bcc('bcc@example.com')
        //->replyTo('fabien@example.com')
        //->priority(Email::PRIORITY_HIGH)
        ->subject("ERROR : création du fichier Excel impossible")
        ->text("Le script de génération des rapports Excel n'a pas fonctionné.");
}

$email = (new Email())
    ->from($_ENV['FROM'])
    ->to($_ENV['TO'])
    //->cc('cc@example.com')
    //->bcc('bcc@example.com')
    //->replyTo('fabien@example.com')
    //->priority(Email::PRIORITY_HIGH)
    ->subject($_ENV['SUBJECT'])
    ->text('Sending emails is fun again!')
    ->html('<p>See Twig integration for better HTML integration!</p>')
    ->attachFromPath($_ENV['REPOSITORY'] . $datej . '.xlsx');

$mailer->send($email);

function addWorkSheet($date) {
	global $countWorksheet, $link, $spreadsheet;
	
	if (!$countWorksheet++)
		$sheet = $spreadsheet->getActiveSheet();
	else
		$sheet = $spreadsheet->createSheet();
	
	$sheet->setTitle('Vente ' . $date);
	
	$data = "SELECT catn.label as CATEGORIE, f.ref, f.datef, f.pos_source as TERMINAL, fd.total_ht as MONTANT, fd.qty as QUANTITE, p.ref as REFERENCE, p.label as DESCRIPTION FROM llx_societe as s LEFT JOIN llx_c_country as c on s.fk_pays = c.rowid LEFT JOIN llx_facture as f ON  s.rowid = f.entity LEFT JOIN llx_c_departements as cd on s.fk_departement = cd.rowid LEFT JOIN llx_projet as pj ON f.fk_projet = pj.rowid LEFT JOIN llx_user as uc ON f.fk_user_author = uc.rowid LEFT JOIN llx_user as uv ON f.fk_user_valid = uv.rowid LEFT JOIN llx_facturedet as fd ON f.rowid =fd.fk_facture LEFT JOIN llx_facture_extrafields as extra ON f.rowid = extra.fk_object LEFT JOIN llx_facturedet_extrafields as extra2 on fd.rowid = extra2.fk_object LEFT JOIN llx_product as p on (fd.fk_product = p.rowid) LEFT JOIN llx_product_extrafields as extra3 on p.rowid = extra3.fk_object LEFT JOIN llx_categorie_product as cat on cat.fk_product = fd.fk_product LEFT JOIN llx_categorie as catn ON catn.rowid = cat.fk_categorie  WHERE f.rowid = fd.fk_facture AND f.entity IN (1) and date_format(f.datef,'%Y%m') = " . $date . " ORDER BY CATEGORIE;";

	$row = 1;
	$sheet->insertNewRowBefore($row);
	$sheet->getColumnDimension('A')->setWidth(20);
	$sheet->getColumnDimension('B')->setWidth(20);
	$sheet->getColumnDimension('C')->setWidth(20);
	$sheet->getColumnDimension('D')->setWidth(10);
	$sheet->getColumnDimension('E')->setWidth(10);
	$sheet->getColumnDimension('F')->setWidth(10);
	$sheet->getColumnDimension('G')->setWidth(30);
	$sheet->getColumnDimension('H')->setWidth(100);
	$sheet->setCellValue('A'.$row, 'CATEGORIE');
	$sheet->setCellValue('B'.$row, 'REF FACTURE');
	$sheet->setCellValue('C'.$row, 'DATE');
	$sheet->setCellValue('D'.$row, 'TERMINAL');
	$sheet->setCellValue('E'.$row, 'MONTANT');
	$sheet->setCellValue('F'.$row, 'QUANTITE');
	$sheet->setCellValue('G'.$row, 'REF PRODUIT');
	$sheet->setCellValue('H'.$row, 'DESCRIPTION');
	
	$result = $link->query($data)->fetch_all(MYSQLI_ASSOC);
	if (count($result) > 0)
	{
		foreach ($result as $data) {
			$row = $sheet->getHighestRow()+1;
			$sheet->insertNewRowBefore($row);
			$sheet->setCellValue('A'.$row, $data['CATEGORIE']);
			$sheet->setCellValue('B'.$row, $data['ref']);
			$sheet->getStyle('C'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
			$sheet->setCellValue('C'.$row, Date::PHPToExcel($data['datef']));
			$sheet->setCellValue('D'.$row, $data['TERMINAL']);
			$sheet->getStyle('E'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
			$sheet->setCellValue('E'.$row, $data['MONTANT']);
			$sheet->setCellValue('F'.$row, $data['QUANTITE']);
			$sheet->setCellValue('G'.$row, $data['REFERENCE']);
			$sheet->setCellValue('H'.$row, $data['DESCRIPTION']);
		}
	}
}

function addTotalWorkSheet($date) {
	global $countWorksheet, $link, $spreadsheet;
	
	if (!$countWorksheet++)
		$sheet = $spreadsheet->getActiveSheet();
	else
		$sheet = $spreadsheet->createSheet();
	
	$sheet->setTitle('TOTAL ' . $date);
	
	$data = "SELECT catn.label as CATEGORIE, sum(fd.total_ht) as MONTANT FROM llx_societe as s LEFT JOIN llx_c_country as c on s.fk_pays = c.rowid LEFT JOIN llx_facture as f ON  s.rowid = f.entity LEFT JOIN llx_c_departements as cd on s.fk_departement = cd.rowid LEFT JOIN llx_projet as pj ON f.fk_projet = pj.rowid LEFT JOIN llx_user as uc ON f.fk_user_author = uc.rowid LEFT JOIN llx_user as uv ON f.fk_user_valid = uv.rowid LEFT JOIN llx_facturedet as fd ON f.rowid =fd.fk_facture LEFT JOIN llx_facture_extrafields as extra ON f.rowid = extra.fk_object LEFT JOIN llx_facturedet_extrafields as extra2 on fd.rowid = extra2.fk_object LEFT JOIN llx_product as p on (fd.fk_product = p.rowid) LEFT JOIN llx_product_extrafields as extra3 on p.rowid = extra3.fk_object LEFT JOIN llx_categorie_product as cat on cat.fk_product = fd.fk_product LEFT JOIN llx_categorie as catn ON catn.rowid = cat.fk_categorie  WHERE f.rowid = fd.fk_facture AND f.entity IN (1) and date_format(f.datef,'%Y%m') = " . $date . " GROUP BY CATEGORIE ORDER BY CATEGORIE;";

	$row = 1;
	$sheet->insertNewRowBefore($row);
	$sheet->getColumnDimension('A')->setWidth(20);
	$sheet->getColumnDimension('E')->setWidth(10);
	$sheet->setCellValue('A'.$row, 'CATEGORIE');
	$sheet->setCellValue('E'.$row, 'MONTANT');
	
	$result = $link->query($data)->fetch_all(MYSQLI_ASSOC);
	if (count($result) > 0)
	{
		foreach ($result as $data) {
			$row = $sheet->getHighestRow()+1;
			$sheet->insertNewRowBefore($row);
			$sheet->setCellValue('A'.$row, $data['CATEGORIE']);
			$sheet->getStyle('E'.$row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR);
			$sheet->setCellValue('E'.$row, $data['MONTANT']);
		}
	}
}

?>
