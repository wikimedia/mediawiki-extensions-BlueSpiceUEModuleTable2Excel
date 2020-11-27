<?php

use BlueSpice\UniversalExport\ExportModule;
use MediaWiki\MediaWikiServices;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Html;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class BsUEModuleTable2Excel extends ExportModule {

	/**
	 * Implementation of BsUniversalExportModule interface.
	 * @param SpecialUniversalExport &$oCaller
	 * @return array array(
	 *     'mime-type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	 *     'filename' => 'Filename.docx',
	 *     'content' => '8F3BC3025A7...'
	 * );
	 */
	public function createExportFile( &$oCaller ) {
		$aResponse = [];
		$sModeFrom = strtolower( $oCaller->getRequest()->getVal(
			'ModeFrom',
			'html'
		) );
		$sModeTo = strtolower( $oCaller->getRequest()->getVal(
			'ModeTo',
			'xls'
		) );
		$sContent = $oCaller->getRequest()->getVal( 'content', '' );

		$aOptions = array_merge_recursive(
			$this->getDefaultOptions( $oCaller ),
			$oCaller->getRequest()->getArray( 'options', [] )
		);
		$tmpFileName = time();

		if ( $sModeFrom == 'html' ) {
			$sContent = $this->prepareHTML( $sContent );
		}

		MediaWikiServices::getInstance()->getHookContainer()->run(
			'BsUEModuleTable2ExcelBeforeProcess',
			[
				$this,
				&$sModeFrom,
				&$sModeTo,
				&$sContent,
				&$tmpFileName
			]
		);

		$oStatus = BsFileSystemHelper::saveToCacheDirectory(
			"$tmpFileName.$sModeFrom",
			$sContent,
			'UEModuleTable2Excel'
		);

		if ( !$oStatus->isGood() ) {
			throw new Exception( $oStatus->getMessage() );
		}

		try {
			$oReader = IOFactory::createReaderForFile(
				"{$oStatus->getValue()}/$tmpFileName.$sModeFrom"
			);
		} catch ( Exception $e ) {
			throw new Exception( "Unsupportet mode $sModeFrom" );
		}

		// $oReader->load( "$tmpFileName.$sModeFrom" );
		$oReader instanceof Html;
		// apply reader options
		foreach ( $aOptions as $key => $val ) {
			$sFName = "set$key";
			if ( !method_exists( $oReader, $sFName ) ) {
				continue;
			}
			call_user_func(
				[ $oReader, $sFName ],
				$val
			);
		}
		$spreadsheet = $oReader->load(
			"{$oStatus->getValue()}/$tmpFileName.$sModeFrom"
		);

		$oProperties = $spreadsheet->getProperties();

		// apply properties
		foreach ( $aOptions as $key => $val ) {
			$sFName = "set$key";
			if ( !method_exists( $oProperties, $sFName ) ) {
				continue;
			}
			call_user_func(
				[ $oProperties, $sFName ],
				$val
			);
		}

		$this->adjustCellContent( $spreadsheet, $oCaller, $aOptions );

		$oWriter = IOFactory::createWriter( $spreadsheet, ucfirst( $sModeTo ) );

		// apply writer options
		foreach ( $aOptions as $key => $val ) {
			$sFName = "set$key";
			if ( !method_exists( $oWriter, $sFName ) ) {
				continue;
			}
			call_user_func(
				[ $oWriter, $sFName ],
				$val
			);
		}

		$oWriter->save( "{$oStatus->getValue()}/$tmpFileName.$sModeTo" );
		$sMimeType = MediaWikiServices::getInstance()->getMimeAnalyzer()->findMediaType(
			$sModeTo
		);

		$aResponse['filename'] = "$tmpFileName.$sModeTo";
		$aResponse['mime-type'] = $sMimeType;

		$oFile = BsFileSystemHelper::getCacheFileContent(
			"$tmpFileName.$sModeTo",
			'UEModuleTable2Excel'
		);

		if ( !$oFile->isGood() ) {
			throw new Exception( $oFile->getMessage() );
		}

		$aResponse['content'] = $oFile->getValue();

		$config = MediaWikiServices::getInstance()->getConfigFactory()->makeConfig(
			'bsg'
		);

		if ( !$config->get( 'TestMode' ) ) {
			unlink( "{$oStatus->getValue()}/$tmpFileName.$sModeTo" );
			unlink( "{$oStatus->getValue()}/$tmpFileName.$sModeFrom" );
		}
		return $aResponse;
	}

	/**
	 *
	 * @param SpecialUniversalExport $oCaller
	 * @param array $aOptions
	 * @return array
	 */
	protected function getDefaultOptions( SpecialUniversalExport $oCaller, $aOptions = [] ) {
		$oTitle = $oCaller->oRequestedTitle;

		if ( $oTitle->getNamespace() < 0 ) {
			$aOptions['Creator'] = $GLOBALS['wgSitename'];
			$aOptions['LastModifiedBy'] = $GLOBALS['wgSitename'];
		} else {
			$oWikiPage = WikiPage::factory( $oTitle );
			$util = MediaWikiServices::getInstance()->getService( 'BSUtilityFactory' );

			$aOptions['Creator'] = $util->getUserHelper(
				$oWikiPage->getCreator()
			)->getDisplayName();

			if ( $oWikiPage->getRevision() ) {
				$lastEditor = $oWikiPage->getRevision()->getUser();
				// sometimes this is an id - possible bug in mw version
				if ( is_int( $lastEditor ) ) {
					$lastEditor = \User::newFromId( $lastEditor );
				}
				$aOptions['LastModifiedBy'] = $util->getUserHelper(
					$lastEditor
				)->getDisplayName();
			}
		}

		$aOptions['Title'] = $oTitle->getFullText();
		$aOptions['Subject'] = $oTitle->getFullText();
		$aOptions['Description'] = '';
		$aOptions['Keywords'] = '';
		$aOptions['Category'] = '';

		$aOptions['Delimiter'] = ';';
		$aOptions['Enclosure'] = '';
		$aOptions['LineEnding'] = "\r\n";
		$aOptions['InputEncoding'] = 'UTF-8';

		return $aOptions;
	}

	/**
	 *
	 * @param string $sContent
	 * @return string
	 */
	protected function prepareHTML( $sContent ) {
		global $wgArticlePath, $wgServer, $wgSitename;

		$sPath = str_replace( '$1', '', $wgArticlePath );
		// replace relative links /1.23/index.php?title=Main_Page
		$sContent = preg_replace(
			'#href="(' . str_replace( '?', '.', $sPath ) . ')(.*?)"#s',
			'href="' . $wgServer . $sPath . '$2"',
			$sContent
		);

		// replace images with teir alt texts
		$sContent = preg_replace(
			'#< ?img.*?.alt=(\'|")(.*?)(\'|") .*?>#s',
			'$2',
			$sContent
		);

		// pwirth: in php versions smaller than 5.4 there is no ENT_HTML401 flag :(
		if ( defined( ENT_HTML401 ) ) {
			$sContent = htmlspecialchars( $sContent, ENT_HTML401, 'UTF-8' );
			$sContent = html_entity_decode( $sContent, ENT_HTML401, 'UTF-8' );
		}
		$sContent = str_replace( '&nbsp;', '', $sContent );
		$sContent = preg_replace(
			"/(^[\r\n]*|[\r\n]+)[\s\t]*[\r\n]+/",
			"\n",
			$sContent
		);
		$sContent = preg_replace( "/<thead>\s*<\/thead>/", "", $sContent );
		$docType = '-//W3C//DTD HTML 4.01 Transitional//EN';

		return <<<EOT
<!DOCTYPE HTML PUBLIC "$docType" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8">
		<title>$wgSitename</title>
	</head>
	<body>
		$sContent
	</body>
</html>
EOT;
	}

	/**
	 *
	 * @param Spreadsheet $spreadsheet
	 * @param SpecialUniversalExport $oCaller
	 * @param array $aOptions
	 */
	protected function adjustCellContent( Spreadsheet $spreadsheet, $oCaller, $aOptions ) {
		$sLastCol = $spreadsheet->getActiveSheet()->getHighestColumn();
		$iLastRow = $spreadsheet->getActiveSheet()->getHighestRow();
		$bFirstRowEmptyForNoReason = true;
		$sLastCol++;
		for ( $sCol = 'A'; $sCol != $sLastCol; $sCol++ ) {
			for ( $iRow = 1; $iRow <= $iLastRow; $iRow++ ) {
				$cell = $spreadsheet->getActiveSheet()->getCell(
					"$sCol$iRow"
				);
				$sVal = trim( $cell->getValue() );
				if ( $iRow == 1 && !empty( $sVal ) ) {
					$bFirstRowEmptyForNoReason = false;
				}
				$sModeTo = $oCaller->getRequest()->getVal( 'ModeTo', 'xls' );
				if ( $sModeTo == 'csv' && !empty( $aOptions['Delimiter'] ) ) {
					// Remove csv Delimiter from cell content or it counts as
					// new cell
					$sVal = str_replace( $aOptions['Delimiter'], '', $sVal );
				}
				$cell->setValue( trim( $sVal ) );
			}
		}
		// Remove first row, when it is empty for unknown reason
		if ( $bFirstRowEmptyForNoReason ) {
			$spreadsheet->getActiveSheet()->removeRow( 1, 1 );
		}
	}

	/**
	 * Implementation of BsUniversalExportModule interface. Creates an overview
	 * over the ExportModule
	 * @return ViewExportModuleOverview
	 */
	public function getOverview() {
		$oModuleOverviewView = new ViewExportModuleOverview();
		$oModuleOverviewView->setOption(
			'module-title',
			wfMessage( 'bs-uemoduletable2excel-overview-title' )->text()
		);
		$oModuleOverviewView->setOption(
			'module-description',
			wfMessage( 'bs-uemoduletable2excel-overview-description' )->text()
		);
		$oModuleOverviewView->setOption(
			'module-bodycontent',
			wfMessage( 'bs-uemoduletable2excel-overview-bodycontent' )->text() . '<br/>'
		);

		return $oModuleOverviewView;
	}

	// </editor-fold>
}
