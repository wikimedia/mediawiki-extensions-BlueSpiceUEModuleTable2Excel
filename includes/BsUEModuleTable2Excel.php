<?php

use BlueSpice\UniversalExport\ExportModule;
use BlueSpice\UniversalExport\ExportSpecification;
use MediaWiki\MediaWikiServices;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Html;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class BsUEModuleTable2Excel extends ExportModule {

	/**
	 * Implementation of BsUniversalExportModule interface.
	 * @param ExportSpecification &$specification
	 * @return array array(
	 *     'mime-type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	 *     'filename' => 'Filename.docx',
	 *     'content' => '8F3BC3025A7...'
	 * );
	 * @throws ConfigException
	 * @throws FatalError
	 * @throws MWException
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function createExportFile( ExportSpecification &$specification ) {
		$aResponse = [];
		$sModeFrom = strtolower( RequestContext::getMain()->getRequest()->getVal(
			'ModeFrom',
			'html'
		) );
		$sModeTo = strtolower( RequestContext::getMain()->getRequest()->getVal(
			'ModeTo',
			'xls'
		) );
		$sContent = RequestContext::getMain()->getRequest()->getVal( 'content', '' );

		$aOptions = array_merge_recursive(
			$this->getDefaultOptions( $specification ),
			RequestContext::getMain()->getRequest()->getArray( 'options', [] )
		);
		$tmpFileName = time();

		if ( $sModeFrom == 'html' ) {
			$sContent = $this->prepareHTML( $sContent );
		}

		$services = MediaWikiServices::getInstance();
		$services->getHookContainer()->run(
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
			throw new Exception( "Unsupported mode $sModeFrom" );
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

		$this->adjustCellContent( $spreadsheet, $specification, $aOptions );

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
		$sMimeType = $services->getMimeAnalyzer()
			->getMimeTypeFromExtensionOrNull( $sModeTo ) ?? MEDIATYPE_UNKNOWN;

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

		$config = $services->getConfigFactory()->makeConfig( 'bsg' );

		if ( !$config->get( 'TestMode' ) ) {
			unlink( "{$oStatus->getValue()}/$tmpFileName.$sModeTo" );
			unlink( "{$oStatus->getValue()}/$tmpFileName.$sModeFrom" );
		}
		return $aResponse;
	}

	/**
	 *
	 * @param ExportSpecification $specification
	 * @param array $aOptions
	 * @return array
	 */
	protected function getDefaultOptions( ExportSpecification $specification, $aOptions = [] ) {
		$oTitle = $specification->getTitle();

		if ( $oTitle->getNamespace() < 0 ) {
			$aOptions['Creator'] = $GLOBALS['wgSitename'];
			$aOptions['LastModifiedBy'] = $GLOBALS['wgSitename'];
		} else {
			$services = MediaWikiServices::getInstance();
			$oWikiPage = $services->getWikiPageFactory()->newFromTitle( $oTitle );
			$util = $services->getService( 'BSUtilityFactory' );

			$user = $services->getUserFactory()->newFromId( $oWikiPage->getCreator()->getId() );
			$aOptions['Creator'] = $util->getUserHelper(
				$user
			)->getDisplayName();

			if ( $oWikiPage->getRevisionRecord() ) {
				$lastEditorId = $oWikiPage->getRevisionRecord()->getUser()->getId();
				$lastEditor = $services->getUserFactory()->newFromId( $lastEditorId );
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
		$sContent = preg_replace_callback(
			'#<img(.*?)/?>#',
			static function ( $matches ) {
				$altText = '';

				$attribs = [];
				preg_match_all( '#\s*(.*?)=\"(.*?)\"\s*#', $matches[1], $attribs );

				if ( !empty( $attribs ) ) {
					$index = array_search( 'alt', $attribs[1] );
					if ( $index !== false ) {
						$altText = $attribs[2][$index];
					}
				}
				return $altText;
			},
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
	 * @param ExportSpecification $specs
	 * @param array $aOptions
	 */
	protected function adjustCellContent(
		Spreadsheet $spreadsheet, ExportSpecification $specs, $aOptions
	) {
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
				$sVal = $this->stripPotentialFormula( $sVal );
				if ( $iRow == 1 && !empty( $sVal ) ) {
					$bFirstRowEmptyForNoReason = false;
				}
				$sModeTo = $specs->getParam( 'ModeTo', 'xls' );
				if ( $sModeTo == 'csv' && !empty( $aOptions['Delimiter'] ) ) {
					// Remove csv Delimiter from cell content or it counts as
					// new cell
					$sVal = str_replace( $aOptions['Delimiter'], '', $sVal );
				}
				$sVal = trim( $sVal );
				$sVal = $this->stripPotentialFormula( $sVal );
				$cell->setValue( $sVal );
			}
		}
		// Remove first row, when it is empty for unknown reason
		if ( $bFirstRowEmptyForNoReason ) {
			$spreadsheet->getActiveSheet()->removeRow( 1, 1 );
		}
	}

	/**
	 *
	 * From https://phpspreadsheet.readthedocs.io/en/latest/topics/accessing-cells/#setting-a-formula-in-a-cell
	 * "if you store a string value with the first character an = in a cell.
	 *  PHPSpreadsheet will treat that value as a formula"
	 *
	 * If the formula is not valid this may lead to an exception.
	 * We currently do not support formulas at all, therefore we explicitly
	 * strip the leading "=".
	 *
	 * @param string $rawCellValue
	 * @return string
	 */
	private function stripPotentialFormula( $rawCellValue ) {
		return preg_replace( '#^=#', '', $rawCellValue );
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

	/**
	 * @inheritDoc
	 */
	public function getExportPermission() {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function getSubactionHandlers() {
		return null;
	}

	/**
	 * @inheritDoc
	 */
	public function getActionButtonDetails() {
		return null;
	}
}
