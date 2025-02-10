<?php

use BlueSpice\UniversalExport\ExportModule;
use BlueSpice\UniversalExport\ExportSpecification;
use BlueSpice\VisualEditorConnector\ColorMapper;
use MediaWiki\Config\ConfigException;
use MediaWiki\Context\RequestContext;
use MediaWiki\MediaWikiServices;
use MediaWiki\Request\WebRequest;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Html;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;

class BsUEModuleTable2Excel extends ExportModule {

	public const MARKER_WRAPPER = '%%';
	public const TAG_BOLD = 'b';
	public const TAG_ITALIC = 'i';
	public const TAG_STRIKETHROUGH = 's';
	public const TAG_UNDERLINE = 'u';

	public const MARKER_BOLD = self::MARKER_WRAPPER . self::TAG_BOLD . self::MARKER_WRAPPER;
	public const MARKER_ITALIC = self::MARKER_WRAPPER . self::TAG_ITALIC . self::MARKER_WRAPPER;
	public const MARKER_STRIKETHROUGH = self::MARKER_WRAPPER . self::TAG_STRIKETHROUGH . self::MARKER_WRAPPER;
	public const MARKER_UNDERLINE = self::MARKER_WRAPPER . self::TAG_UNDERLINE . self::MARKER_WRAPPER;

	/** @var array */
	protected array $colorMarkers = [];

	/**
	 * @inheritDoc
	 */
	public function __construct(
		$name, MediaWikiServices $services, Config $config, WebRequest $request
	) {
		parent::__construct( $name, $services, $config, $request );
		$this->buildColorMarkers();
	}

	private function buildColorMarkers() {
		foreach ( ColorMapper::COLORS as $class => $code ) {
			$this->colorMarkers[$class] = self::MARKER_WRAPPER . $code . self::MARKER_WRAPPER;
		}
	}

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
	 * @param string $content
	 * @return string
	 */
	protected function prepareHTML( $content ) {
		global $wgArticlePath, $wgServer, $wgSitename;

		$sPath = str_replace( '$1', '', $wgArticlePath );
		// replace relative links /1.23/index.php?title=Main_Page
		$content = preg_replace(
			'#href="(' . str_replace( '?', '.', $sPath ) . ')(.*?)"#s',
			'href="' . $wgServer . $sPath . '$2"',
			$content
		);

		$content = $this->replaceImagesWithAltText( $content );
		$content = $this->replaceStyleTags( $content );

		$content = htmlspecialchars( $content, ENT_HTML401, 'UTF-8' );
		$content = html_entity_decode( $content, ENT_HTML401, 'UTF-8' );
		$content = str_replace( '&nbsp;', '', $content );
		$content = preg_replace(
			"/(^[\r\n]*|[\r\n]+)[\s\t]*[\r\n]+/",
			"\n",
			$content
		);
		$content = preg_replace( "/<thead>\s*<\/thead>/", "", $content );
		$docType = '-//W3C//DTD HTML 4.01 Transitional//EN';

		// A sitename with more than 31 characters breaks the export
		$sitename = $wgSitename;
		if ( strlen( $sitename ) > 31 ) {
			$sitename = substr( $sitename, 0, 31 );
		}

		return <<<EOT
<!DOCTYPE HTML PUBLIC "$docType" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8">
		<title>$sitename</title>
	</head>
	<body>
		$content
	</body>
</html>
EOT;
	}

	/**
	 * Replaces image tags with their alt text
	 *
	 * @param string $content
	 * @return string
	 */
	private function replaceImagesWithAltText( $content ) {
		return preg_replace_callback(
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
			$content
		) ?: $content;
	}

	/**
	 * @param string $content
	 * @return string
	 */
	public function replaceStyleTags( string $content ): string {
		$dom = new DOMDocument();
		// phpcs:ignore Generic.PHP.NoSilencedErrors.Discouraged
		$success = @$dom->loadHTML( $content, LIBXML_HTML_NOIMPLIED | LIBXML_HTML_NODEFDTD );
		if ( !$success ) {
			return $content;
		}

		$tdElements = $dom->getElementsByTagName( 'td' );

		foreach ( $tdElements as $td ) {
			// Clear the existing content
			$innerHTML = '';
			foreach ( $td->childNodes as $child ) {
				$innerHTML .= $this->processNode( $child );
			}

			// Set the new content with markers
			$td->nodeValue = '';
			$td->appendChild( $dom->createTextNode( $innerHTML ) );
		}

		return $dom->saveHTML() ?: $content;
	}

	/**
	 * @param DOMNode $node
	 * @return string
	 */
	private function processNode( DOMNode $node ): string {
		if ( $node->nodeType === XML_TEXT_NODE ) {
			return $node->textContent;
		}

		$output = '';

		if ( $node->nodeType === XML_ELEMENT_NODE ) {
			/** @var DOMElement $node */
			$tagName = $node->tagName;

			// Recursively process child nodes
			foreach ( $node->childNodes as $child ) {
				$output .= $this->processNode( $child );
			}

			// Wrap output with markers based on the tag
			switch ( $tagName ) {
				case self::TAG_BOLD:
					$output = self::MARKER_BOLD . $output . self::MARKER_BOLD;
					break;
				case self::TAG_ITALIC:
					$output = self::MARKER_ITALIC . $output . self::MARKER_ITALIC;
					break;
				case self::TAG_STRIKETHROUGH:
					$output = self::MARKER_STRIKETHROUGH . $output . self::MARKER_STRIKETHROUGH;
					break;
				case self::TAG_UNDERLINE:
					$output = self::MARKER_UNDERLINE . $output . self::MARKER_UNDERLINE;
					break;
				case 'span':
					// Check for specific color class in the span element
					$output = $this->replaceColorSpans( $node, $output );
					break;
			}
		}

		return $output;
	}

	/**
	 * @param DOMElement $node
	 * @param string $output
	 * @return string
	 */
	public function replaceColorSpans( DOMElement $node, string $output ): string {
		if ( $node->hasAttribute( 'class' ) ) {
			$spanClass = $node->getAttribute( 'class' );

			// Check if the class exists in ColorMapper::COLORS
			if ( array_key_exists( $spanClass, ColorMapper::COLORS ) ) {
				$colorMarker = $this->colorMarkers[$spanClass];
				return $colorMarker . $output . $colorMarker;
			}
		}

		return $output;
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
				$cell = $spreadsheet->getActiveSheet()->getCell( "$sCol$iRow" );
				$val = trim( $cell->getValue() );

				// Replace markers if present
				if ( str_contains( $val, self::MARKER_BOLD ) ||
					str_contains( $val, self::MARKER_ITALIC ) ||
					str_contains( $val, self::MARKER_STRIKETHROUGH ) ||
					str_contains( $val, self::MARKER_UNDERLINE ) ||
					$this->stringInArray( $val, array_values( $this->colorMarkers ) )
				) {
					$richText = $this->replaceStyleMarkers( $val );
					$cell->setValue( $richText );
					$bFirstRowEmptyForNoReason = false;
					continue;
				}

				$val = $this->stripPotentialFormula( $val );
				if ( $iRow == 1 && !empty( $val ) ) {
					$bFirstRowEmptyForNoReason = false;
				}

				$sModeTo = $specs->getParam( 'ModeTo', 'xls' );
				if ( $sModeTo == 'csv' && !empty( $aOptions['Delimiter'] ) ) {
					// Remove csv Delimiter from cell content or it counts as
					// new cell
					$val = str_replace( $aOptions['Delimiter'], '', $val );
				}

				$val = trim( $val );
				$cell->setValue( $val );
			}
		}

		// Remove first row, when it is empty for unknown reason
		if ( $bFirstRowEmptyForNoReason ) {
			$spreadsheet->getActiveSheet()->removeRow( 1, 1 );
		}
	}

	/**
	 * @param string $value
	 * @param array $colorMarkers
	 * @return bool
	 */
	private function stringInArray( string $value, array $colorMarkers ): bool {
		foreach ( $colorMarkers as $marker ) {
			if ( str_contains( $value, $marker ) ) {
				return true;
			}
		}

		return false;
	}

	/**
	 * Replace bold, italic, strikethrough, underline and color markers in the string and return a RichText object
	 *
	 * @param string $value The string containing bold, italic, strikethrough, underline and color markers
	 * @return RichText
	 */
	public function replaceStyleMarkers( string $value ): RichText {
		$richText = new RichText();

		// Dynamically create a regex pattern for color markers
		$colorMarkers = array_values( $this->colorMarkers );
		$colorMarkersRegex = implode( '|', array_map( 'preg_quote', $colorMarkers ) );

		// Combine all markers into a regex pattern
		$b = preg_quote( self::MARKER_BOLD );
		$i = preg_quote( self::MARKER_ITALIC );
		$s = preg_quote( self::MARKER_STRIKETHROUGH );
		$u = preg_quote( self::MARKER_UNDERLINE );

		// Create the final regex pattern
		$pattern = "/($b|$i|$s|$u|$colorMarkersRegex)/";

		// Split the string by markers
		$segments = preg_split( $pattern, $value, -1, PREG_SPLIT_DELIM_CAPTURE | PREG_SPLIT_NO_EMPTY );

		// Array to track style states
		$styleStates = [
			'bold' => false,
			'italic' => false,
			'strikethrough' => false,
			'underline' => false,
			'color' => null
		];

		foreach ( $segments as $segment ) {
			if ( !$segment ) {
				continue;
			}
			if ( $this->isStyleMarker( $segment ) ) {
				$this->toggleStyleState( $styleStates, $segment );
				continue;
			}

			// Create a new TextRun for each segment
			$textRun = $richText->createTextRun( $segment );
			$font = $textRun->getFont();

			// Apply styles based on the current state
			$font->setBold( $styleStates['bold'] );
			$font->setItalic( $styleStates['italic'] );
			$font->setStrikethrough( $styleStates['strikethrough'] );
			$font->setUnderline( $styleStates['underline'] );

			if ( isset( $styleStates['color'] ) ) {
				$font->setColor( new Color( "FF{$styleStates['color']}" ) );
			}
		}

		return $richText;
	}

	/**
	 * Toggle the style states based on the segment
	 *
	 * @param array &$styleStates Current style states
	 * @param string $segment The segment containing the style marker
	 */
	public function toggleStyleState( array &$styleStates, string $segment ) {
		switch ( $segment ) {
			case self::MARKER_BOLD:
				$styleStates['bold'] = !$styleStates['bold'];
				break;
			case self::MARKER_ITALIC:
				$styleStates['italic'] = !$styleStates['italic'];
				break;
			case self::MARKER_STRIKETHROUGH:
				$styleStates['strikethrough'] = !$styleStates['strikethrough'];
				break;
			case self::MARKER_UNDERLINE:
				$styleStates['underline'] = !$styleStates['underline'];
				break;
			default:
				// Handle colors dynamically based on ColorMapper
				foreach ( $this->colorMarkers as $class => $marker ) {
					if ( $segment == $marker ) {
						$colorCode = ColorMapper::COLORS[$class];
						// If the current color is already set, unset it
						if ( isset( $styleStates['color'] ) &&
							$styleStates['color'] === $colorCode
						) {
							unset( $styleStates['color'] );
						} else {
							// Otherwise, set it to the new color
							$styleStates['color'] = $colorCode;
						}
						break;
					}
				}
				break;
		}
	}

	/**
	 * Check if the segment is a style marker
	 *
	 * @param string $segment The segment to check
	 * @return bool
	 */
	public function isStyleMarker( string $segment ): bool {
		// Combine static markers with dynamic color markers from $this->colorMarkers
		$allMarkers = array_merge(
			[
				self::MARKER_BOLD,
				self::MARKER_ITALIC,
				self::MARKER_STRIKETHROUGH,
				self::MARKER_UNDERLINE
			],
			array_values( $this->colorMarkers )
		);

		return in_array( $segment, $allMarkers );
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
