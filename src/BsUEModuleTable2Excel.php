<?php

namespace BlueSpice\UEModuleTable2Excel;

use BlueSpice\VisualEditorConnector\ColorMapper;
use BsFileSystemHelper;
use DOMDocument;
use DOMElement;
use DOMNode;
use Exception;
use MediaWiki\Config\ConfigFactory;
use MediaWiki\Context\RequestContext;
use MediaWiki\HookContainer\HookContainer;
use MediaWiki\Language\FormatterFactory;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Html;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use Wikimedia\Mime\MimeAnalyzer;

class BsUEModuleTable2Excel {

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

	/** @var ConfigFactory */
	protected $configFactory;

	/** @var HookContainer */
	protected $hookContainer;

	/** @var MimeAnalyzer */
	protected $mimeAnalyzer;

	/** @var FormatterFactory */
	protected $formatterFactory;

	public function __construct(
		ConfigFactory $configFactory, HookContainer $hookContainer,
		MimeAnalyzer $mimeAnalyer, FormatterFactory $formatterFactory
	) {
		$this->configFactory = $configFactory;
		$this->hookContainer = $hookContainer;
		$this->mimeAnalyzer = $mimeAnalyer;
		$this->formatterFactory = $formatterFactory;
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
	 * @param string $content
	 * @return array array(
	 *     'mime-type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	 *     'filename' => 'Filename.docx',
	 *     'content' => '8F3BC3025A7...'
	 * );
	 * @throws ConfigException
	 * @throws FatalError
	 * @throws Exception
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function createExportFile( ExportSpecification &$specification, string $content ) {
		$response = [];
		$modeFrom = $specification->getModeFrom();
		$modeTo = $specification->getModeTo();
		$formatter = $this->formatterFactory->getStatusFormatter( RequestContext::getMain() );

		$options = $this->getDefaultOptions();

		$tmpFileName = time();

		if ( $modeFrom == 'html' ) {
			$content = $this->prepareHTML( $content );
		}

		$this->hookContainer->run(
			'BsUEModuleTable2ExcelBeforeProcess',
			[
				$this,
				&$modeFrom,
				&$modeTo,
				&$content,
				&$tmpFileName
			]
		);

		$status = BsFileSystemHelper::saveToCacheDirectory(
			"$tmpFileName.$modeFrom",
			$content,
			'UEModuleTable2Excel'
		);

		if ( !$status->isGood() ) {
			throw new Exception( $formatter->getMessage( $status ) );
		}

		try {
			$reader = IOFactory::createReaderForFile(
				"{$status->getValue()}/$tmpFileName.$modeFrom"
			);
		} catch ( Exception $e ) {
			throw new Exception( "Unsupported mode $modeFrom" );
		}

		$reader instanceof Html;
		// apply reader options
		foreach ( $options as $key => $val ) {
			$sFName = "set$key";
			if ( !method_exists( $reader, $sFName ) ) {
				continue;
			}
			call_user_func(
				[ $reader, $sFName ],
				$val
			);
		}
		$spreadsheet = $reader->load(
			"{$status->getValue()}/$tmpFileName.$modeFrom"
		);

		$properties = $spreadsheet->getProperties();

		// apply properties
		foreach ( $options as $key => $val ) {
			$sFName = "set$key";
			if ( !method_exists( $properties, $sFName ) ) {
				continue;
			}
			call_user_func(
				[ $properties, $sFName ],
				$val
			);
		}

		$this->adjustCellContent( $spreadsheet, $specification, $options );

		$writer = IOFactory::createWriter( $spreadsheet, ucfirst( $modeTo ) );

		// apply writer options
		foreach ( $options as $key => $val ) {
			$sFName = "set$key";
			if ( !method_exists( $writer, $sFName ) ) {
				continue;
			}
			call_user_func(
				[ $writer, $sFName ],
				$val
			);
		}

		$writer->save( "{$status->getValue()}/$tmpFileName.$modeTo" );
		$sMimeType = $this->mimeAnalyzer
			->getMimeTypeFromExtensionOrNull( $modeTo ) ?? MEDIATYPE_UNKNOWN;

		$response['filename'] = "$tmpFileName.$modeTo";
		$response['mime-type'] = $sMimeType;

		$file = BsFileSystemHelper::getCacheFileContent(
			"$tmpFileName.$modeTo",
			'UEModuleTable2Excel'
		);

		if ( !$file->isGood() ) {
			throw new Exception( $formatter->getMessage( $file ) );
		}

		$response['content'] = $file->getValue();

		$config = $this->configFactory->makeConfig( 'bsg' );

		if ( !$config->get( 'TestMode' ) ) {
			unlink( "{$status->getValue()}/$tmpFileName.$modeTo" );
			unlink( "{$status->getValue()}/$tmpFileName.$modeFrom" );
		}
		return $response;
	}

	/**
	 *
	 * @param array $options
	 * @return array
	 */
	protected function getDefaultOptions( $options = [] ) {
		$options['Description'] = '';
		$options['Keywords'] = '';
		$options['Category'] = '';

		$options['Delimiter'] = ';';
		$options['Enclosure'] = '';
		$options['LineEnding'] = "\r\n";
		$options['InputEncoding'] = 'UTF-8';

		return $options;
	}

	/**
	 * @param string $content
	 * @return string
	 */
	protected function prepareHTML( $content ) {
		$path = str_replace( '$1', '', $GLOBALS['wgArticlePath'] );
		// replace relative links /1.23/index.php?title=Main_Page
		$content = preg_replace(
			'#href="(' . str_replace( '?', '.', $path ) . ')(.*?)"#s',
			'href="' . $GLOBALS['wgServer'] . $path . '$2"',
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
		$sitename = $GLOBALS['wgSitename'];
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
	 * @param array $options
	 */
	protected function adjustCellContent(
		Spreadsheet $spreadsheet, ExportSpecification $specs, $options
	) {
		$lastCol = $spreadsheet->getActiveSheet()->getHighestColumn();
		$lastRow = $spreadsheet->getActiveSheet()->getHighestRow();
		$firstRowEmptyForNoReason = true;

		$lastCol++;
		for ( $col = 'A'; $col != $lastCol; $col++ ) {
			for ( $row = 1; $row <= $lastRow; $row++ ) {
				$cell = $spreadsheet->getActiveSheet()->getCell( "$col$row" );
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
					$firstRowEmptyForNoReason = false;
					continue;
				}

				$val = $this->stripPotentialFormula( $val );
				if ( $row == 1 && !empty( $val ) ) {
					$firstRowEmptyForNoReason = false;
				}

				$modeTo = $specs->getParam( 'ModeTo', 'xls' );
				if ( $modeTo == 'csv' && !empty( $options['Delimiter'] ) ) {
					// Remove csv Delimiter from cell content or it counts as
					// new cell
					$val = str_replace( $options['Delimiter'], '', $val );
				}

				$val = trim( $val );
				$cell->setValue( $val );
			}
		}

		// Remove first row, when it is empty for unknown reason
		if ( $firstRowEmptyForNoReason ) {
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
}
