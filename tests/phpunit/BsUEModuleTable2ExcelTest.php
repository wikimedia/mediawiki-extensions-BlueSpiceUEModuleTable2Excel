<?php

namespace BlueSpice\UEModuleTable2Excel\Tests;

use BlueSpice\UEModuleTable2Excel\BsUEModuleTable2Excel;
use BlueSpice\VisualEditorConnector\ColorMapper;
use MediaWiki\MediaWikiServices;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PHPUnit\Framework\TestCase;

class BsUEModuleTable2ExcelTest extends TestCase {

	/** @var string */
	protected $bold = BsUEModuleTable2Excel::MARKER_BOLD;
	/** @var string */
	protected $italic = BsUEModuleTable2Excel::MARKER_ITALIC;
	/** @var string */
	protected $strikethrough = BsUEModuleTable2Excel::MARKER_STRIKETHROUGH;
	/** @var string */
	protected $underline = BsUEModuleTable2Excel::MARKER_UNDERLINE;
	/** @var string */
	protected const RED = ColorMapper::COLORS['col-red'];
	/** @var string */
	protected $red = BsUEModuleTable2Excel::MARKER_WRAPPER . self::RED . BsUEModuleTable2Excel::MARKER_WRAPPER;
	/** @var string */
	protected const BLUE = ColorMapper::COLORS['col-blue'];
	/** @var string */
	protected $blue = BsUEModuleTable2Excel::MARKER_WRAPPER . self::BLUE . BsUEModuleTable2Excel::MARKER_WRAPPER;

	/**
	 * @covers \BlueSpice\UEModuleTable2Excel\BsUEModuleTable2Excel::replaceStyleTags
	 * @dataProvider provideStyleTagsData
	 */
	public function testreplaceStyleTags( $content, $expected ) {
		$bsUEModuleTable2Excel = $this->createBsUEModuleTable2Excel();

		$result = $bsUEModuleTable2Excel->replaceStyleTags( $content );

		$this->assertSame( $expected, $result );
	}

	/**
	 * @covers \BlueSpice\UEModuleTable2Excel\BsUEModuleTable2Excel::replaceStyleMarkers
	 * @dataProvider provideStyleMarkerData
	 */
	public function testReplaceStyleMarkers( $value, $expected ) {
		$bsUEModuleTable2Excel = $this->createBsUEModuleTable2Excel();

		$richText = $bsUEModuleTable2Excel->replaceStyleMarkers( $value );

		$this->assertEquals( $expected, $richText );
	}

	/**
	 * Creates a BsUEModuleTable2Excel object.
	 *
	 * @return BsUEModuleTable2Excel The BsUEModuleTable2Excel object.
	 */
	private function createBsUEModuleTable2Excel() {
		$services = MediaWikiServices::getInstance();
		$configFactory = $services->getConfigFactory();
		$hookContainer = $services->getHookContainer();
		$mimeAnalyzer = $services->getMimeAnalyzer();
		$formatter = $services->getFormatterFactory();
		return new BsUEModuleTable2Excel( $configFactory, $hookContainer, $mimeAnalyzer, $formatter );
	}

	/**
	 * @return array[]
	 * phpcs:disable Generic.Files.LineLength.TooLong
	 */
	public function provideStyleTagsData() {
		return [
			'bold tag' => [
				<<<EOT
<table>

<tbody><tr>
<td><b>bold</b>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->bold}bold{$this->bold}
</td></tr></tbody></table>

EOT
			],
			'italic tag' => [
				<<<EOT
<table>

<tbody><tr>
<td><i>italic</i>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->italic}italic{$this->italic}
</td></tr></tbody></table>

EOT
			],
			'strikethrough tag' => [
				<<<EOT
<table>

<tbody><tr>
<td><s>strikethrough</s>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->strikethrough}strikethrough{$this->strikethrough}
</td></tr></tbody></table>

EOT
			],
			'underline tag' => [
				<<<EOT
<table>

<tbody><tr>
<td><u>underline</u>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->underline}underline{$this->underline}
</td></tr></tbody></table>

EOT
			],
			'combined tags' => [
				<<<EOT
<table>

<tbody><tr>
<td><b>bold</b> <i>italic</i> <s>strikethrough</s> <u>underline</u>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->bold}bold{$this->bold} {$this->italic}italic{$this->italic} {$this->strikethrough}strikethrough{$this->strikethrough} {$this->underline}underline{$this->underline}
</td></tr></tbody></table>

EOT
			],
			'mixed tags' => [
				<<<EOT
<table>

<tbody><tr>
<td><b><i><s><u>boldItalicStrikethroughUnderline</u></s></i></b>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->bold}{$this->italic}{$this->strikethrough}{$this->underline}boldItalicStrikethroughUnderline{$this->underline}{$this->strikethrough}{$this->italic}{$this->bold}
</td></tr></tbody></table>

EOT
			],
			'col-red class' => [
				<<<EOT
<table>

<tbody><tr>
<td><span class="col-red">col-red</span>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->red}col-red{$this->red}
</td></tr></tbody></table>

EOT
			],
			'col-blue class' => [
				<<<EOT
<table>

<tbody><tr>
<td><span class="col-blue">col-blue</span>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->blue}col-blue{$this->blue}
</td></tr></tbody></table>

EOT
			],
			'combined classes' => [
				<<<EOT
<table>

<tbody><tr>
<td><span class="col-red">col-red</span> <span class="col-blue">col-blue</span>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->red}col-red{$this->red} {$this->blue}col-blue{$this->blue}
</td></tr></tbody></table>

EOT
			],
			'combined classes and plain' => [
				<<<EOT
<table>

<tbody><tr>
<td><span class="col-red">col-red</span> plain <span class="col-blue">col-blue</span>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->red}col-red{$this->red} plain {$this->blue}col-blue{$this->blue}
</td></tr></tbody></table>

EOT
			],
			'tags, classes and plain' => [
				<<<EOT
<table>

<tbody><tr>
<td><i><b><span class="col-red">bolditalicred</span></b></i> plain <s><u><span class="col-blue">strikethroughunderlineblue</span></u></s>
</td></tr></tbody></table>
EOT
				,
				<<<EOT
<?xml encoding="UTF-8"><table>

<tbody><tr>
<td>{$this->italic}{$this->bold}{$this->red}bolditalicred{$this->red}{$this->bold}{$this->italic} plain {$this->strikethrough}{$this->underline}{$this->blue}strikethroughunderlineblue{$this->blue}{$this->underline}{$this->strikethrough}
</td></tr></tbody></table>

EOT
			]
		];
	}

	/**
	 * Data provider for testReplaceStyleMarkers.
	 * phpcs:disable Generic.Files.LineLength.TooLong
	 *
	 * @return array[]
	 */
	public function provideStyleMarkerData() {
		return [
			'bold marker' => [
				"{$this->bold}bold{$this->bold}",
				$this->createRichText( [
					'bold' => [
						'bold' => true
					]
				] )
			],
			'italic marker' => [
				"{$this->italic}italic{$this->italic}",
				$this->createRichText( [
					'italic' => [
						'italic' => true
					]
				] )
			],
			'strikethrough marker' => [
				"{$this->strikethrough}strikethrough{$this->strikethrough}",
				$this->createRichText( [
					'strikethrough' => [
						'strikethrough' => true
					]
				] )
			],
			'underline marker' => [
				"{$this->underline}underline{$this->underline}",
				$this->createRichText( [
					'underline' => [
						'underline' => true
					]
				] )
			],
			'combined markers' => [
				"{$this->bold}bold{$this->bold}{$this->italic}italic{$this->italic}{$this->strikethrough}strikethrough{$this->strikethrough}{$this->underline}underline{$this->underline}",
				$this->createRichText( [
					'bold' => [
						'bold' => true
					],
					'italic' => [
						'italic' => true
					],
					'strikethrough' => [
						'strikethrough' => true
					],
					'underline' => [
						'underline' => true
					]
				] )
			],
			'mixed markers' => [
				"{$this->bold}{$this->italic}{$this->strikethrough}{$this->underline}bold italic strikethrough underline{$this->underline}{$this->strikethrough}{$this->italic}{$this->bold}",
				$this->createRichText( [
					'bold italic strikethrough underline' => [
						'bold' => true,
						'italic' => true,
						'strikethrough' => true,
						'underline' => true
					]
				] )
			],
			'col-red marker' => [
				"{$this->red}col-red{$this->red}",
				$this->createRichText( [
					'col-red' => [
						'color' => self::RED
					]
				] )
			],
			'col-blue marker' => [
				"{$this->blue}col-blue{$this->blue}",
				$this->createRichText( [
					'col-blue' => [
						'color' => self::BLUE
					]
				] )
			],
			'combined color markers' => [
				"{$this->red}col-red{$this->red}{$this->blue}col-blue{$this->blue}",
				$this->createRichText( [
					'col-red' => [
						'color' => self::RED
					],
					'col-blue' => [
						'color' => self::BLUE
					]
				] )
			],
			'combined color markers and plain' => [
				"{$this->red}col-red{$this->red} plain {$this->blue}col-blue{$this->blue}",
				$this->createRichText( [
					'col-red' => [
						'color' => self::RED
					],
					' plain ' => [],
					'col-blue' => [
						'color' => self::BLUE
					]
				] )
			],
			'mixed markers and plain' => [
				"{$this->bold}{$this->italic}{$this->red}bold italic red{$this->red}{$this->italic}{$this->bold} plain {$this->strikethrough}{$this->underline}{$this->blue}strikethrough underline blue{$this->blue}{$this->underline}{$this->strikethrough}",
				$this->createRichText( [
					'bold italic red' => [
						'bold' => true,
						'italic' => true,
						'color' => self::RED
					],
					' plain ' => [],
					'strikethrough underline blue' => [
						'strikethrough' => true,
						'underline' => true,
						'color' => self::BLUE
					]
				] )
			],
		];
	}

	/**
	 * Creates a RichText object with styled text segments.
	 *
	 * @param array $segments An associative array of text segments and their style flags
	 * @return RichText The RichText object containing the styled segments.
	 */
	private function createRichText( array $segments ) {
		$richText = new RichText();

		foreach ( $segments as $text => $style ) {
			$textRun = $richText->createTextRun( $text );
			$font = $textRun->getFont();

			$font->setBold( $style['bold'] ?? false );
			$font->setItalic( $style['italic'] ?? false );
			$font->setStrikethrough( $style['strikethrough'] ?? false );
			$font->setUnderline( $style['underline'] ?? false );

			if ( isset( $style['color'] ) ) {
				$font->setColor( new Color( "FF{$style['color']}" ) );
			}
		}

		return $richText;
	}
}
