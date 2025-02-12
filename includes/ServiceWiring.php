<?php

use BlueSpice\UEModuleTable2Excel\BsUEModuleTable2Excel;
use MediaWiki\MediaWikiServices;

return [
	'BlueSpiceUEModuleTable2Excel' => static function ( MediaWikiServices $services ) {
		return new BsUEModuleTable2Excel(
			$services->getConfigFactory(),
			$services->getHookContainer(),
			$services->getMimeAnalyzer(),
			$services->getFormatterFactory()
		);
	}
];
