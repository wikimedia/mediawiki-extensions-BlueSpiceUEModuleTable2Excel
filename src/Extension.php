<?php

/**
 * BlueSpiceExportTables extension for BlueSpice
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, version 3.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License along
 * with this program; if not, write to the Free Software Foundation, Inc.,
 * 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.
 *
 * This file is part of BlueSpice MediaWiki
 * For further information visit https://bluespice.com
 *
 * @author     Patric Wirth
 * @package    BluespiceExportTables
 * @subpackage ExportTables
 * @copyright  Copyright (C) 2016 Hallo Welt! GmbH, All rights reserved.
 * @license    http://www.gnu.org/copyleft/gpl.html GPL-3.0-only
 * @filesource
 */

namespace BlueSpice\UEModuleTable2Excel;
/**
 * Base class for BlueSpiceExportTables extension
 * @package BlueSpiceExportTables
 * @subpackage ExportTables
 */
class Extension extends \BlueSpice\Extension {

	/**
	 *
	 * @param SpecialUniversalExport $oSpecialPage
	 * @param string $sParam
	 * @param array $aModules
	 * @return true
	 */
	public static function onBSUniversalExportSpecialPageExecute( $oSpecialPage, $sParam, &$aModules ) {
		$aModules['table2excel'] = new \BsUEModuleTable2Excel();
		return true;
	}
}