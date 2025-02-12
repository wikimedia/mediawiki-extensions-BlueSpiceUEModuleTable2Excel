<?php

namespace BlueSpice\UEModuleTable2Excel;

use MediaWiki\Config\Config;
use MediaWiki\User\User;

class ExportSpecification {

	/** @var Config */
	private $config;
	/** @var User */
	private $user = null;
	/** @var string */
	private $modeFrom;
	/** @var string */
	private $modeTo;
	/** @var array */
	private $params;

	/**
	 *
	 * @param Config $config
	 * @param User $user
	 * @param string $modeFrom
	 * @param string $modeTo
	 * @param array $params
	 */
	public function __construct(
		Config $config, User $user, string $modeFrom, string $modeTo, $params = []
	) {
		$this->config = $config;
		$this->user = $user;
		$this->modeFrom = $modeFrom;
		$this->modeTo = $modeTo;
		$this->params = $params;

		$this->setDefaults();
	}

	/**
	 * @return User|null
	 */
	public function getUser() {
		return $this->user;
	}

	/**
	 * @return string
	 */
	public function getModeFrom() {
		return $this->modeFrom;
	}

	/**
	 * @return string
	 */
	public function getModeTo() {
		return $this->modeTo;
	}

	/**
	 * @return array
	 */
	public function getParams() {
		return $this->params;
	}

	/**
	 * @param string $param
	 * @param null $default
	 * @return mixed
	 */
	public function getParam( $param, $default = null ) {
		if ( isset( $this->params[$param] ) ) {
			return $this->params[$param];
		}

		return $default;
	}

	/**
	 * @param string $param
	 * @param mixed $value
	 */
	public function setParam( $param, $value ) {
		$this->params[$param] = $value;
	}

	private function setDefaults() {
		// Set up default parameters and metadata
		$this->params = [];

		$this->setParam( 'webroot-filesystempath', $this->config->get( 'UploadDirectory' ) );
	}
}
