<?php

namespace BlueSpice\UEModuleTable2Excel\Rest;

use BlueSpice\UEModuleTable2Excel\BsUEModuleTable2Excel;
use BlueSpice\UEModuleTable2Excel\ExportSpecification;
use MediaWiki\Config\Config;
use MediaWiki\Context\RequestContext;
use MediaWiki\Rest\SimpleHandler;
use Wikimedia\ParamValidator\ParamValidator;

class Table2Excel extends SimpleHandler {

	/** @var Config */
	private $config;

	/** @var BsUEModuleTable2Excel */
	private $module;

	public function __construct( Config $config, BsUEModuleTable2Excel $module ) {
		$this->config = $config;
		$this->module = $module;
	}

	public function run() {
		$validated = $this->getValidatedParams();
		$modeFrom = 'html';
		$modeTo = $validated['modeTo'];
		$user = RequestContext::getMain()->getUser();

		$body = $this->getValidatedBody();
		$data = $body['data'];
		if ( !isset( $data['content'] ) ) {
			return $this->getResponseFactory()->createHttpError( 404, [ 'No export content set' ] );
		}
		$content = $body['data']['content'];
		if ( !$content ) {
			return $this->getResponseFactory()->createHttpError( 404, [ 'No data found' ] );
		}
		$specs = new ExportSpecification( $this->config, $user, $modeFrom, $modeTo );
		$export = $this->module->createExportFile( $specs, $content );
		$filename = $export['filename'];
		$data = $export['content'];

		$response = $this->getResponseFactory()->create();
		$response->setHeader( 'Content-Type', $export['mime-type'] );
		$response->setHeader( 'Content-Disposition', 'attachment; filename="' . $filename . '"' );
		$response->setHeader( 'X-Filename', $filename );
		$response->getBody()->write( $data );

		return $response;
	}

	/**
	 * @inheritDoc
	 */
	public function getParamSettings() {
		return [
			'modeTo' => [
				self::PARAM_SOURCE => 'path',
				ParamValidator::PARAM_TYPE => 'string',
				ParamValidator::PARAM_REQUIRED => true
			]
		];
	}

	/**
	 * @inheritDoc
	 */
	public function getBodyParamSettings(): array {
		return [
			'data' => [
				static::PARAM_SOURCE => 'body',
				ParamValidator::PARAM_TYPE => 'array',
				ParamValidator::PARAM_REQUIRED => false
			]
		];
	}
}
