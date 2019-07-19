<?php
error_reporting(E_ALL & ~E_DEPRECATED & ~E_STRICT);
ini_set('display_errors', '1');
ini_set('log_errors', '0');
define('ROOT_PATH', dirname(__DIR__).DIRECTORY_SEPARATOR);

use Core\Git\Log;
use Core\Report\Export as ReportExport;

try {
    if (version_compare(PHP_VERSION, '7.1.0', '<')) {
        throw new Exception('low version of PHP is not supported,upgrade your PHP version >= 7.1.0 :(');
    }
    include_once ROOT_PATH.'vendor/autoload.php';
    //Check configuration
    $config = include_once ROOT_PATH.'configs/config.php';
    if (!isset($config['download_folder']) || !$config['download_folder']) {
        throw new Exception('PLZ,configure your download folder first :(');
    }
    if (!isset($config['repo']) || !is_array($config['repo'])) {
        throw new Exception('PLZ,configure your repo info first :(');
    }
    //Query available logs
    $outputs = [];
    foreach ($config['repo'] as $repo) {
        $generator = new Log([
            'limit' => $argv[1] ?? 10,
            'after' => $argv[2] ?? date('Y-m-d', strtotime('-1 week')),
            'identifier' => $config['identifier'] ?? '',
            'repo_path' => $repo['path'] ?? '',
            'project' => $repo['name'] ?? '',
            'auto_pull' => $config['auto_pull'] ?? true,
            'author' => $config['author'] ?? '',
        ]);
        $outputs = array_merge($outputs, $generator->getDone());
    }
    if (!$outputs) {
        throw new Exception('Sorry,I find nothing to write :(');
    }
    $distPath = ROOT_PATH.$config['download_folder'].DIRECTORY_SEPARATOR;
    if (!is_dir($distPath)) {
        mkdir($distPath);
    }
    //Generate report
    $title = sprintf($config['title'] ?? '', date('Y'));
    $reportExport = new ReportExport(ROOT_PATH.'assets'.DIRECTORY_SEPARATOR.'tpl.xlsx', $outputs);
    $reportExport->setTitle($title);
    $reportExport->setDepartment($config['department'] ?? '');
    $reportExport->setRealName($config['real_name'] ?? '');
    if (($config['debug'] ?? true)) {
        $reportExport->markdown();
    }
    $reportExport->save($distPath.$title.'-'.$config['real_name'].date('m.d').'.xlsx');
    echo PHP_EOL.'Your report has been created!find it in "'.$distPath.'",have a nice weekend :)'.PHP_EOL.PHP_EOL;
} catch (Exception $e) {
    echo $e->getMessage().PHP_EOL;
}

