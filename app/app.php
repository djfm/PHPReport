<?php

require_once __DIR__.'/../vendor/autoload.php';

require_once __DIR__.'/lib/Report.php';

$app = new Silex\Application();

ini_set('display_errors','on');
$app['debug'] = true;


$app->register(new Silex\Provider\TwigServiceProvider(), array(
    'twig.path' => __DIR__.'/views',
));

$app->register(new Silex\Provider\UrlGeneratorServiceProvider());

$app->get('/{rep}/html', function ($rep) use ($app){
	$report = new Report;
	$report->load(__DIR__ . "/reports/$rep.xml");
	
	$report->compute();

	return $report->renderAsHTML();
})->bind('html');

$app->get('/{rep}/excel', function ($rep) use ($app){
	$report = new Report;
	$report->load(__DIR__ . "/reports/$rep.xml");
	
	$report->compute();
	$report->sendAsExcel();
})->bind('excel');

$app->get('/', function() use ($app){

	$reports = array_map(function($p){return basename($p,'.xml');}, array_filter(scandir(__DIR__ . '/reports/'),function($fname){return $fname[0] != '.' and preg_match('/\.xml$/', $fname);}));

	return $app['twig']->render('index.html.twig', array('reports' => $reports));
})->bind('home');

$app->run();