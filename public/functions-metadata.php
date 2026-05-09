<?php

declare(strict_types=1);

/**
 * Sert functions.json. Les en-têtes CORS viennent de public/.htaccess (mod_headers) ;
 * ne pas les renvoyer ici : paloaddin_api_cors_send() doublonnait Access-Control-Allow-Origin
 * (Apache + PHP), ce que les navigateurs rejettent.
 */
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

$path = __DIR__ . '/functions.json';
if (!is_readable($path)) {
    http_response_code(500);
    header('Content-Type: text/plain; charset=utf-8');
    echo 'functions.json missing';
    exit;
}

header('Content-Type: application/json; charset=utf-8');
readfile($path);
