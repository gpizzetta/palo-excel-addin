<?php

header('Content-Type: application/json; charset=utf-8');

echo json_encode([
    'status' => 'ok',
    'service' => 'paloaddin-php',
    'time' => date('c'),
]);

