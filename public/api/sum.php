<?php

header('Content-Type: application/json; charset=utf-8');

$raw = file_get_contents('php://input');
$payload = json_decode($raw, true);

$a = 0.0;
$b = 0.0;
if (is_array($payload)) {
    $a = (float)($payload['a'] ?? 0);
    $b = (float)($payload['b'] ?? 0);
}

echo json_encode([
    'status' => 'ok',
    'a' => $a,
    'b' => $b,
    'sum' => $a + $b,
]);

