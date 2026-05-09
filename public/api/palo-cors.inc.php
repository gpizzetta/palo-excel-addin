<?php

declare(strict_types=1);

/**
 * Schéma public (https derrière proxy TLS, etc.).
 */
function paloaddin_request_scheme(): string
{
    if (!empty($_SERVER['HTTP_X_FORWARDED_PROTO'])) {
        $p = strtolower(trim(explode(',', (string) $_SERVER['HTTP_X_FORWARDED_PROTO'])[0]));

        return $p === 'https' ? 'https' : 'http';
    }
    if (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off') {
        return 'https';
    }
    if (isset($_SERVER['SERVER_PORT']) && (string) $_SERVER['SERVER_PORT'] === '443') {
        return 'https';
    }

    return 'http';
}

/**
 * Hôte public (Host ou X-Forwarded-Host derrière reverse proxy).
 */
function paloaddin_request_host(): string
{
    if (!empty($_SERVER['HTTP_X_FORWARDED_HOST'])) {
        return trim(explode(',', (string) $_SERVER['HTTP_X_FORWARDED_HOST'])[0]);
    }
    $h = $_SERVER['HTTP_HOST'] ?? '';

    return trim((string) $h);
}

/**
 * Origine attendue pour ce vhost (ex. https://portal.berdoz.local).
 */
function paloaddin_cors_site_origin(): string
{
    $host = paloaddin_request_host();
    if ($host === '') {
        return '';
    }

    return paloaddin_request_scheme() . '://' . $host;
}

function paloaddin_cors_origin_is_this_site(string $origin): bool
{
    if ($origin === '') {
        return false;
    }
    $expected = paloaddin_cors_site_origin();
    if ($expected === '') {
        return false;
    }
    if (strcasecmp($origin, $expected) === 0) {
        return true;
    }
    $host = paloaddin_request_host();
    $scheme = paloaddin_request_scheme();
    if ($scheme === 'https' && strcasecmp($origin, 'https://' . preg_replace('/:443$/', '', $host)) === 0) {
        return true;
    }
    if ($scheme === 'http' && strcasecmp($origin, 'http://' . preg_replace('/:80$/', '', $host)) === 0) {
        return true;
    }

    return false;
}

function paloaddin_api_cors_allowed(string $o): bool
{
    if ($o === '') {
        return false;
    }
    if (paloaddin_cors_origin_is_this_site($o)) {
        return true;
    }
    if (preg_match('#^https://palo\.berdoz\.local$#i', $o) === 1) {
        return true;
    }
    if (preg_match('#^http://palo\.berdoz\.local$#i', $o) === 1) {
        return true;
    }
    if (preg_match('#^https://([a-z0-9-]+\.)?officeapps\.live\.com$#i', $o) === 1) {
        return true;
    }

    return false;
}

function paloaddin_api_cors_send(): void
{
    $origin = $_SERVER['HTTP_ORIGIN'] ?? '';
    if (paloaddin_api_cors_allowed($origin)) {
        header('Access-Control-Allow-Origin: ' . $origin);
        header('Access-Control-Allow-Credentials: true');
        header('Vary: Origin');
    }
    header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
    header('Access-Control-Allow-Headers: Content-Type, Authorization, X-Requested-With');
}
