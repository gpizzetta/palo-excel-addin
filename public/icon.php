<?php

// Simple transparent PNG placeholder for Office manifest icons.
// 1x1 PNG scaled by Office, good enough for local testing.
header('Content-Type: image/png');
echo base64_decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO5Vn6sAAAAASUVORK5CYII='
);

