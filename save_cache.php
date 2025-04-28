<?php
$data = file_get_contents("php://input");

file_put_contents("geo_cache.json", $data);

echo "Cache enregistré avec succès.";
?>
