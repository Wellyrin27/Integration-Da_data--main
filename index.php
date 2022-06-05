<?php

require 'vendor/autoload.php';
require 'Parser.php';

//Путь до xlsx файла из первого аргумента программы.
$inputFileName = $argv[1];

//Создание экземпляра класса parser с передачей в конструктор путь до xlsx файла.
$parser = new Parser($inputFileName);
//Получения списка адресов из xlsx файла.
$addresses = $parser->getAddresses();
//Получение GeoJson из DaData сервиса для каждого адреса из списка.
$result = $parser->getGeoData($addresses);

//Открытие файла json на запись.
$fp = fopen('results.json', 'w');
//Запись в файл полученного массива данных.
fwrite($fp, json_encode($result, JSON_UNESCAPED_UNICODE));
//закрытие файла.
fclose($fp);
