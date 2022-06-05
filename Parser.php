<?php

use PhpOffice\PhpSpreadsheet\IOFactory;


class Parser
{
    //Путь до xlsx файла.
    private string $path;
    
    //Адрес до сервиса DaData.
    private string $apiUrl = 'https://cleaner.dadata.ru/api/v1/clean';
    
    //Обработчик http запросов.
    private CurlHandle $handle;

    //Конструктор, принимающий путь до xlsx файла.
    public function __construct(string $path)
    {
        //Определяем переменную path значением из конструктора.
        $this->path = $path;

        //Считываем данные из файла config.php.
        $creds = require_once 'config.php';

        //Получение токена из файла config.php.
        $token = $creds['token'];

        //Получение секретного ключа из файла config.php.
        $secret = $creds['secret'];

        //Настройка обработчика http запросов.
        self::initDaData($token, $secret);
    }

    //Настройка обработчика http запросов.
    private function initDaData(string $token, string $secret)
    {
        //Инициализация обработчика http запросов.
        $this->handle = curl_init();

        //Указание параметра о необходимости возврата данных.
        curl_setopt($this->handle, CURLOPT_RETURNTRANSFER, 1);

        //Указание http заголовка.
        curl_setopt($this->handle, CURLOPT_HTTPHEADER, array(
            "Content-Type: application/json",
            "Accept: application/json",
            "Authorization: Token " . $token,
            "X-Secret: " . $secret,
        ));

        //Указание http метода отправки запроса (POST).
        curl_setopt($this->handle, CURLOPT_POST, 1);
    }

    //Метод для получения списка адресов из xlsx файла.
    public function getAddresses(): array
    {
        //Список считанных адресов из xlsx файла.
        $addressArray = [];
        
        //Получение "сырых" данных из xlsx файла.
        $xlsxData = $this->parseXlsx();

        //Адрес ячейки содержащей адрес объекта.
        $it = 6;

        //обработка значений из файла.
        for (; $it < count($xlsxData); $it++) {

            //Если указаная ячейка содержит адрес объекта.
            if (isset($xlsxData[$it][2])) {
                $addressArray[] = $xlsxData[$it][2];
            }
        }

        //Возвращает список адресов объектов.
        return $addressArray;
    }

    //Получение "сырых" данных из xlsx файла.
    private function parseXlsx(): array
    {
        try {
            //Получение пути до xlsx файла.
            $inputFileName = $this->path;

            //Открытие файла и получение страницы с данными.
            $inputFileType = IOFactory::identify($inputFileName);
            $reader = IOFactory::createReader($inputFileType);
            $spreadsheet = $reader->load($inputFileName);
            
            //Получение всех данных из активной xlsx страницы.
            return $spreadsheet->getActiveSheet()->toArray();
        } catch (Exception $e) {
            //В случае ошибки возвращается пустой массив данных.
            return [];
        }
    }

    //Получение GeoJson из сервиса DaData по указанному списку адресов.
    public function getGeoData(array $addresses): array
    {
        try {
            //формирование адреса запроса для получения данных.
            $url = $this->apiUrl . "/address";

            //Указание списка адресов.
            $fields = $addresses;

            //Вернуть полученный результат выполнения запроса.
            return $this->executeRequest($url, $fields);
        } catch (Exception $e) {
            //В случае ошибки возвращается пустой массив данных.
            return [];
        }
    }

    //Выполнение http запроса с указанными данными.
    private function executeRequest($url, $fields)
    {
        //Указание адреса запроса.
        curl_setopt($this->handle, CURLOPT_URL, $url);

        //Указание метода http запроса (POST).
        curl_setopt($this->handle, CURLOPT_POST, 1);

        //Указание тела запроса (в формате Json).
        curl_setopt($this->handle, CURLOPT_POSTFIELDS, json_encode($fields));

        //Отключение проверки SSL (разрешение http запросов).
        curl_setopt($this->handle, CURLOPT_SSL_VERIFYPEER, 0);

        //Получение результата выполнения запроса.
        $result = $this->exec();

        //Возвращает полученного ответа в формате ассоциативного массива.
        return json_decode($result, true);
    }

    //Выполнение конкретного http запроса
    private function exec()
    {
        //Получение результата выполнения подготовленного запроса.
        $result = curl_exec($this->handle);

        //Получение информации о выполненном запросе.
        $info = curl_getinfo($this->handle);

        //Проверка статус-кода на выполнение запроса.
        if ($info['http_code'] == 429) {
            throw new Exception('Too many requests');
        } elseif ($info['http_code'] != 200) {
            throw new Exception('Request failed with http code ' . $info['http_code'] . ': ' . $result);
        }
        //Возвращает результат выполнения http запроса.
        return $result;
    }
}