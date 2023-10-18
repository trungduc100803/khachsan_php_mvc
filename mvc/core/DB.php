<?php
class DB
{
    public $connection;
    public $serverName = "localhost";
    public $userName = "root";
    public $password = "";
    public $dataBaseName = "quan_ly_khach_san";

    public function __construct()
    {
        $this->connection = new mysqli($this->serverName, $this->userName, $this->password, $this->dataBaseName);
        if ($this->connection->connect_error) {
            die("Connection failed: " . $this->connection->connect_error);
        }

        return $this->connection;
    }
}
