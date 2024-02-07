<?php
class insert extends config{
    public $task;

    public function __construct($task){
    $this->task = $task;
    }
    public function insertTask() {
    $con = $this->con();
    $sql = "INSERT INTO `todolist` (`items`) VALUES ('$this->task')";
    $data = $con->prepare($sql);
//Error Handling and Security
    try {
        $data->execute();
        return true;
    } catch (PDOException $e) {
        // Log or handle the error
        return false;
    }
}
}
?>