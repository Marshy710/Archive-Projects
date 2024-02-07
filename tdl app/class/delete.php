<?php
class delete extends config{
    public $id;

    public function __construct($id){
    $this->id = $id;
    }
    public function deleteTask() {
    $con = $this->con();
    $sql = "DELETE FROM `todolist` WHERE `id` = $this->id";
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