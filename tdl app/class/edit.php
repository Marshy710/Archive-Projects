<?php
class edit extends config{
    public $id;

    public function __construct($id){
    $this->id = $id;
    }
    public function editTask() {
    $con = $this->con();
    $sql = "UPDATE `todolist` SET `status`='COMPLETED', `date_completed`=NOW() WHERE `id` = '$this->id'";
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