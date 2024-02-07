<?php
function actionTable(){
      insertT();
      deleteT();
      editT();
}     
function insertT(){
    if (!empty ($_GET['items'])) {
        $insert = new insert($_GET['items']);
        if ($insert-> insertTask()){
          echo'<div class="col-md-9 alert alert-success alert-dismissible fade show" role="alert">
          <strong>Goods!</strong> Dungag napod trabahuon.
          <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
        </div>'; 
        }else{
            echo '<div class="col-md-9 alert alert-danger alert-dismissible fade show" role="alert">
      <strong>Agoi!</strong> Dili kadungag usba ba.
      <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    </div>';
        }
    }   
}
function deleteT(){
  if (!empty ($_GET['delete'])) {
      $delete = new delete($_GET['delete']);
      if ($delete-> deleteTask()){
        echo'<div class="col-md-9 alert alert-warning alert-dismissible fade show" role="alert">
        <strong>Ay!</strong> Imu gitang tangan ug usa.
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
      </div>'; 
      }else{
          echo '<div class="col-md-9 alert alert-danger alert-dismissible fade show" role="alert">
    <strong>Agoy!</strong> Dili ma delete.
    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
  </div>';
      }
  }   
}
function editT(){
  if (!empty ($_GET['edit'])) {
      $edit = new edit($_GET['edit']);
      if ($edit-> editTask()){
        echo'<div class="col-md-9 alert alert-info alert-dismissible fade show" role="alert">
        <strong>Goods!</strong> gamay nalang imu buluhaton.
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
      </div>'; 
      }else{
          echo '<div class="col-md-9 alert alert-danger alert-dismissible fade show" role="alert">
    <strong>Agoi!</strong> Wala nadayun usba ba.
    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
  </div>';
      }
  }   
}
function viewTable(){
  $view = new View();
  $view->viewData();
  $view->viewCompletedData();
}
?>