<?php 
  require_once 'php/init.php';
?>
<!doctype html>
<html lang="eng">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.rtl.min.css" 
    integrity="sha384-nU14brUcp6StFntEOOEBvcJm4huWjB0OcIeQ3fltAfSmuZFrkAif0T+UtNGlKKQv" crossorigin="anonymous">
    <title>Task Manager</title>
  </head>
  <body>
    <nav class="navbar navbar-dark bg-dark shadow">
        <span class="navbar-brand mb-0 h1">Task Manager</span>
    </nav>
    <div class="container  mt-5">
      <?php actionTable();?>
        <form action="" method="GET">
            <div class="row">
                <div class="col-md-9 form-group mb-2">
                  <input class="form-control" type="text" name="items" placeholder="Enter" required/>
                </div>
                <div class="col-md">
                    <input class="btn btn-success" type="submit" value="Add Task">
                </div>
            </div> 
           </form>
         <?php viewTable();?>  
    </div>

    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.11.8/dist/umd/popper.min.js" 
    integrity="sha384-I7E8VVD/ismYTF4hNIPjVp/Zjvgyol6VFvRkX/vR+Vc4jQkC+hVqc2pM8ODewa9r" crossorigin="anonymous">
    </script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.min.js" 
    integrity="sha384-BBtl+eGJRgqQAUMxJ7pMwbEyER4l1g+O15P+16Ep7Q9Q+zqX6gSbd85u4mG4QzX+" crossorigin="anonymous">
    </script>
    
  </body>
</html>