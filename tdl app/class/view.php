<?php 
class view extends config{

    public function viewData(){
        $con = $this->con();
        $sql ="SELECT * FROM `todolist` WHERE `status` = 'PENDING';";
        $data = $con->prepare($sql);
        $data->execute();
        $result = $data->fetchAll(PDO::FETCH_ASSOC);
    
        echo "<h3 class='mb-7'>Pending Task</h3>";
    
        if (count($result) > 0) {
            echo "<table class='table table-dark table-striped table-responsive-sm table-hover'>";
            echo "<thead>
                    <tr>       
                        <th>Task Item</th>
                        <th>Action</th>
                    </tr>    
                    </thead><tbody>";
    
            foreach ($result as $data) {
                echo "<tr>";
                echo "<td>$data[items]</td>";
                echo "<td>
                        <a class='btn btn-info btn-sm' href='index.php?edit=$data[id]'> Mark Completed</a>
                        <a class='btn btn-danger btn-sm' href='index.php?delete=$data[id]'> Delete Task</a>
                    </td>";
                echo "</tr>";
            }
    
            echo "</tbody></table>";
        } else {
            echo "<p>No pending tasks available.</p>";
        }
    }
    public function viewCompletedData(){
        $con = $this->con();
        $sql ="SELECT * FROM `todolist` WHERE `status` = 'COMPLETED';";
        $data = $con->prepare($sql);
        $data->execute();
        $result = $data->fetchAll(PDO::FETCH_ASSOC);
        echo "<h3 class='mb-7 mt-5'>Completed Task</h3>";
        echo "<table class='table table-dark table-striped table-sm table-hover'>";
        echo "<thead>
                <tr>       
                    <th>Task Item</th>
                    <th>Date Completed</th>
                    <th>Action</th>
                </tr>    
                </thead><tbody>";
    foreach($result as $data)   {
            echo "<tr>";
            echo "<td>$data[items]</td>";
                echo "<td>$data[date_completed]</td>";
                echo "<td>
                        <a class='btn btn-danger btn-sm' href='index.php?delete=$data[id]'> Delete Task</a>
                    </td>";
            }
        echo"</tbody></table>";        
    } 
}

?>